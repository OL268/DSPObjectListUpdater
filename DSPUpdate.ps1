# ============================================
# DSPUpdate.ps1 - Sync DSP ObjectList (v1.1+ contention-safe, PS 5.1)
# ============================================

function Sync-DSPList {
    param (
        [string]$ListName,
        [array]$ADSnapshot,    # rows contain: ObjectGUID, Name, sAMAccountName, distinguishedName
        [object]$Connection
    )

    # --- constants ---
    $MAX_LIST_SIZE = 10000

    # --- helpers ---
    function NowStamp { (Get-Date -Format o) }

    function New-TempCsvFromGuids {
        param([string[]]$Guids)
        if (-not $Guids -or $Guids.Count -eq 0) { return $null }
        $tmp = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), ".csv")
        # one GUID per line, no header
        [System.IO.File]::WriteAllLines($tmp, $Guids)
        return $tmp
    }

    # Fast unique GUIDs (PS 5.1)
    function Get-UniqueGuids {
        param([array]$Rows)
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($r in $Rows) {
            if ($null -ne $r -and $r.PSObject.Properties['ObjectGUID'] -and $r.ObjectGUID) {
                [void]$set.Add([string]$r.ObjectGUID)
            }
        }
        $arr = New-Object 'System.Collections.Generic.List[string]'
        foreach ($g in $set) { [void]$arr.Add($g) }
        return $arr.ToArray()
    }

    # Set difference using HashSet
# Set difference using HashSet (BUGFIX: null-safe when B is $null)
function Get-Minus {
    param([string[]]$A, [string[]]$B)  # return A \ B

    if ($null -eq $A) { $A = @() }
    if ($null -eq $B) { $B = @() }

    # Create set with comparer only; then union to avoid null ctor arg
    $bSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($B.Count -gt 0) { $bSet.UnionWith([string[]]$B) }

    $out = New-Object 'System.Collections.Generic.List[string]'
    foreach ($x in $A) {
        if ($null -ne $x -and -not $bSet.Contains($x)) { [void]$out.Add($x) }
    }
    return $out.ToArray()
}


    # CSV chunked pipeline runner with retry/backoff for "Concurrent update blocked"
    function Invoke-DspCsvPipeline {
        param(
            [ValidateSet('Add','Remove')] [string]$Mode,
            [string]$ListName,
            [object]$Connection,
            [string[]]$Guids,
            [int]$InitialChunkSize = 1000,
            [int]$MinChunkSize = 100,
            [int]$MaxAttempts = 5
        )

        if (-not $Guids -or $Guids.Count -eq 0) { return }

        $chunkSize = $InitialChunkSize
        $index = 0
        $total = $Guids.Count

        while ($index -lt $total) {
            # Determine this chunk slice
            $count = [Math]::Min($chunkSize, $total - $index)
            $slice = $Guids[$index..($index + $count - 1)]

            $currentChunk = [int]([Math]::Floor($index / [double]$chunkSize)) + 1
            $totalChunks  = [int]([Math]::Ceiling($total / [double]$chunkSize))
            $verb = ('➕ Adding','➖ Removing')[$Mode -ne 'Add']

            # PS-safe formatting (avoid [string]::Format index errors)
            Write-Host ("{0} ({1}) via CSV pipeline chunk {2}/{3}..." -f $verb, $slice.Count, $currentChunk, $totalChunks)

            $csv = $null
            try {
                $csv = New-TempCsvFromGuids -Guids $slice

                $attempt = 0
                $completedThisSlice = $false
                while (-not $completedThisSlice) {
                    try {
                        if ($Mode -eq 'Add') {
                            Import-Csv -Path $csv -Header 'ObjectGuid' |
                                Add-DSPObjectListMember -ListName $ListName -DirectoryType AD -Connection $Connection -WarningAction Continue -ErrorAction Stop |
                                Out-Null
                        } else {
                            Import-Csv -Path $csv -Header 'ObjectGuid' |
                                Remove-DSPObjectListMember -ListName $ListName -DirectoryType AD -Connection $Connection -WarningAction Continue -ErrorAction Stop |
                                Out-Null
                        }
                        $completedThisSlice = $true
                    } catch {
                        $msg = $_.Exception.Message
                        if ($msg -match 'Concurrent update blocked') {
                            if ($attempt -lt $MaxAttempts) {
                                $attempt++
                                Start-Sleep -Seconds ([int]([math]::Pow(2, $attempt)))  # 2,4,8,16,32 sec
                                continue
                            }
                            if ($chunkSize -gt $MinChunkSize) {
                                $newSize = [int]([Math]::Max($MinChunkSize, [Math]::Floor($chunkSize / 2)))
                                Write-Warning ("Contention persists — reducing chunk size from {0} to {1} and retrying." -f $chunkSize, $newSize)
                                $chunkSize = $newSize
                                # break; outer while will re-slice with new chunk size (don't advance $index)
                                $completedThisSlice = $true
                            } else {
                                throw
                            }
                        } else {
                            throw
                        }
                    }
                }

            } finally {
                if ($csv) { Remove-Item -Path $csv -Force -ErrorAction SilentlyContinue }
            }

            # Advance index only if the slice actually finished with current chunk size
            if ($count -eq [Math]::Min($chunkSize, $total - $index)) {
                $index += $count
            }
        }
    }

    # --- connect & ensure list exists ---
    Write-Host "$(NowStamp) [INFO] 🔎 Checking if DSP List '$ListName' exists..."
    Write-Host "$(NowStamp) [INFO] 🔗 Connecting to DSP Server 'localhost'..."
    try {
        $Connection = Connect-DSPServer -ComputerName "localhost"
        Write-Host "$(NowStamp) [SUCCESS] ✅ Connected."
    } catch {
        Write-Log -Message ("❌ Failed to connect to DSP Server: " + $_.Exception.Message) -Level "ERROR" -ToRun -ToGlobal -Echo
        return
    }

    $existingList = Get-DSPObjectList -Name $ListName -Connection $Connection
    if (-not $existingList -or -not $existingList.Id) {
        Write-Host "$(NowStamp) [WARN] ⚠️ List '$ListName' does not exist. Creating..."
        try {
            $null = New-DSPObjectList -Name $ListName -Monitored $true -MonitoringSecurityNotification $true -Connection $Connection
            Write-Host "$(NowStamp) [SUCCESS] ✅ List '$ListName' created."
        } catch {
            Write-Log -Message "❌ Failed to create list '$ListName': $($_.Exception.Message)" -Level "ERROR" -ToRun -ToGlobal -Echo
            return
        }
    } else {
        Write-Host "$(NowStamp) [SUCCESS] ✅ List '$ListName' exists."
    }

    # --- get current members ---
    Write-Host "$(NowStamp) [INFO] 🔎 Retrieving current list members..."
    $currentMembers = Get-DSPObjectListMember -ListName $ListName -Connection $Connection -ErrorAction SilentlyContinue
    if ($null -eq $currentMembers) { $currentMembers = @() }
    $currentCount = ($currentMembers | Measure-Object).Count
    Write-Host "$(NowStamp) [INFO] 🔎 Current members: $currentCount"

    $currentGuids = @()
    if ($currentCount -gt 0) {
        $currentGuids = $currentMembers | Where-Object { $_.ObjectGuid } | Select-Object -ExpandProperty ObjectGuid -Unique
    }

    # --- AD snapshot GUIDs (fast unique) ---
    Write-Host "$(NowStamp) [INFO] 🔎 Preparing AD snapshot GUIDs..."
    $adGuids = Get-UniqueGuids -Rows $ADSnapshot
    Write-Host "$(NowStamp) [INFO] 📊 AD snapshot GUIDs: $($adGuids.Count)"

    # --- build plan (remove first, then add; respect 10K) ---
    Write-Host "$(NowStamp) [INFO] 🔁 Computing diff (adds/removes)..."
    $toRemove = Get-Minus -A $currentGuids -B $adGuids       # present in list but NOT in AD snapshot
    $toAdd    = Get-Minus -A $adGuids     -B $currentGuids   # present in AD snapshot but NOT in list

    # capacity after removals
    $expectedAfterRemovals = [Math]::Max(0, $currentCount - $toRemove.Count)
    $allowedAdds = $MAX_LIST_SIZE - $expectedAfterRemovals
    if ($allowedAdds -lt 0) { $allowedAdds = 0 }

    $leftovers = @()
    if ($toAdd.Count -gt $allowedAdds) {
        $leftovers = $toAdd | Select-Object -Skip $allowedAdds
        $toAdd     = $toAdd | Select-Object -First $allowedAdds
        Write-Host "$(NowStamp) [WARN] ⚠️ Trimming $($leftovers.Count) additions to respect 10K cap."
    }

    Write-Host "$(NowStamp) [INFO] 🛠️ Sync Plan: Add=$($toAdd.Count), Remove=$($toRemove.Count)"

    # --- execute removals (CSV, adaptive, contention-safe) ---
    if ($toRemove.Count -gt 0) {
        Write-Host "$(NowStamp) [INFO] ➖ Removing ($($toRemove.Count)) via CSV pipeline..."
        try {
            Invoke-DspCsvPipeline -Mode Remove -ListName $ListName -Connection $Connection -Guids $toRemove -InitialChunkSize 1000 -MinChunkSize 100 -MaxAttempts 5
            Write-Host "$(NowStamp) [SUCCESS] ➖ Remove pipeline completed."
        } catch {
            Write-Log -Message ("❌ Remove pipeline error: " + $_.Exception.Message) -Level "ERROR" -ToRun -ToGlobal -Echo
        }
    }

    # --- execute additions (CSV, adaptive, fast path) ---
    if ($toAdd.Count -gt 0) {
        Write-Host "$(NowStamp) [INFO] ➕ Adding ($($toAdd.Count)) via CSV pipeline..."
        try {
            Invoke-DspCsvPipeline -Mode Add -ListName $ListName -Connection $Connection -Guids $toAdd -InitialChunkSize 1000 -MinChunkSize 500 -MaxAttempts 5
            Write-Host "$(NowStamp) [SUCCESS] ➕ Add pipeline completed."
        } catch {
            Write-Log -Message ("❌ Add pipeline error: " + $_.Exception.Message) -Level "ERROR" -ToRun -ToGlobal -Echo
        }
    }

    # --- log leftovers (over 10K) ---
    if ($leftovers.Count -gt 0) {
        $first50 = ($leftovers | Select-Object -First 50) -join ', '
        Write-Log -Message ("⚠️ Leftover (not added due to 10K): {0}" -f $first50) -Level "WARN" -ToRun -ToGlobal -Echo
        Write-Host "$(NowStamp) [WARN] ⚠️ Skipped $($leftovers.Count) addition(s) due to 10K cap (first 50 in log)."
    }

    # --- verify final count ---
    $finalMembers = Get-DSPObjectListMember -ListName $ListName -Connection $Connection -ErrorAction SilentlyContinue
    $finalCount = ($finalMembers | Measure-Object).Count
    Write-Host "$(NowStamp) [✅ SUCCESS] ✅ DSP List '$ListName' synchronization finished. Final count=$finalCount"
}
