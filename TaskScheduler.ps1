function Update-OrCreate-ScheduledTask {
    param(
        [Parameter(Mandatory=$true)]
        [psobject]$Policy,
        [string]$PolicyPath
    )

    Write-Log -Message "🔎 Starting Scheduled Task management..." -Level "INFO" -ToRun -ToGlobal -Echo

    try {
        # --- Resolve policy path from Main if not passed ---
        if ([string]::IsNullOrWhiteSpace($PolicyPath) -and $script:global:PolicyPath) { $PolicyPath = $script:global:PolicyPath }
        if ([string]::IsNullOrWhiteSpace($Policy.TaskName) -or [string]::IsNullOrWhiteSpace($PolicyPath)) {
            throw "Policy missing required fields: TaskName or PolicyPath."
        }

        $taskFolderPath = '\DSPObjectListAutoUpdater'
        $taskName       = $Policy.TaskName

        # --- Resolve paths ---
        $policyFolder = Split-Path -Path $PolicyPath -Parent
        $rootFolder   = Split-Path -Path $policyFolder -Parent
        $scriptPath   = Join-Path $rootFolder 'DSPObjectListAutoUpdater.ps1'
        if (-not (Test-Path $scriptPath)) { throw "Main script not found: $scriptPath" }
        $workingDir   = $rootFolder

        # --- Build command (ensures correct working dir) ---
        $quotedScript = '"' + $scriptPath + '"'
        $quotedPolicy = '"' + $PolicyPath + '"'
        $preSetLoc    = "Set-Location -LiteralPath '$workingDir'; & $quotedScript -Policy $quotedPolicy"
        $psExe        = 'powershell.exe'
        $psArgs       = "-NoProfile -ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -Command $preSetLoc"

        # --- Schedule: convert UTC->local for trigger registration ---
        if (-not $Policy.Schedule -or -not $Policy.Schedule.Frequency -or -not $Policy.Schedule.TimeUTC) {
            throw "Schedule block must contain Frequency and TimeUTC."
        }
        $freq      = ($Policy.Schedule.Frequency + '')
        $timeUtc   = [datetime]::ParseExact($Policy.Schedule.TimeUTC, 'HH:mm', $null, 'None')
        $utcToday  = [datetime]::UtcNow.Date.AddHours($timeUtc.Hour).AddMinutes($timeUtc.Minute)
        $localTime = $utcToday.ToLocalTime()
        $atLocal   = Get-Date $localTime -Format 'HH:mm'

        switch ($freq.ToLower()) {
            'daily'   { $desiredKind='Daily';   $newTrigger = New-ScheduledTaskTrigger -Daily   -At $atLocal }
            'weekly'  { $desiredKind='Weekly';  $newTrigger = New-ScheduledTaskTrigger -Weekly  -DaysOfWeek Monday -At $atLocal }
            'monthly' { $desiredKind='Monthly'; $newTrigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At $atLocal }
            default   { throw "Unsupported frequency: $freq" }
        }

        # --- Action/Principal/Settings ---
        try { $action = New-ScheduledTaskAction -Execute $psExe -Argument $psArgs -WorkingDirectory $workingDir }
        catch { $action = New-ScheduledTaskAction -Execute $psExe -Argument $psArgs } # older OS fallback

        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 6)

        # --- Ensure task folder exists (COM path is most reliable) ---
        try {
            $svc = New-Object -ComObject Schedule.Service
            $svc.Connect()
            $root = $svc.GetFolder('\')
            try { $null = $root.GetFolder($taskFolderPath) } catch { $null = $root.CreateFolder($taskFolderPath) }
        } catch { }

        # --- Helper to compare existing trigger to desired ---
        function Get-TriggerInfo($t) {
            $kind = $t | Select-Object -ExpandProperty TriggerType
            $start = [datetime]$t.StartBoundary
            [pscustomobject]@{
                Kind = "$kind"
                Time = $start.ToLocalTime().ToString('HH:mm')
            }
        }

        # --- Detect if task exists (robust) ---
        $existing = Get-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -ErrorAction SilentlyContinue
        $utcInfo  = ("{0} (UTC {1})" -f $atLocal, $Policy.Schedule.TimeUTC)
        $cmdPrev  = "$psExe $psArgs"
        Write-Log -Message "🧭 Task action: $cmdPrev" -Level "INFO" -ToRun -ToGlobal -Echo

        if ($existing) {
            # Compare schedule
            $curTrig = $existing.Triggers | Select-Object -First 1
            $curInfo = Get-TriggerInfo $curTrig
            $sameKind = ($curInfo.Kind -match $desiredKind)  # types like 'Daily'/'Weekly'
            $sameTime = ($curInfo.Time -eq $atLocal)

            if ($sameKind -and $sameTime) {
                Write-Log -Message "✅ Scheduled Task '$taskName' already up-to-date. No changes needed." -Level "SUCCESS" -ToRun -ToGlobal -Echo
                return
            }

            Write-Log -Message "♻️ Updating task '$taskName' → $($desiredKind) at $utcInfo" -Level "INFO" -ToRun -ToGlobal -Echo
            $definition = New-ScheduledTask -Action $action -Trigger $newTrigger -Settings $settings -Principal $principal
            try {
                Unregister-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -Confirm:$false -ErrorAction Stop
                Start-Sleep -Seconds 2
                Register-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -InputObject $definition -ErrorAction Stop | Out-Null
                Write-Log -Message "✅ Scheduled Task updated: $taskName" -Level "SUCCESS" -ToRun -ToGlobal -Echo
            } catch {
                # As a safety, force register (no error log)
                Register-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -InputObject $definition -Force | Out-Null
                Write-Log -Message "✅ Scheduled Task updated (forced): $taskName" -Level "SUCCESS" -ToRun -ToGlobal -Echo
            }
            return
        }

        # --- Create new task ---
        Write-Log -Message "🆕 Creating task '$taskName' → $($desiredKind) at $utcInfo" -Level "INFO" -ToRun -ToGlobal -Echo
        $definitionNew = New-ScheduledTask -Action $action -Trigger $newTrigger -Settings $settings -Principal $principal
        try {
            Register-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -InputObject $definitionNew -ErrorAction Stop | Out-Null
            Write-Log -Message "✅ Created Scheduled Task: $taskName" -Level "SUCCESS" -ToRun -ToGlobal -Echo
        } catch {
            # If it already exists (race), just go through update path quietly
            Register-ScheduledTask -TaskPath $taskFolderPath -TaskName $taskName -InputObject $definitionNew -Force | Out-Null
            Write-Log -Message "✅ Scheduled Task already existed; validated/updated: $taskName" -Level "SUCCESS" -ToRun -ToGlobal -Echo
        }
    }
    catch {
        Write-Log -Message "❌ ERROR in Update-OrCreate-ScheduledTask: $($_.Exception.Message)" -Level "ERROR" -ToRun -ToGlobal -Echo
    }
}
