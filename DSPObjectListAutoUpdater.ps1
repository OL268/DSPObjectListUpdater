param([string]$Policy)  # v1.1-compatible parameter

<#
.SYNOPSIS
  DSP ObjectList AutoUpdater - v1.2
  (Adds Scope.Exclude; post-query flow identical to v1.1)

.DESCRIPTION
  - Loads policy JSON (v1.2 schema with Scope.Exclude).
  - Builds LDAP filter from policy Filters/Logic.
  - Queries AD per scope (Domains/SearchBases) + Exclude.
  - AFTER QUERY (same as v1.1):
        1) Sync-DSPList
        2) Save-ResultSnapshot
        3) Update-OrCreate-ScheduledTask
#>

# --- Imports (same layout as v1.1) ---
. "$PSScriptRoot\Modules\Logging.ps1"
. "$PSScriptRoot\Modules\Utilities.ps1"
. "$PSScriptRoot\Modules\Policy.ps1"
. "$PSScriptRoot\Modules\QueryAD.ps1"
. "$PSScriptRoot\Modules\DSPUpdate.ps1"
. "$PSScriptRoot\Modules\Snapshot.ps1"
. "$PSScriptRoot\Modules\TaskScheduler.ps1"

# Avoid errors if Stop-GlobalLogs is not exported
function Invoke-StopGlobalLogsSafe {
    try {
        $f = Get-Command -Name Stop-GlobalLogs -ErrorAction SilentlyContinue
        if ($null -ne $f) { Stop-GlobalLogs }
    } catch {}
}

# Robust policy path resolution (only used if -Policy missing or invalid)
function Resolve-PolicyPath {
    param([string]$PathIn)

    if (-not [string]::IsNullOrWhiteSpace($PathIn) -and (Test-Path -LiteralPath $PathIn)) {
        return (Resolve-Path -LiteralPath $PathIn).Path
    }

    $candidates = @()
    try {
        $candidates = Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.json -File -ErrorAction SilentlyContinue |
                      Select-Object -ExpandProperty FullName
    } catch {}

    if ($candidates -and $candidates.Count -gt 0) {
        $preferred = @($candidates | Where-Object { $_ -match 'Policy' })
        if ($preferred -and $preferred.Count -gt 0) { return $preferred[0] }
        return $candidates[0]
    }

    throw ("Policy path not provided and no *.json policy found in folder: {0}" -f $PSScriptRoot)
}

function Main {
    param([string]$ConfigPath)

    # Logging session (unchanged)
    $runID        = Get-ShortRunID
    $timestampUTC = Get-TimestampUTC
    Start-GlobalLogs -RunID $runID -TimestampUTC $timestampUTC
    Clean-OldLogs -DaysToKeep 30

    try {
        # Resolve policy
        try {
            $ConfigPath = Resolve-PolicyPath -PathIn $ConfigPath
        } catch {
            Write-Log -Message $_.Exception.Message -Level "ERROR" -ToRun -ToGlobal -Echo
            Invoke-StopGlobalLogsSafe
            return
        }
        $global:PolicyPath = $ConfigPath

        # Load + summarize policy
        $config = Load-Policy -Path $ConfigPath
        Print-PolicySummary -Policy $config

        # Build LDAP filter (same semantics as v1.1)
        $ldap = Build-LDAPFilter -Filters $config.Filters -Logic $config.Logic
        Write-Log -Message ("LDAP Filter built: {0}" -f $ldap) -Level "INFO" -ToRun -ToGlobal

        # Scope parts (only addition is Exclude pass-through)
        $searchBases = $null
        $domains     = $null
        $exclude     = $null
        if ($config.Scope) {
            if ($config.Scope.SearchBases) { $searchBases = $config.Scope.SearchBases }
            if ($config.Scope.Domains)     { $domains     = $config.Scope.Domains }
            if ($config.Scope.Exclude)     { $exclude     = $config.Scope.Exclude }
        }

        # === AD QUERY ===
        $adObjects = Query-ADObjects `
                        -Classes $config.Classes `
                        -LDAPFilter $ldap `
                        -SearchBases $searchBases `
                        -Domains $domains `
                        -Exclude $exclude `
                        -Properties @('distinguishedName','samAccountName','name','objectClass')

        Write-Log -Message ("Query returned {0} objects" -f ($adObjects.Count)) -Level "INFO" -ToRun -ToGlobal -Echo

    

        # 1) Sync DSPlist
        Sync-DSPList -ListName $config.ListName -ADSnapshot $adObjects -Connection $global:DSPConnection

        # 2) Snapshot after sync
        $null = Save-ResultSnapshot -ListName $config.ListName -Classes $config.Classes -RunID $runID -TimestampUTC $timestampUTC -FullObjects $adObjects

        # 3) Schedule task
        Update-OrCreate-ScheduledTask -Policy $config -PolicyPath $global:PolicyPath

        Write-Log -Message("✅ Finished DSP Object List AutoUpdater Run $runID ." -f  $runID) -Level "SUCCESS" -ToRun -ToGlobal -Echo
    }
    catch {
        Write-Log -Message ("❌ FATAL: {0}" -f $_.Exception.Message) -Level "ERROR" -ToRun -ToGlobal -Echo
        throw
    }
    finally {
        Invoke-StopGlobalLogsSafe
    }
}

Main -ConfigPath $Policy
