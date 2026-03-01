# ============================================
# Snapshot.ps1 - Snapshot and Results Saving Module
# ============================================

function Save-ResultSnapshot {
    param (
        [string]$ListName,
        [string[]]$Classes,
        [string]$RunID,
        [string]$TimestampUTC,
        [object[]]$FullObjects
    )

    # Ensure folders
    if (-not (Test-Path ".\Snapshots")) { New-Item ".\Snapshots" -ItemType Directory | Out-Null }
    if (-not (Test-Path ".\Results")) { New-Item ".\Results" -ItemType Directory | Out-Null }

    # Build filenames
    $classCode = Join-ClassNames -Classes $Classes
    $snapshotFile = ".\Snapshots\Snapshot_${ListName}_${classCode}_${RunID}_${TimestampUTC}.json"
    $resultsFile = ".\Results\FullADResults_${classCode}_${RunID}_${TimestampUTC}.json"

    # Save Full Result
    try {
        $FullObjects | ConvertTo-Json -Depth 5 | Set-Content -Path $resultsFile -Encoding UTF8
        Write-Log -Message "💾 Full AD query results saved: $resultsFile" -Level "SUCCESS" -ToRun -ToGlobal -Echo
    } catch {
        Write-Log -Message "❌ Failed to save full results: $($_.Exception.Message)" -Level "ERROR" -ToRun -ToGlobal -Echo
    }

    # Save Snapshot (only critical fields)
    $snapshot = $FullObjects | Select-Object objectGUID, distinguishedName

    try {
        $snapshot | ConvertTo-Json -Depth 5 | Set-Content -Path $snapshotFile -Encoding UTF8
        Write-Log -Message "💾 Snapshot saved: $snapshotFile" -Level "SUCCESS" -ToRun -ToGlobal -Echo
    } catch {
        Write-Log -Message "❌ Failed to save snapshot: $($_.Exception.Message)" -Level "ERROR" -ToRun -ToGlobal -Echo
    }

    return $snapshotFile
}
