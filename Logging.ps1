# ============================================
# Logging.ps1 - Professional Logging Module (FINAL)
# ============================================

function Start-GlobalLogs {
    param (
        [string]$RunID,
        [string]$TimestampUTC
    )

    $global:LogRoot = ".\Logs"
    if (-not (Test-Path $LogRoot)) { 
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
        Write-Host "📁 Created Logs folder: $LogRoot" -ForegroundColor Green
    }

    $global:RunLogPath = Join-Path $LogRoot ("Run_${RunID}_${TimestampUTC}.log")
    $global:GlobalLogPath = Join-Path $LogRoot "GlobalExecutionLog.txt"

    if (-not (Test-Path $GlobalLogPath)) {
        New-Item -Path $GlobalLogPath -ItemType File | Out-Null
        Write-Host "📄 Created global log file: $GlobalLogPath" -ForegroundColor Green
    }

    $global:RunID = $RunID
    $global:RunTimestampUTC = $TimestampUTC

    Write-Log -Message "🔵 Logging initialized. Session log: $global:RunLogPath" -Level "INFO" -ToRun -ToGlobal -Echo
}

function Write-Log {
    param (
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',
        [switch]$ToRun,
        [switch]$ToGlobal,
        [switch]$Echo
    )

    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    switch ($Level) {
        "INFO"    { $prefix = "[ℹ️ INFO]" }
        "WARN"    { $prefix = "[⚠️ WARN]" }
        "ERROR"   { $prefix = "[❌ ERROR]" }
        "SUCCESS" { $prefix = "[✅ SUCCESS]" }
    }

    $consoleLine = "$timestamp $prefix $Message"
    $logLine     = "$timestamp $prefix [RUN:$global:RunID] $Message"

    # Echo to console if requested
    if ($Echo) {
        switch ($Level) {
            "INFO"    { Write-Host $consoleLine -ForegroundColor Cyan }
            "WARN"    { Write-Host $consoleLine -ForegroundColor Yellow }
            "ERROR"   { Write-Host $consoleLine -ForegroundColor Red }
            "SUCCESS" { Write-Host $consoleLine -ForegroundColor Green }
        }
    }

    # Write to Session Run Log if needed
    if ($global:RunLogPath -and $ToRun) {
        try {
            Add-Content -Path $global:RunLogPath -Value $logLine
        } catch {
            Write-Host "❌ ERROR writing to Run log: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Write to Global Execution Log if needed
    if ($global:GlobalLogPath -and $ToGlobal) {
        try {
            Add-Content -Path $global:GlobalLogPath -Value $logLine
        } catch {
            Write-Host "❌ ERROR writing to Global log: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Clean-OldLogs {
    param (
        [int]$DaysToKeep = 30
    )

    $cutoff = (Get-Date).AddDays(-$DaysToKeep)

    # --- Clean Logs ---
    $logFiles = Get-ChildItem -Path ".\Logs" -Filter "Run_*.log" -ErrorAction SilentlyContinue
    foreach ($file in $logFiles) {
        if ($file.LastWriteTime -lt $cutoff) {
            try {
                Remove-Item $file.FullName -Force
                Write-Log -Message "🧹 Deleted old log file: $($file.Name)" -Level "INFO" -ToGlobal
            } catch {
                Write-Host "⚠️ Failed to delete old log: $($file.Name)" -ForegroundColor Yellow
            }
        }
    }

    # --- Clean Snapshots ---
    $snapshotFiles = Get-ChildItem -Path ".\Snapshots" -Filter "*.json" -ErrorAction SilentlyContinue
    foreach ($file in $snapshotFiles) {
        if ($file.LastWriteTime -lt $cutoff) {
            try {
                Remove-Item $file.FullName -Force
                Write-Log -Message "🧹 Deleted old snapshot file: $($file.Name)" -Level "INFO" -ToGlobal
            } catch {
                Write-Host "⚠️ Failed to delete old snapshot: $($file.Name)" -ForegroundColor Yellow
            }
        }
    }

    # --- Clean Results ---
    $resultFiles = Get-ChildItem -Path ".\Results" -Filter "*.json" -ErrorAction SilentlyContinue
    foreach ($file in $resultFiles) {
        if ($file.LastWriteTime -lt $cutoff) {
            try {
                Remove-Item $file.FullName -Force
                Write-Log -Message "🧹 Deleted old result file: $($file.Name)" -Level "INFO" -ToGlobal
            } catch {
                Write-Host "⚠️ Failed to delete old result: $($file.Name)" -ForegroundColor Yellow
            }
        }
    }
}


function Write-RunSummary {
    param (
        [int]$QueriedObjects = 0,
        [int]$ObjectsAdded = 0,
        [int]$ObjectsRemoved = 0,
        [datetime]$StartTime
    )

    $endTime = Get-Date
    $duration = New-TimeSpan -Start $StartTime -End $endTime

    $summary = @"
📋 Run Summary:
- Queried Objects: $QueriedObjects
- Objects Added to DSP: $ObjectsAdded
- Objects Removed from DSP: $ObjectsRemoved
- Total Duration: {0:D2}:{1:D2}:{2:D2}
"@ -f $duration.Hours, $duration.Minutes, $duration.Seconds

    Write-Log -Message $summary -Level "INFO" -ToRun -ToGlobal -Echo
}
