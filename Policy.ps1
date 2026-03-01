<#
.SYNOPSIS
  Policy load/validate helpers (v1.2)

.DESCRIPTION
  Loads policy JSON, validates v1.2 schema (including Scope.Exclude),
  and prints a concise summary via the existing logging pipeline.

.AUTHOR
  DSP PM Team

.VERSION
  1.2
#>

function Load-Policy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    } catch {
        throw "Unable to read policy file '$Path': $($_.Exception.Message)"
    }

    try {
        # PS 5.1: ConvertFrom-Json has no -Depth switch
        $policy = $raw | ConvertFrom-Json
    } catch {
        throw "Invalid JSON in policy file '$Path': $($_.Exception.Message)"
    }

    # Strict v1.2 schema validation (no legacy shims)
    Validate-ConfigStructure -Config $policy
    return $policy
}

function New-PolicyFile {
    param (
        [hashtable]$Config,
        [string]$RunID,
        [string]$TimestampUTC
    )

    # Save as-is (v1.2 schema)
    $json = ($Config | ConvertTo-Json -Depth 10)

    # PS 5.1-safe null fallback (no '??')
    $taskName = if ($null -ne $Config -and $Config.ContainsKey('TaskName') -and -not [string]::IsNullOrWhiteSpace($Config.TaskName)) {
        [string]$Config.TaskName
    } else {
        'Policy'
    }

    $name = ('{0}_{1}.json' -f $taskName, $TimestampUTC)
    $out  = Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath $name
    Set-Content -LiteralPath $out -Value $json -Encoding UTF8
    return $out
}

function Print-PolicySummary {
    param([psobject]$Policy)

    # Use central logging mechanism; do not use Write-Host
    Write-Log -Message "Policy Summary:" -Level "INFO" -ToRun -ToGlobal -Echo
    Write-Log -Message (" TaskName   : {0}" -f $Policy.TaskName) -Level "INFO" -ToRun -ToGlobal -Echo
    Write-Log -Message (" ListName   : {0}" -f $Policy.ListName) -Level "INFO" -ToRun -ToGlobal -Echo

    if ($Policy.Classes) {
        Write-Log -Message (" Classes    : {0}" -f ($Policy.Classes -join ', ')) -Level "INFO" -ToRun -ToGlobal -Echo
    }

    if ($Policy.Filters) {
        foreach ($f in $Policy.Filters) {
            if ($null -ne $f) {
                Write-Log -Message (" Filter     : {0} {1} {2}" -f $f.Attribute, $f.Operator, $f.Value) -Level "INFO" -ToRun -ToGlobal -Echo
            }
        }
    }

    if ($Policy.Scope) {
        if ($Policy.Scope.SearchBases) {
            foreach ($b in $Policy.Scope.SearchBases) {
                if ($null -ne $b) {
                    Write-Log -Message (" Scope      : {0} :: {1}" -f $b.Domain, $b.BaseDN) -Level "INFO" -ToRun -ToGlobal -Echo
                }
            }
        }
        if ($Policy.Scope.Exclude) {
            foreach ($e in $Policy.Scope.Exclude) {
                if ($null -ne $e) {
                    Write-Log -Message (" Exclude    : {0} :: {1}" -f $e.Domain, $e.BaseDN) -Level "INFO" -ToRun -ToGlobal -Echo
                }
            }
        }
        if ($Policy.Scope.Domains) {
            Write-Log -Message (" Domains    : {0}" -f ($Policy.Scope.Domains -join ', ')) -Level "INFO" -ToRun -ToGlobal -Echo
        }
    }

    if ($Policy.Schedule) {
        Write-Log -Message (" Schedule   : {0} at {1} UTC" -f $Policy.Schedule.Frequency, $Policy.Schedule.TimeUTC) -Level "INFO" -ToRun -ToGlobal -Echo
    }
}
