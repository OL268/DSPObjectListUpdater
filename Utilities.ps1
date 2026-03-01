<#
.SYNOPSIS
  Utilities & validators (v1.2)

.DESCRIPTION
  Strict schema validation for v1.2 including Scope.Exclude.
  Also exposes helpers used by the runner and AD query.

.AUTHOR
  DSP PM Team

.VERSION
  1.2
#>

# ---------- General helpers (restored) ----------

function Get-TimestampUTC {
    # Filename-safe UTC timestamp (no colons)
    # Example: 2025-08-19T225514Z
    (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHHmmssZ")
}

function Get-ShortRunID {
    # Compact run ID for filenames/logs
    $now = Get-Date
    "{0:yyyyMMdd}T{0:HHmmss}_{1}" -f $now, ($now.Millisecond.ToString("0000"))
}

# ---------- Validators ----------

function Validate-Classes {
    param([object[]]$Classes)
    if (-not $Classes -or $Classes.Count -eq 0) {
        throw "Policy.Classes must contain at least one class."
    }
}
function Validate-Filters {
    param([object[]]$Filters, [string]$Logic)

    if (-not $Filters -or $Filters.Count -eq 0) {
        throw "Policy.Filters must contain at least one filter."
    }

    foreach ($f in $Filters) {
        if (-not $f) { throw "Filter entry is null." }

        $attr = [string]$f.Attribute
        $op   = ([string]$f.Operator).ToLower()

        if ([string]::IsNullOrWhiteSpace($attr)) { throw "Filter missing Attribute." }
        if ([string]::IsNullOrWhiteSpace($op))   { throw "Filter missing Operator." }

        switch ($op) {
            '-present' {
                # OK with empty or missing Value
                continue
            }
            '-in' { 
                if ($null -eq $f.Value) { throw "Filter '$attr' with '-in' requires Value (array or string)." }
                # arrays are fine; a single string is also accepted (treated as 1-item list)
                if (($f.Value -is [System.Collections.IEnumerable]) -and -not ($f.Value -is [string])) {
                    if ((@($f.Value) | Measure-Object).Count -eq 0) { throw "Filter '$attr' with '-in' requires at least one value." }
                } else {
                    if ([string]::IsNullOrWhiteSpace([string]$f.Value)) { throw "Filter '$attr' with '-in' requires a non-empty value." }
                }
            }
            '-notin' {
                if ($null -eq $f.Value) { throw "Filter '$attr' with '-notin' requires Value (array or string)." }
                if (($f.Value -is [System.Collections.IEnumerable]) -and -not ($f.Value -is [string])) {
                    if ((@($f.Value) | Measure-Object).Count -eq 0) { throw "Filter '$attr' with '-notin' requires at least one value." }
                } else {
                    if ([string]::IsNullOrWhiteSpace([string]$f.Value)) { throw "Filter '$attr' with '-notin' requires a non-empty value." }
                }
            }
            default {
                # All other operators require a non-empty string value
                if ($null -eq $f.Value -or [string]::IsNullOrWhiteSpace([string]$f.Value)) {
                    throw "Filter '$attr' with operator '$op' requires a non-empty Value."
                }
            }
        }
    }

    if ($Logic -and ($Logic -ne 'AND' -and $Logic -ne 'OR')) {
        throw "Policy.Logic must be 'AND' or 'OR'."
    }
}


function Validate-Scope {
    param([psobject]$Scope)
    if (-not $Scope) { return } # Scope is optional

    if ($Scope.Domains -and $Scope.SearchBases) {
        throw "Scope cannot specify both 'Domains' and 'SearchBases'. Choose one."
    }
    if (-not $Scope.Domains -and -not $Scope.SearchBases) {
        throw "Scope must specify either 'Domains' or 'SearchBases'."
    }

    if ($Scope.SearchBases) {
        foreach ($b in $Scope.SearchBases) {
            if (-not $b.Domain -or -not $b.BaseDN) {
                throw "Each SearchBases item must include 'Domain' and 'BaseDN'."
            }
            if ($b.BaseDN -notmatch '(^[A-Za-z]+=.+,DC=.+)') {
                throw "Invalid BaseDN in SearchBases: $($b.BaseDN)"
            }
        }
    }

    if ($Scope.Exclude) {
        foreach ($e in $Scope.Exclude) {
            if (-not $e.Domain -or -not $e.BaseDN) {
                throw "Each Exclude item must include 'Domain' and 'BaseDN'."
            }
            if ($e.BaseDN -notmatch '(^[A-Za-z]+=.+,DC=.+)') {
                throw "Invalid BaseDN in Exclude: $($e.BaseDN)"
            }
        }
    }

    if ($Scope.Domains) {
        foreach ($d in $Scope.Domains) {
            if ([string]::IsNullOrWhiteSpace($d)) {
                throw "Empty domain name in 'Domains'."
            }
        }
    }
}

function Validate-Schedule {
    param([psobject]$Schedule)
    if ($Schedule) {
        if (-not $Schedule.Frequency -or -not $Schedule.TimeUTC) {
            throw "Schedule requires Frequency and TimeUTC."
        }
    }
}

function Validate-ConfigStructure {
    param([psobject]$Config)

    if (-not $Config.ListName) { throw "Policy.ListName is required." }
    Validate-Classes  -Classes  $Config.Classes
    Validate-Filters  -Filters  $Config.Filters -Logic $Config.Logic
    Validate-Scope    -Scope    $Config.Scope
    Validate-Schedule -Schedule $Config.Schedule
}


function Join-ClassNames {
    <#
    .SYNOPSIS
      Build a stable code from class names (used by Snapshot.ps1).

    .PARAMETER Classes
      Array of class names (e.g., 'Users','Groups','Computers').

    .OUTPUTS
      String like "Computers_Groups_Users". Returns "All" if empty.
    #>
    param([object[]]$Classes)

    if (-not $Classes -or $Classes.Count -eq 0) { return "All" }

    $clean = @()
    foreach ($c in $Classes) {
        if ($null -eq $c) { continue }
        $s = [string]$c
        if (-not [string]::IsNullOrWhiteSpace($s)) { $clean += $s }
    }

    if ($clean.Count -eq 0) { return "All" }

    # Stable, deduped, filesystem-safe-ish token
    $tokens = $clean | Sort-Object -Unique | ForEach-Object {
        ($_ -replace '[^\w\-\.]', '_')
    }

    return ($tokens -join '_')
}


# ---------- Scope/DN helper ----------

function Test-DistinguishedNameInScope {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistinguishedName,
        [object[]]$IncludeBases,
        [object[]]$ExcludeBases
    )

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
        return $false
    }

    $dn = $DistinguishedName.ToLowerInvariant()
    $included = $false

    # Include logic: if SearchBases are provided, DN must be under at least one of them.
    if ($IncludeBases -and $IncludeBases.Count -gt 0) {
        foreach ($b in $IncludeBases) {
            if ($null -ne $b -and $null -ne $b.BaseDN) {
                $base = $b.BaseDN.ToLowerInvariant()
                if ($dn -like ("*" + $base)) { $included = $true; break }
            }
        }
    } else {
        # Domain-scope: include by default (will be filtered by Exclude if present)
        $included = $true
    }

    if (-not $included) { return $false }

    # Exclude logic: if DN matches any Exclude.BaseDN, drop it.
    if ($ExcludeBases -and $ExcludeBases.Count -gt 0) {
        foreach ($e in $ExcludeBases) {
            if ($null -ne $e -and $null -ne $e.BaseDN) {
                $exc = $e.BaseDN.ToLowerInvariant()
                if ($dn -like ("*" + $exc)) { return $false }
            }
        }
    }

    return $true
}

# ---------- LDAP filter builder ----------

function Escape-LdapValue {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '(' { [void]$sb.Append('\28') }
            ')' { [void]$sb.Append('\29') }
            '*' { [void]$sb.Append('\2a') }
            '\' { [void]$sb.Append('\5c') }
            ([char]0) { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    $sb.ToString()
}

function Escape-LdapValueForLike {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '*' { [void]$sb.Append('*') }          # keep wildcards for -like
            '(' { [void]$sb.Append('\28') }
            ')' { [void]$sb.Append('\29') }
            '\' { [void]$sb.Append('\5c') }
            ([char]0) { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    $sb.ToString()
}

function Build-LDAPFilter {
    <#
    .SYNOPSIS
      Build an LDAP filter string from policy Filters[] and Logic.

    .PARAMETER Filters
      Array of @{ Attribute='<attr>'; Operator='-like|-eq|-ne|-notlike|-startswith|-endswith|-in|-notin|-present'; Value='<val or array>' }

    .PARAMETER Logic
      'AND' (default) or 'OR' to combine conditions.

    .OUTPUTS
      [string] LDAP filter, e.g. (&(attr=value)(cn=*admin*))
    #>
    param(
        [object[]]$Filters,
        [string]$Logic = 'AND'
    )

    if (-not $Filters -or $Filters.Count -eq 0) {
        return "(objectClass=*)"
    }

    $clauses = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $Filters.Count; $i++) {
        $f = $Filters[$i]
        if ($null -eq $f) { continue }

        $attr = [string]$f.Attribute
        $op   = [string]$f.Operator
        $val  = $f.Value

        if ([string]::IsNullOrWhiteSpace($attr)) { continue }
        if ([string]::IsNullOrWhiteSpace($op))   { $op = '-eq' }

        switch ($op.ToLower()) {
            '-eq' {
                $v = Escape-LdapValue ([string]$val)
                $clauses.Add("($attr=$v)")
            }
            '-ne' {
                $v = Escape-LdapValue ([string]$val)
                $clauses.Add("(!($attr=$v))")
            }
            '-like' {
                $v = Escape-LdapValueForLike ([string]$val)
                $clauses.Add("($attr=$v)")
            }
            '-notlike' {
                $v = Escape-LdapValueForLike ([string]$val)
                $clauses.Add("(!($attr=$v))")
            }
            '-startswith' {
                $v = Escape-LdapValue ([string]$val)
                $clauses.Add("($attr=${v}*)")
            }
            '-endswith' {
                $v = Escape-LdapValue ([string]$val)
                $clauses.Add("($attr=*${v})")
            }
            '-present' {
                # Value ignored: attribute existence
                $clauses.Add("($attr=*)")
            }
            '-in' {
                # Value should be array-like
                $vals = @()
                if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) { $vals = $val } else { $vals = @($val) }
                $inner = New-Object System.Collections.Generic.List[string]
                foreach ($vv in $vals) {
                    $inner.Add("($attr=$(Escape-LdapValue ([string]$vv)))")
                }
                if ($inner.Count -gt 1) {
                    $clauses.Add("(|{0})" -f ($inner -join ""))
                } elseif ($inner.Count -eq 1) {
                    $clauses.Add($inner[0])
                }
            }
            '-notin' {
                $vals = @()
                if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) { $vals = $val } else { $vals = @($val) }
                $inner = New-Object System.Collections.Generic.List[string]
                foreach ($vv in $vals) {
                    $inner.Add("($attr=$(Escape-LdapValue ([string]$vv)))")
                }
                if ($inner.Count -gt 1) {
                    $clauses.Add("(!(|{0}))" -f ($inner -join ""))  # not (val1 OR val2)
                } elseif ($inner.Count -eq 1) {
                    $clauses.Add("!{0}" -f $inner[0])
                }
            }
            default {
                # Fallback to equality
                $v = Escape-LdapValue ([string]$val)
                $clauses.Add("($attr=$v)")
            }
        }

        # Log each clause build (INFO, as logger has no DEBUG)
        try { Write-Log -Message ("Build-LDAPFilter: {0} {1} {2}" -f $attr,$op,$val) -Level "INFO" -ToRun -ToGlobal } catch {}
    }

    if ($clauses.Count -eq 0) { return "(objectClass=*)" }
    if ($clauses.Count -eq 1) { return $clauses[0] }

    if ($Logic -and $Logic.ToUpper() -eq 'OR') {
        return "(|{0})" -f ($clauses -join "")
    } else {
        return "(&{0})" -f ($clauses -join "")
    }
}
