<#
.SYNOPSIS
  AD query helpers (v1.2)

.DESCRIPTION
  Queries Active Directory for objects based on policy inputs.
  Adds support for Scope.Exclude (array of { Domain, BaseDN }).
  Uses existing logging pipeline (Write-Log). No Write-Host.

.AUTHOR
  DSP PM Team

.VERSION
  1.2
#>

function Convert-DomainToBaseDN {
    param([Parameter(Mandatory = $true)][string]$DomainFqdn)
    $parts = $DomainFqdn.Split('.') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (-not $parts -or $parts.Count -eq 0) { return $null }
    return ($parts | ForEach-Object { 'DC=' + $_ }) -join ','
}

function Get-DomainRootDN {
    param([Parameter(Mandatory = $true)][string]$ServerFqdn)
    try {
        $dom = Get-ADDomain -Server $ServerFqdn -ErrorAction Stop
        return $dom.DistinguishedName
    } catch {
        return Convert-DomainToBaseDN -DomainFqdn $ServerFqdn
    }
}

function Query-ADObjects {
    <#
    .SYNOPSIS
      Query AD using LDAP filter and policy scope with Exclude support.
    #>
    param(
        [string[]]$Classes,
        [Parameter(Mandatory = $true)][string]$LDAPFilter,
        [object[]]$SearchBases,
        [string[]]$Domains,
        [object[]]$Exclude,
        [object[]]$Properties = @('distinguishedName'),
        [int]$PageSize = 200   # kept for compatibility; mapped to -ResultPageSize
    )

    $all = New-Object System.Collections.Generic.List[object]

    $includeBases = @()
    if ($SearchBases -and $SearchBases.Count -gt 0) {
        foreach ($b in $SearchBases) {
            if ($null -ne $b -and $b.Domain -and $b.BaseDN) { $includeBases += $b }
        }
    }
    $excludeBases = @()
    if ($Exclude -and $Exclude.Count -gt 0) {
        foreach ($e in $Exclude) {
            if ($null -ne $e -and $e.Domain -and $e.BaseDN) { $excludeBases += $e }
        }
    }

    if ($Classes -and $Classes.Count -gt 0) {
        Write-Log -Message ("Query-ADObjects: Classes={0}" -f ($Classes -join ',')) -Level "INFO" -ToRun -ToGlobal -Echo
    }
    Write-Log -Message ("Query-ADObjects: LDAPFilter={0}" -f $LDAPFilter) -Level "INFO" -ToRun -ToGlobal

    if ($includeBases -and $includeBases.Count -gt 0) {
        foreach ($b in $includeBases) {
            Write-Log -Message ("Scope Include: {0} :: {1}" -f $b.Domain, $b.BaseDN) -Level "INFO" -ToRun -ToGlobal
        }
    } elseif ($Domains -and $Domains.Count -gt 0) {
        Write-Log -Message ("Scope Domains: {0}" -f ($Domains -join ', ')) -Level "INFO" -ToRun -ToGlobal
    } else {
        Write-Log -Message "Scope: NONE (caller must provide SearchBases or Domains)" -Level "WARN" -ToRun -ToGlobal -Echo
    }

    if ($excludeBases -and $excludeBases.Count -gt 0) {
        foreach ($e in $excludeBases) {
            Write-Log -Message ("Scope Exclude: {0} :: {1}" -f $e.Domain, $e.BaseDN) -Level "INFO" -ToRun -ToGlobal
        }
    }

    $targets = @()
    if ($includeBases -and $includeBases.Count -gt 0) {
        $targets = $includeBases | ForEach-Object { @{ Domain = $_.Domain; BaseDN = $_.BaseDN } }
    } elseif ($Domains -and $Domains.Count -gt 0) {
        foreach ($d in $Domains) {
            $root = Get-DomainRootDN -ServerFqdn $d
            if ([string]::IsNullOrWhiteSpace($root)) {
                Write-Log -Message ("Unable to resolve root DN for domain {0}" -f $d) -Level "WARN" -ToRun -ToGlobal -Echo
                continue
            }
            $targets += @{ Domain = $d; BaseDN = $root }
        }
    } else {
        $targets += @{ Domain = $null; BaseDN = $null }
    }

    foreach ($t in $targets) {
        $server    = $t.Domain
        $base      = $t.BaseDN
        $svrLabel  = if ($server) { $server } else { 'default' }
        $baseLabel = if ($base) { " / $base" } else { "" }

        try {
            $msg = if ($server -and $base) {
                "Querying AD: Server=$server Base=$base"
            } elseif ($server) {
                "Querying AD: Server=$server (root)"
            } elseif ($base) {
                "Querying AD: Base=$base (default server)"
            } else {
                "Querying AD: default server & base"
            }
            Write-Log -Message $msg -Level "INFO" -ToRun -ToGlobal

            $args = @{
                LDAPFilter      = $LDAPFilter
                Properties      = $Properties
                SearchScope     = 'Subtree'
                ResultPageSize  = $PageSize     # <-- FIX: map to -ResultPageSize
                ErrorAction     = 'Stop'
            }
            if ($server) { $args['Server'] = $server }
            if ($base)   { $args['SearchBase'] = $base }

            $batch = Get-ADObject @args

            foreach ($obj in $batch) {
                if ($null -eq $obj -or [string]::IsNullOrWhiteSpace($obj.DistinguishedName)) { continue }
                $inScope = Test-DistinguishedNameInScope `
                              -DistinguishedName $obj.DistinguishedName `
                              -IncludeBases $includeBases `
                              -ExcludeBases $excludeBases
                if ($inScope) { [void]$all.Add($obj) }
            }

            $count = if ($batch) { $batch.Count } else { 0 }
            Write-Log -Message ("Fetched {0} objects from {1}{2}" -f $count, $svrLabel, $baseLabel) -Level "INFO" -ToRun -ToGlobal

        } catch {
            Write-Log -Message ("AD query failed for {0}{1}: {2}" -f $svrLabel, $baseLabel, $_.Exception.Message) -Level "ERROR" -ToRun -ToGlobal -Echo
        }
    }

    return $all.ToArray()
}
