#Requires -Version 5.1
<#
.SYNOPSIS
    Runs every PowerShell check of data\hardening.csv and reports the ones that fail.

.DESCRIPTION
    Use it on a domain controller (or an admin server with RSAT) of a test domain, to make sure
    every check of the Hardening guide works on a real Active Directory.
    All the checks are read-only: they only read AD, GPOs, the registry or certificate stores.
    Some checks read the local registry or local services, so they give the real answer only on a DC.

.EXAMPLE
    .\tests\Test-HardeningChecks.ps1
    .\tests\Test-HardeningChecks.ps1 -RiskId A-Krbtgt, P-SchemaAdmin -ShowOutput
#>
param(
    [string]$Csv = (Join-Path $PSScriptRoot '..\data\hardening.csv'),
    [string[]]$RiskId,
    [switch]$ShowOutput,
    [string]$ExportPath
)

$rows = Import-Csv -Path $Csv -Delimiter ';' -Encoding UTF8
if ($RiskId) { $rows = $rows | Where-Object { $_.RiskId -in $RiskId } }

foreach ($module in 'ActiveDirectory', 'GroupPolicy') {
    if (-not (Get-Module -ListAvailable -Name $module)) { Write-Warning "Module $module is not installed: checks that use it will fail (install RSAT)." }
}

$results = foreach ($row in $rows) {
    $errorsBefore = $Error.Count
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $status = 'OK'; $message = ''; $output = @(); $global:LASTEXITCODE = 0
    try {
        $output = @(& ([scriptblock]::Create($row.Check)) 2>&1)
        $errs = @($output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
        $output = @($output | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })
        if ($errs) { $status = 'ERROR'; $message = $errs[0].ToString() }
        elseif ($LASTEXITCODE -and $row.Check -match '^(auditpol|repadmin|certutil)') { $status = 'ERROR'; $message = "Exit code $LASTEXITCODE" }
    }
    catch { $status = 'ERROR'; $message = $_.Exception.Message }
    $watch.Stop()
    # errors written with -ErrorAction SilentlyContinue are expected (for example Enterprise Admins in a child domain)
    $Error.RemoveRange(0, [Math]::Max(0, $Error.Count - $errorsBefore))

    if ($ShowOutput -and $output) {
        Write-Host "`n--- $($row.RiskId)" -ForegroundColor Cyan
        $output | Out-String -Width 300 | Out-Host
    }
    [PSCustomObject]@{
        RiskId  = $row.RiskId
        Status  = $status
        Results = $output.Count
        Seconds = [Math]::Round($watch.Elapsed.TotalSeconds, 1)
        Message = $message
    }
}

$results | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Out-Host
$failed = @($results | Where-Object Status -ne 'OK')
"{0} checks run, {1} OK, {2} with errors" -f $results.Count, ($results.Count - $failed.Count), $failed.Count
if ($ExportPath) { $results | Export-Csv -Path $ExportPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 }
if ($failed) { exit 1 }
