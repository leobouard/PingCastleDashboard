#Requires -Version 5.1
<#
.SYNOPSIS
    Checks that every PowerShell command of data\hardening.csv is valid, without running it.

.DESCRIPTION
    For each check: the code must parse, every cmdlet must exist and every parameter must be a real
    parameter of that cmdlet (checked with Get-Command, so the ActiveDirectory and GroupPolicy modules
    must be installed). Words inside -Filter { ... } blocks are AD filter terms and are skipped.
    Used by the CI pipeline. Nothing is executed.
#>
param([string]$Csv = (Join-Path $PSScriptRoot '..\data\hardening.csv'))

$native = 'auditpol', 'repadmin', 'certutil', 'netdom', 'wbadmin'
$problems = [System.Collections.Generic.List[string]]::new()
$rows = Import-Csv -Path $Csv -Delimiter ';' -Encoding UTF8

foreach ($row in $rows) {
    $tokens = $null; $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($row.Check, [ref]$tokens, [ref]$errs)
    if ($errs) { $problems.Add("$($row.RiskId): parse error: $($errs[0].Message)"); continue }

    $commands = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($cmd in $commands) {
        # skip AD filter scriptblocks: -Filter { Attribute -eq $value }
        $inFilter = $false; $node = $cmd.Parent
        while ($node) {
            if ($node -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) {
                $owner = $node.Parent
                if ($owner -is [System.Management.Automation.Language.CommandAst]) {
                    $i = $owner.CommandElements.IndexOf($node)
                    if ($i -gt 0 -and $owner.CommandElements[$i - 1] -is [System.Management.Automation.Language.CommandParameterAst] -and
                        $owner.CommandElements[$i - 1].ParameterName -eq 'Filter') { $inFilter = $true }
                }
                break
            }
            $node = $node.Parent
        }
        if ($inFilter) { continue }

        $name = $cmd.GetCommandName()
        if (-not $name -or $name -in $native) { continue }
        $info = Get-Command -Name $name -ErrorAction SilentlyContinue
        if (-not $info) { $problems.Add("$($row.RiskId): unknown command '$name'"); continue }
        if ($info.CommandType -eq 'Alias') { $info = $info.ResolvedCommand }
        $valid = @($info.Parameters.Keys) + @($info.Parameters.Values | ForEach-Object { $_.Aliases })
        foreach ($p in $cmd.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] }) {
            if ($valid -notcontains $p.ParameterName) { $problems.Add("$($row.RiskId): '$name' has no parameter -$($p.ParameterName)") }
        }
    }
}

"PowerShell $($PSVersionTable.PSVersion) - $($rows.Count) checks validated, $($problems.Count) problem(s)"
$problems | ForEach-Object { "  $_" }
if ($problems.Count) { exit 1 }
