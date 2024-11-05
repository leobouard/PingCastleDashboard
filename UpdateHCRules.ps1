$uri = "https://api.github.com/repos/netwrix/pingcastle/contents/Healthcheck/Rules"
$results = Invoke-RestMethod -Method GET -Uri $uri | Where-Object { $_.Name -like '*.cs' }

$i = 0
$total = ($results | Measure-Object).Count


$HCRules = $results | ForEach-Object {
    
    # Progress bar
    $i++
    $percent = $i / $total * 100
    Write-Progress -Activity $_.Name -PercentComplete $percent

    # Get RAW code
    $content = (Invoke-RestMethod -Method GET -Uri $_.'download_url')
    $content = $content -split ' '

    # RiskId
    [string]$riskId = $content | Select-String -SimpleMatch 'RuleModel'
    $riskId = ($riskId.Replace('[RuleModel("', '')).Replace('",', '')

    # Level
    [string]$anssiLevel = $content | Select-String -SimpleMatch 'RuleDurAnssi'
    if ($anssiLevel) {
        [string]$level = $anssiLevel.Replace('[RuleDurANSSI(', '')
    }
    else {
        [string]$level = $content | Select-String -SimpleMatch 'RuleMaturityLevel'
        [string]$level = $level.Replace('[RuleMaturityLevel(', '')
    }

    if ($riskId) {
        # Create the PSCustomObject
        [PSCustomObject]@{
            RiskId = $riskId
            Level  = $level.SubString(0, 1)
        } 
    }
}

$ref = Get-Content -Path "$PSScriptRoot\data\HCRules.csv" -Encoding utf8
$dif = $HCRules | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
$comp = Compare-Object -ReferenceObject $ref -DifferenceObject $dif
Write-Output "$(($comp | Measure-Object).Count) difference(s) found!"
$comp | Format-Table

$HCRules | Sort-Object RiskId | Export-Csv -Path "$PSScriptRoot\data\HCRules.csv" -Delimiter ';' -Encoding utf8 -NoTypeInformation

pause
