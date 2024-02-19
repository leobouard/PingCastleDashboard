#Requires -Version 5.1
#Requires -Modules @{ModuleName='PSWriteHTML';ModuleVersion='1.17.0'}

$xmlFiles = Get-ChildItem -Path .\xml -Filter '*.xml'

$reports = $xmlFiles | ForEach-Object {

    $domain = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/DomainFQDN').Node.'#text'
    $date   = Get-Date (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/GenerationDate').Node.'#text'

    $xpath = '/HealthcheckData/RiskRules/HealthcheckRiskRule'
    $riskRules = (Select-Xml -Path $_.FullName -XPath $xpath).Node

    [PSCustomObject]@{
        Domain    = $domain
        Date      = $date
        RiskRules = $riskRules
    }
    
}

$comp = Compare-Object -ReferenceObject $reports[0].RiskRules -DifferenceObject $reports[1].RiskRules
# Issues/misconfiguration that have been solved
$old = ($comp | Where-Object {$_.SideIndicator -eq '=>'}).InputObject
# New issues or misconfiguration found
$new = ($comp | Where-Object {$_.SideIndicator -eq '<='}).InputObject





<#
($reports.RiskRules.Points | Measure-Object -Sum).Sum
(($reports.RiskRules | Where-Object {$_.Category -eq 'Anomalies'}).Points | Measure-Object -Sum).Sum
(($reports.RiskRules | Where-Object {$_.Category -eq 'PrivilegedAccounts'}).Points | Measure-Object -Sum).Sum
(($reports.RiskRules | Where-Object {$_.Category -eq 'StaleObjects'}).Points | Measure-Object -Sum).Sum
(($reports.RiskRules | Where-Object {$_.Category -eq 'Trusts'}).Points | Measure-Object -Sum).Sum

Dashboard -Name 'PingCastle dashboard' -FilePath '.\dashboard.html' {
    Section -Invisible {
        Panel -Invisible {
            Chart -Title 'Evolution of the cumulated points' {
                ChartAxisX -Name 'aoû. 20','sept. 20'
                ChartLine -Name 'Total' -Value 766,694
                ChartLine -Name 'Anomalies' -Value 184,184
                ChartLine -Name 'PrivilegedAccounts' -Value 360,300
                ChartLine -Name 'StaleObjects' -Value 97,90
                ChartLine -Name 'Trusts' -Value 125,120
            }
        }
    }
    Section -Invisible {
        Panel -Invisible {
            Chart -Title 'Incidents Reported vs Solved' -TitleAlignment center {
                ChartAxisX -Name 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'
                ChartLine -Name 'Incidents per month' -Value 10, 41, 35, 51, 49, 62, 69, 91, 148
                ChartLine -Name 'Incidents per month resolved' -Value 5, 10, 20, 31, 49, 62, 69, 91, 148
            }
        }
    }
}
#>