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

$reports = $reports | Sort-Object Date


Dashboard -Name 'PingCastle dashboard' -FilePath '.\dashboard.html' -Show {

    # Main page
    New-HtmlTab -Name 'Main page' {

        $chartAxisX = $reports | ForEach-Object { Get-Date $_.date -Format 'yyyy-MM-dd HH:mm' }
        $chartLineTotal = $reports | ForEach-Object { ($_.RiskRules.Points | Measure-Object -Sum).Sum }
        $chartLineAnoma = $reports | ForEach-Object { (($_.RiskRules | Where-Object {$_.Category -eq 'Anomalies'}).Points | Measure-Object -Sum).Sum }
        $chartLinePrivi = $reports | ForEach-Object { (($_.RiskRules | Where-Object {$_.Category -eq 'PrivilegedAccounts'}).Points | Measure-Object -Sum).Sum }
        $chartLineStale = $reports | ForEach-Object { (($_.RiskRules | Where-Object {$_.Category -eq 'StaleObjects'}).Points | Measure-Object -Sum).Sum }
        $chartLineTrust = $reports | ForEach-Object { (($_.RiskRules | Where-Object {$_.Category -eq 'Trusts'}).Points | Measure-Object -Sum).Sum }

        New-HTMLChart -Title 'Evolution of the cumulated points' {
            New-ChartAxisX -Name $chartAxisX
            New-ChartLine -Value $chartLineTotal -Name 'Total'
            New-ChartLine -Value $chartLineAnoma -Name 'Anomalies'
            New-ChartLine -Value $chartLinePrivi -Name 'PrivilegedAccounts'
            New-ChartLine -Value $chartLineStale -Name 'StaleObjects'
            New-ChartLine -Value $chartLineTrust -Name 'Trusts'
        }
        
    }

    # Initial situation
    New-HtmlTab -Name 'Initial situation' {
        
    }

    # Create a new tab for all other reports
    $i = 0
    $reports | Select-Object -Skip 1 | ForEach-Object {

        New-HtmlTab -Name (Get-Date $_.date -Format 'yyyy-MM-dd HH:mm') {

            $currentReport  = $_
            $previousReport = $reports[$i] 

            # Comparison between current report and previous one
            $comp = Compare-Object -ReferenceObject $previousReport.RiskRules -DifferenceObject $currentReport.RiskRules
            $old = ($comp | Where-Object {$_.SideIndicator -eq '=>'}).InputObject | Select-Object Points,Category,Model,RiskId,Rationale
            $new = ($comp | Where-Object {$_.SideIndicator -eq '<='}).InputObject | Select-Object Points,Category,Model,RiskId,Rationale
        
            New-HTMLTable -Title 'Risk rules resolved' -DataTable $old -DefaultSortIndex Points -HideFooter
            New-HTMLTable -Title 'New risk rules triggered' -DataTable $new -DefaultSortIndex Points -HideFooter
        
        }

        $i++
    }

}





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