#Requires -Version 5.1
#Requires -Modules @{ModuleName='PSWriteHTML';ModuleVersion='1.17.0'}

$PSDefaultParameterValues = @{
    'New-HTMLSection:HeaderBackGroundColor' = '#3D3834'
    'New-HTMLSection:HeaderTextSize'        = '16'
    'New-ChartBar:Color'                    = '#783CBD'
    'New-ChartLine:Color'                   = '#783CBD'
}

$xmlFiles = Get-ChildItem -Path .\xml -Filter '*.xml'

$reports = $xmlFiles | ForEach-Object {

    $domain = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/DomainFQDN').Node.'#text'
    $date = Get-Date (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/GenerationDate').Node.'#text'

    $xpath = '/HealthcheckData/RiskRules/HealthcheckRiskRule'
    $riskRules = (Select-Xml -Path $_.FullName -XPath $xpath).Node

    [PSCustomObject]@{
        Domain    = $domain
        Date      = $date
        Scores    = [PSCustomObject]@{
            Global           = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/GlobalScore').Node.'#text'
            StaleObjects     = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/StaleObjectsScore').Node.'#text'
            PrivilegiedGroup = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/PrivilegiedGroupScore').Node.'#text'
            Trust            = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/TrustScore').Node.'#text'
            Anomaly          = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/AnomalyScore').Node.'#text'
        }
        RiskRules = $riskRules | ForEach-Object {
            [PSCustomObject]@{
                Points    = [int]($_.Points)
                Category  = $_.Category
                Model     = $_.Model
                RiskId    = $_.RiskId
                Rationale = $_.Rationale
            }
        }
    }
}

$reports = $reports | Sort-Object Date
$allRiskRules = $reports.RiskRules | Sort-Object -Unique -Property RiskId

# Create one dashboard foreach domain
$reports.Domain | Sort-Object -Unique | ForEach-Object {

    $domain = $_
    $domainReports = $reports | Where-Object { $_.Domain -eq $domain }
    New-HTML -Name 'PingCastle dashboard' -FilePath ".\output\dashboard_$domain.html" -Encoding UTF8 -Author 'Léo Bouard' -DateFormat 'yyyy-MM-dd HH:mm:ss' -Show {
        
        # Header
        New-HTMLHeader -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path '.\html\header.html')) }

        # Main
        New-HTMLMain {

            # Main page tab
            New-HTMLTab -Name 'Main page' -IconSolid home {

                $firstReport = $domainReports[0]
                $lastReport = $domainReports[-1]
                $comp = Compare-Object -ReferenceObject $firstReport.RiskRules.RiskId -DifferenceObject $lastReport.RiskRules.RiskId
                $riskSolvedSince = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                    $riskId = $_.InputObject
                    $lastAppearance = ($domainReports | Where-Object { $_.RiskRules.RiskId -eq $riskId })[-1].Date
                    $allRiskRules | Where-Object { $_.RiskId -eq $riskId } | Select-Object *, @{Name = 'LastAppearance'; Expression = { $lastAppearance } }
                }

                <# $scoreRecap = $domainReports | Select-Object Date,
                    @{N='Global';E={$_.Scores.Global}},
                    @{N='Anomalies';E={$_.Scores.Anomaly}},
                    @{N='PrivilegedAccounts';E={$_.Scores.PrivilegiedGroup}},
                    @{N='StaleObjects';E={$_.Scores.StaleObjects}},
                    @{N='Trusts';E={$_.Scores.Trust}} #>

                $chartAxisX = $domainReports | ForEach-Object { Get-Date $_.date -Format 'yyyy-MM-dd HH:mm' }
                $chartLineTotal = $domainReports | ForEach-Object { ($_.RiskRules.Points | Measure-Object -Sum).Sum }
                $chartLineAnoma = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum }
                $chartLinePrivi = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum }
                $chartLineStale = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum }
                $chartLineTrust = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum }

                # Diagram for global score
                New-HTMLSection -HeaderText 'Evolution of global score' {
                    New-HTMLChart {
                        New-ChartAxisX -Name $chartAxisX
                        New-ChartLine -Value $chartLineTotal -Name 'Point(s)'
                    }
                }

                # Diagrams per category
                New-HTMLSection -HeaderText 'Evolution per category' {
                    New-HTMLPanel {
                        New-HTMLChart -Title 'Anomalies' {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartLine -Value $chartLineAnoma -Name 'Point(s)'
                        }
                        New-HTMLChart -Title 'Privileged Accounts' {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartLine -Value $chartLinePrivi -Name 'Point(s)'
                        }
                    }
                    New-HTMLPanel {
                        New-HTMLChart -Title 'Stale Objects' {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartLine -Value $chartLineStale -Name 'Point(s)'
                        }
                        New-HTMLChart -Title 'Trusts' {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartLine -Value $chartLineTrust -Name 'Point(s)'
                        }
                    }
                }

                # Work-in-progress
                New-HTMLSection -HeaderText 'Remediations' {
                    New-HTMLPanel {
                        New-HTMLHeading h3 -HeadingText 'All risks solved'
                        New-HTMLTable -DataTable $riskSolvedSince -DefaultSortIndex 0 -DefaultSortOrder Descending
                    }
                }
            }

            # Create a new tab for all other reports
            $i = 0
            $domainReports | Select-Object -Skip 1 | ForEach-Object {
    
                $currentReport = $_
                $previousReport = $domainReports[$i]
                $i++

                $comp = Compare-Object -ReferenceObject $previousReport.RiskRules.RiskId -DifferenceObject $currentReport.RiskRules.RiskId
                $riskSolved = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                    $riskId = $_.InputObject
                    $allRiskRules | Where-Object { $_.RiskId -eq $riskId }
                }
                $riskNew = $comp | Where-Object { $_.SideIndicator -eq '=>' } | ForEach-Object {
                    $riskId = $_.InputObject
                    $allRiskRules | Where-Object { $_.RiskId -eq $riskId }
                }

                New-HTMLTab -Name (Get-Date $_.date -Format 'yyyy-MM-dd HH:mm') {

                    # Show PingCastle scores between 0 and 100 pts
                    New-HTMLSection -HeaderText 'Scores' {
                        New-HTMLGage -Label 'Anomalies' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.Anomaly
                        New-HTMLGage -Label 'Privileged Accounts' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.PrivilegiedGroup
                        New-HTMLGage -Label 'Stale Objects' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.StaleObjects
                        New-HTMLGage -Label 'Trusts' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.Trust
                    }

                    # Show evolution per item between initial, previous and current report
                    New-HTMLSection -HeaderText 'Comparison with previous reports' {
                        New-HTMLChart -Title 'Anomalies' {
                            New-ChartBarOptions -Vertical
                            New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum
                        }
                        New-HTMLChart -Title 'Privileged Accounts' {
                            New-ChartBarOptions -Vertical
                            New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum
                        }
                        New-HTMLChart -Title 'Stale Objects' {
                            New-ChartBarOptions -Vertical
                            New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum
                        }
                        New-HTMLChart -Title 'Trusts' {
                            New-ChartBarOptions -Vertical
                            New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum
                            New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum
                        }
                    }

                    # Comparison between previous and current report
                    New-HTMLSection -HeaderText 'Improvement & deterioration' {
                        New-HTMLPanel {
                            # The following risk rules have been resolved since the last report (improvement)
                            New-HTMLHeading h3 -HeadingText 'Risk rules resolved'
                            New-HTMLTable -DataTable $riskSolved -DefaultSortIndex 0 -DefaultSortOrder Descending -HideButtons
                        }
                        New-HTMLPanel {
                            # The following risk rules have been discovered since the last report (deterioration)
                            New-HTMLHeading h3 -HeadingText 'New risk rules triggered'
                            New-HTMLTable -DataTable $riskNew -DefaultSortIndex 0 -DefaultSortOrder Descending -HideButtons
                        }
                    }

                    # Show all risk rules
                    New-HTMLSection -HeaderText 'All current risk rules' {
                        New-HTMLTable -DataTable $currentReport.RiskRules -DefaultSortIndex 0 -DefaultSortOrder Descending
                    }
                }
            }
        }

        # Footer
        New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path '.\html\footer.html')) }
    }
}