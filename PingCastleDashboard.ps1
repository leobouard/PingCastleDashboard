#Requires -Version 5.1
#Requires -Modules @{ModuleName='PSWriteHTML';ModuleVersion='1.17.0'}

$PSDefaultParameterValues = @{
    'New-HTMLSection:HeaderBackGroundColor' = '#3D3834'
    'New-HTMLSection:HeaderTextSize'        = '16'
    'New-ChartBar:Color'                    = '#783CBD'
    'New-ChartLine:Color'                   = '#783CBD'
    'New-HTMLTable:HTML'                    = { {
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 1 -BackgroundColor '#fd0100'
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 2 -BackgroundColor '#ffa500'
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 3 -BackgroundColor '#f0e68c'
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 4 -BackgroundColor '#007bff'
        } }
    'New-HTMLTable*:WarningAction'          = 'SilentlyContinue'
}

$xmlFiles = Get-ChildItem -Path "$PSScriptRoot\xml" -Filter '*.xml' -Recurse
$hcRules = Import-Csv -Path "$PSScriptRoot\HCRules.csv" -Delimiter ';' -Encoding utf8

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
            $riskId = $_.RiskId
            [PSCustomObject]@{
                Points    = [int]($_.Points)
                Level     = ($hcRules | Where-Object { $_.RiskId -eq $riskId }).Level
                Category  = $_.Category
                Model     = $_.Model
                RiskId    = $riskId
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
    New-HTML -Name 'PingCastle dashboard' -FilePath "$PSScriptRoot\output\dashboard_$domain.html" -Encoding UTF8 -Author 'Léo Bouard' -DateFormat 'yyyy-MM-dd HH:mm:ss' -Show {
        
        # Header
        New-HTMLHeader -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\html\header.html")) }

        # Main
        New-HTMLMain {

            # Home tab
            New-HTMLTab -Name 'Home' -IconSolid home {

                $firstReport = $domainReports[0]
                $lastReport = $domainReports[-1]
                $comp = Compare-Object -ReferenceObject $firstReport.RiskRules.RiskId -DifferenceObject $lastReport.RiskRules.RiskId
                $riskSolvedSince = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                    $riskId = $_.InputObject
                    $lastAppearance = ($domainReports | Where-Object { $_.RiskRules.RiskId -eq $riskId })[-1].Date
                    $allRiskRules | Where-Object { $_.RiskId -eq $riskId } | Select-Object *, @{Name = 'LastAppearance'; Expression = { $lastAppearance } }
                }

                $chartAxisX = $domainReports | ForEach-Object { Get-Date $_.date -Format 'yyyy-MM-dd HH:mm' }

                $chartLineTotal = $domainReports | ForEach-Object { ($_.RiskRules.Points | Measure-Object -Sum).Sum }

                $chartLineMaturity1 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 1 } | Measure-Object -Sum).Sum }
                $chartLineMaturity2 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 2 } | Measure-Object -Sum).Sum }
                $chartLineMaturity3 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 3 } | Measure-Object -Sum).Sum }
                $chartLineMaturity4 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 4 } | Measure-Object -Sum).Sum }

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

                # Diagram for maturity level rule matching
                New-HTMLSection -HeaderText 'Maturity rule matching' {
                    New-HTMLPanel {
                        New-HTMLChart {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartLine -Value $chartLineMaturity1 -Name 'Maturity 1' -Color '#fd0100'
                            New-ChartLine -Value $chartLineMaturity2 -Name 'Maturity 2' -Color '#ffa500'
                            New-ChartLine -Value $chartLineMaturity3 -Name 'Maturity 3' -Color '#f0e68c'
                            New-ChartLine -Value $chartLineMaturity4 -Name 'Maturity 4' -Color '#007bff'
                        }
                    }
                    New-HTMLPanel {
                        New-HTMLChart {
                            New-ChartPie -Value $chartLineMaturity1[-1] -Name 'Maturity 1' -Color '#fd0100'
                            New-ChartPie -Value $chartLineMaturity2[-1] -Name 'Maturity 2' -Color '#ffa500'
                            New-ChartPie -Value $chartLineMaturity3[-1] -Name 'Maturity 3' -Color '#f0e68c'
                            New-ChartPie -Value $chartLineMaturity4[-1] -Name 'Maturity 4' -Color '#007bff'
                        }
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

                # Remediations
                New-HTMLSection -HeaderText 'Remediations' {
                    New-HTMLPanel {
                        New-HTMLHeading h2 -HeadingText 'All risks solved'
                        New-HTMLTable -DataTable $riskSolvedSince -DefaultSortIndex 1
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
                        New-HTMLPanel {
                            New-HTMLGage -Label 'Anomalies' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.Anomaly
                            New-HTMLText -TextBlock { 'It is about specific security control points' }
                        }
                        New-HTMLPanel {
                            New-HTMLGage -Label 'Privileged Accounts' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.PrivilegiedGroup
                            New-HTMLText -TextBlock { 'It is about administrators of the Active Directory' }
                        }
                        New-HTMLPanel {
                            New-HTMLGage -Label 'Stale Objects' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.StaleObjects
                            New-HTMLText -TextBlock { 'It is about operations related to user or computer objects' }
                        }
                        New-HTMLPanel {
                            New-HTMLGage -Label 'Trusts' -MinValue 0 -MaxValue 100 -Value $currentReport.Scores.Trust
                            New-HTMLText -TextBlock { 'It is about connections between two Active Directories' }
                        }
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
                            New-HTMLSection -Invisible -AlignItems center -JustifyContent flex-start -BackgroundColor '#cfe9cf' {
                                New-HTMLFontIcon -IconSize 20 -IconSolid check-circle -IconColor 'Green'
                                New-HTMLHeading h2 -HeadingText 'Risk rules resolved'
                            }
                            New-HTMLTable -DataTable $riskSolved -DefaultSortIndex 1 -HideButtons
                        }
                        New-HTMLPanel {
                            # The following risk rules have been discovered since the last report (deterioration)
                            New-HTMLSection -Invisible -AlignItems center -JustifyContent flex-start -BackgroundColor '#ffcece' {
                                New-HTMLFontIcon -IconSize 20 -IconSolid arrow-circle-down -IconColor 'DarkRed'
                                New-HTMLHeading h2 -HeadingText 'New risk rules triggered'
                            }
                            New-HTMLTable -DataTable $riskNew -DefaultSortIndex 1 -HideButtons
                        }
                    }

                    # Show all risk rules
                    New-HTMLSection -HeaderText 'All current risk rules' {
                        New-HTMLTable -DataTable $currentReport.RiskRules -DefaultSortIndex 1
                    }
                }
            }
        }

        # Footer
        New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\html\footer.html")) }
    }
}