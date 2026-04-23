#Requires -Version 5.1
#Requires -Modules @{ModuleName='PSWriteHTML';ModuleVersion='1.17.0'}

param(
    [System.IO.DirectoryInfo]$XMLPath,
    [System.IO.DirectoryInfo]$OutputPath = "$PSScriptRoot\output",
    [string]$DateFormat = 'yyyy-MM-dd',
    [string]$URI = 'https://blog.metsys.fr',
    [string]$Logo = 'https://www.metsys.fr/wp-content/themes/metsys/images/svg/metsys-logo-white.svg',
    [string]$Author = 'METSYS',
    [int]$MaxWidth = 1400,
    [switch]$DoNotShow,
    [switch]$InvertChartLine
)

function Get-File {
    param (
        [string]$Directory = 'C:\',
        [string]$Filter = 'All files (*.*)|*.*'
    )

    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = (Get-Item $Directory).FullName
    $OpenFileDialog.Filter = $Filter
    $OpenFileDialog.Multiselect = $true
    $null = $OpenFileDialog.ShowDialog()
    
    $OpenFileDialog.FileNames
}

$Colors = @{
    'Primary'   = '#783CBD'
    'Secondary' = '#3845AB'
    'Neutral'   = '#3D3834'
    'Positive'  = '#CFE9CF'
    'Negative'  = '#FFCECE'
    'Level1'    = '#F94144'
    'Level2'    = '#F8961E'
    'Level3'    = '#F9C74F'
    'Level4'    = '#43AA8B'
    'Level5'    = '#277DA1'
    'Highest'   = 'darkred'
    'High'      = 'darkorange'
    'Medium'    = 'darkgoldenrod'
    'Low'       = 'darkgreen'
    'Lowest'    = 'darkcyan'
}

$Palette = @(
    '#F94144',
    '#90BE6D',
    '#256EFF'
    '#F9C74F',
    '#DA659A',
    '#43AA8B',
    '#1e93c5',
    '#8338EC',
    '#A9DEF9',
    '#F9844A'
)

$PSDefaultParameterValues = @{
    'New-HTMLSection:HeaderBackGroundColor' = $Colors.Neutral
    'New-HTMLSection:HeaderTextSize'        = 16
    'New-HTMLSection:Margin'                = 20
    'New-ChartBar:Color'                    = $Colors.Primary
    'New-ChartLine:Color'                   = $Colors.Primary
    'New-HTMLTable:HTML'                    = { {
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 1 -BackgroundColor $Colors.Level1
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 2 -BackgroundColor $Colors.Level2
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 3 -BackgroundColor $Colors.Level3
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 4 -BackgroundColor $Colors.Level4
            New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 5 -BackgroundColor $Colors.Level5
        } }
    'New-HTMLTable*:WarningAction'          = 'SilentlyContinue'
    'New-HTMLGage:MinValue'                 = 0
    'New-HTMLGage:MaxValue'                 = 100
    'New-HTMLGage:Pointer'                  = $true
}

if (!$XMLPath) {
    $xmlFiles = Get-File -Directory $PSScriptRoot -Filter 'Extensible Markup Language (*.xml)|*.xml' | ForEach-Object { Get-Item -Path $_ }
}
else {
    $xmlFiles = Get-ChildItem -Path $xmlPath -Filter '*.xml' -Recurse
}

$hcRules = Import-Csv -Path "$PSScriptRoot\data\HCRules.csv" -Delimiter ';' -Encoding utf8
$functionalLevels = 'Windows2000', 'Windows2003Interim', 'Windows2003', 'Windows2008', 'Windows2008R2', 'Windows2012', 'Windows2012R2', 'Windows2016', 'Windows2025'

if (!(Test-Path -Path $OutputPath.FullName -PathType Container)) {
    $null = New-Item -Path $OutputPath.FullName -ItemType Directory
}

$reports = $xmlFiles | ForEach-Object {
    [PSCustomObject]@{
        Domain     = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/DomainFQDN').Node.'#text'
        Date       = Get-Date (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/GenerationDate').Node.'#text'
        Version    = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/EngineVersion').Node.'#text'
        Maturity   = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/MaturityLevel').Node.'#text'
        DomainMode = $functionalLevels[(Select-Xml -Path $_.FullName -XPath '/HealthcheckData/DomainFunctionalLevel').Node.'#text']
        ForestMode = $functionalLevels[(Select-Xml -Path $_.FullName -XPath '/HealthcheckData/ForestFunctionalLevel').Node.'#text']
        Scores     = [PSCustomObject]@{
            Global           = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/GlobalScore').Node.'#text'
            StaleObjects     = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/StaleObjectsScore').Node.'#text'
            PrivilegiedGroup = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/PrivilegiedGroupScore').Node.'#text'
            Trust            = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/TrustScore').Node.'#text'
            Anomaly          = [int](Select-Xml -Path $_.FullName -XPath '/HealthcheckData/AnomalyScore').Node.'#text'
        }
        RiskRules  = (Select-Xml -Path $_.FullName -XPath '/HealthcheckData/RiskRules/HealthcheckRiskRule').Node | ForEach-Object {
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

# Add month and year properties to reports
$reports | Add-Member -MemberType NoteProperty -Name 'Month' -Value $null -Force
$reports | ForEach-Object { $_.Month = Get-Date $_.Date -Format 'yyyy-MM' }

# Keep only the last report for each domain and month
$reports = $reports | Sort-Object -Property Date -Descending | Group-Object -Property Domain, Month | ForEach-Object { $_.Group | Select-Object -First 1 }

# Add empty report for each month and domain without report
$nullReports = $reports.Domain | Sort-Object -Unique | ForEach-Object {
    $domain = $_
    $reports.Month | Sort-Object -Unique | ForEach-Object {
        $month = $_
        if (!($reports | Where-Object { $_.Domain -eq $domain -and $_.Month -eq $month })) {
            [PSCustomObject]@{
                Domain     = $domain
                Date       = Get-Date $month
                Month      = $month
                DomainMode = 'Unknown'
                ForestMode = 'Unknown'
                Maturity   = 0
                Scores     = [PSCustomObject]@{
                    Global           = 0
                    StaleObjects     = 0
                    PrivilegiedGroup = 0
                    Trust            = 0
                    Anomaly          = 0
                }
                RiskRules  = $null
            }
        }
    }
}

# Final sorting
$reports = $reports + $nullReports | Sort-Object -Property Date

$domains = $reports.Domain | Sort-Object -Unique
$allRiskRules = $reports.RiskRules | Sort-Object -Unique -Property RiskId

# Limit to 10 domains for better readability of charts
if ($domains.Count -gt 10) {
    Write-Warning "More than 10 domains found, only the 10 first will be kept for charts readability"
    $domains = $domains | Select-Object -First 10
}

New-HTML -Name 'Global - PingCastle dashboard' -FilePath "$OutputPath\dashboard_global.html" -Encoding UTF8 -Author $Author -DateFormat 'dd/MM/yyyy HH:mm:ss' {
       
    # Header
    New-HTMLHeader -HTMLContent {
        $domain = 'Global report'
        $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\header.html"))
    }

    # Main
    New-HTMLMain {

        New-HTMLTab -Name 'Home' -IconSolid home {

            $firstReports = $domains | ForEach-Object { $domain = $_ ; ($reports | Where-Object { $_.Domain -eq $domain -and $_.DomainMode -ne 'Unknown' } | Select-Object -First 1) }
            $lastReports = $domains | ForEach-Object { $domain = $_ ; ($reports | Where-Object { $_.Domain -eq $domain -and $_.DomainMode -ne 'Unknown' } | Select-Object -Last 1) }
            $ref = $firstReports.RiskRules.RiskId | Sort-Object -Unique
            $dif = $lastReports.RiskRules.RiskId | Sort-Object -Unique
            $comp = Compare-Object -ReferenceObject $ref -DifferenceObject $dif
            $riskSolvedSince = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                $riskId = $_.InputObject
                $lastAppearance = ($reports | Where-Object { $_.RiskRules.RiskId -eq $riskId })[-1].Month
                $lastDomain = ($reports | Where-Object { $_.RiskRules.RiskId -eq $riskId })[-1].Domain
                $allRiskRules | Where-Object { $_.RiskId -eq $riskId } | Select-Object *,
                @{Name = 'LastAppearance'; Expression = { $lastAppearance } },
                @{Name = 'LastDomain'; Expression = { $lastDomain } }
            }
            $riskRulesEvolution = $allRiskRules | Select-Object * -ExcludeProperty Points
            $domains | ForEach-Object { $riskRulesEvolution | Add-Member -Name $_ -MemberType NoteProperty -Value $null }
            $lastReports | ForEach-Object {
                $report = $_
                $member = $_.Domain
                $riskRulesEvolution | ForEach-Object {
                    $riskId = $_.RiskId
                    $_.$member = ($report.RiskRules | Where-Object { $_.RiskId -eq $riskId }).Points
                }
            }
            $riskRulesEvolution | Add-Member -Name 'Global' -MemberType NoteProperty -Value $null
            $riskRulesEvolution | ForEach-Object {
                $risk = $_
                $global = $domains | ForEach-Object { $risk.$_ }
                if (($global | Measure-Object).Count -ge 1) { $_.Global = ($global | Measure-Object -Sum).Sum }
            }

            $chartAxisX = $reports.Month | Sort-Object -Unique
            
            # Diagram for global score
            New-HTMLSection -HeaderText 'Evolution of global score and criticity rule matching' {
                New-HTMLPanel {
                    New-HTMLSection -Invisible {
                        New-HTMLChart -Title 'Uncapped total of the four items score' {
                            New-ChartAxisX -Name $chartAxisX
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                            $i = 0
                            $domains | ForEach-Object {
                                $domain = $_
                                $chartLineTotal = $reports | Where-Object { $_.Domain -eq $domain } | ForEach-Object { ($_.RiskRules.Points | Measure-Object -Sum).Sum }
                                New-ChartLine -Value $chartLineTotal -Name $domain -Color $Palette[$i]
                                $i++
                            }
                        }
                    }
                    New-HTMLSection -Invisible {
                        New-HTMLChart -Title 'Domain maturity' {
                            New-ChartAxisX -Name $chartAxisX
                            New-ChartAxisY -MinValue 0 -MaxValue 4 -Reversed
                            $i = 0
                            $domains | ForEach-Object {
                                $domain = $_
                                $chartLineMaturity = ($reports | Where-Object { $_.Domain -eq $domain }).Maturity
                                New-ChartLine -Value $chartLineMaturity -Name $domain -Curve stepline -Color $Palette[$i]
                                $i++
                            }
                        }
                        New-HTMLChart -Title 'Criticity rule matching' {
                            New-ChartAxisX -Name $chartAxisX
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                            $i = 0
                            $domains | ForEach-Object {
                                $domain = $_
                                $chartLineCriticity = ($reports | Where-Object { $_.Domain -eq $domain }) | ForEach-Object { ($_.RiskRules | Measure-Object).Count }
                                New-ChartLine -Value $chartLineCriticity -Name $domain -Color $Palette[$i]
                                $i++
                            }
                        }
                    }
                }
            }

            # Diagrams per category
            New-HTMLSection -HeaderText 'Evolution per category' {
                New-HTMLPanel {
                    New-HTMLChart -Title 'Anomalies' {
                        New-ChartAxisX -Name $chartAxisX
                        if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                        $i = 0
                        $domains | ForEach-Object {
                            $domain = $_
                            $chartLineAnoma = ($reports | Where-Object { $_.Domain -eq $domain }).Scores.Anomaly
                            New-ChartLine -Value $chartLineAnoma -Name $domain -Color $Palette[$i]
                            $i++
                        }
                    }
                    New-HTMLChart -Title 'Privileged Accounts' {
                        New-ChartAxisX -Name $chartAxisX
                        if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                        $i = 0
                        $domains | ForEach-Object {
                            $domain = $_
                            $chartLinePrivi = ($reports | Where-Object { $_.Domain -eq $domain }).Scores.PrivilegiedGroup
                            New-ChartLine -Value $chartLinePrivi -Name $domain -Color $Palette[$i]
                            $i++
                        }
                    }
                }
                New-HTMLPanel {
                    New-HTMLChart -Title 'Stale Objects' {
                        New-ChartAxisX -Name $chartAxisX
                        if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                        $i = 0
                        $domains | ForEach-Object {
                            $domain = $_
                            $chartLineStale = ($reports | Where-Object { $_.Domain -eq $domain }).Scores.StaleObjects
                            New-ChartLine -Value $chartLineStale -Name $domain -Color $Palette[$i]
                            $i++
                        }
                    }
                    New-HTMLChart -Title 'Trusts' {
                        New-ChartAxisX -Name $chartAxisX
                        if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                        $i = 0
                        $domains | ForEach-Object {
                            $domain = $_
                            $chartLineTrust = ($reports | Where-Object { $_.Domain -eq $domain }).Scores.Trust
                            New-ChartLine -Value $chartLineTrust -Name $domain -Color $Palette[$i]
                            $i++
                        }
                    }
                }
            }

            # Remediations
            New-HTMLSection -HeaderText 'Remediations' {
                New-HTMLTable -Title 'All risks solved' -DataTable $riskSolvedSince -DefaultSortIndex 1 -DisablePaging
            }

            # Risk rules evolution
            New-HTMLSection -HeaderText 'Risk rules evolution' {
                New-HTMLTable -DataTable $riskRulesEvolution -DefaultSortIndex 0 -DisablePaging {
                    New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 1 -BackgroundColor $Colors.Level1
                    New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 2 -BackgroundColor $Colors.Level2
                    New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 3 -BackgroundColor $Colors.Level3
                    New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 4 -BackgroundColor $Colors.Level4
                    New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 5 -BackgroundColor $Colors.Level5
                    New-HTMLTableCondition -Name 'Global' -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray'
                    $domains | ForEach-Object { New-HTMLTableCondition -Name $_ -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray' }
                    New-HTMLTableCondition -Name 'Global' -ComparisonType string -Operator eq -Value '' -Color 'darkgray' -Row
                }
            }
        }

        # Create a new tab for all other reports
        $domains | ForEach-Object {

            $domain = $_
            $currentReport = $reports | Where-Object { $_.Domain -eq $domain -and $_.DomainMode -ne 'Unknown' } | Select-Object -Last 1

            New-HTMLTab -Name $domain -IconSolid globe {

                # Show main informations about report and domain
                New-HTMLSection -HeaderText 'Latest report and domain information' -Direction column {
                    New-HTMLSection -Invisible {
                        New-HTMLPanel {
                            $mainInfo = [PSCustomObject]@{
                                "PingCastle version" = $currentReport.Version
                                "Generated on"       = Get-Date $currentReport.Date -Format D
                                "Report age"         = "$([int]((New-TimeSpan -Start $currentReport.Date).TotalDays)) day(s)"
                                "Domain maturity"    = $currentReport.Maturity
                                "Domain mode"        = $currentReport.DomainMode
                                "Forest mode"        = $currentReport.ForestMode
                            }
                            # Using raw HTML instead of New-HTMLImage because -Inline parameter doesn't work as excpected
                            '<img src="https://raw.githubusercontent.com/leobouard/PingCastleDashboard/refs/heads/main/data/ping-castle-logo.png" alt="PingCastle logo" style="height:32px; margin: 2em auto; display: block;" />'
                            New-HTMLTable -Title 'Report information' -DataTable $mainInfo -HideFooter -Transpose -Simplify
                        }
                        New-HTMLPanel {
                            New-HTMLGage -Label 'Global score' -Value $currentReport.Scores.Global
                            New-HTMLText -Alignment center -TextBlock { 'The worst score out of the four items' }
                        }
                        New-HTMLPanel {
                            New-HTMLChart {
                                New-ChartPie -Value ($currentReport.RiskRules | Where-Object { $_.Level -eq 1 }).Count -Name 'Criticity 1' -Color $Colors.Level1
                                New-ChartPie -Value ($currentReport.RiskRules | Where-Object { $_.Level -eq 2 }).Count -Name 'Criticity 2' -Color $Colors.Level2
                                New-ChartPie -Value ($currentReport.RiskRules | Where-Object { $_.Level -eq 3 }).Count -Name 'Criticity 3' -Color $Colors.Level3
                                New-ChartPie -Value ($currentReport.RiskRules | Where-Object { $_.Level -eq 4 }).Count -Name 'Criticity 4' -Color $Colors.Level4
                                New-ChartPie -Value ($currentReport.RiskRules | Where-Object { $_.Level -eq 5 }).Count -Name 'Criticity 5' -Color $Colors.Level5
                            }
                        }
                    }

                    New-HTMLSection -Invisible {
                        New-HTMLPanel {
                            $totalScore = ($currentReport.RiskRules.Points | Measure-Object -Sum).Sum
                            New-HTMLText -Text "$totalScore pt(s)" -Alignment center -FontSize 22 -FontWeight bold
                            New-HTMLText -Text "Total score" -Alignment center -FontSize 12
                        }
                        1..5 | ForEach-Object {
                            $i = $_
                            New-HTMLPanel {
                                $critXpoints = (($currentReport.RiskRules | Where-Object { $_.Level -eq $i }).Points | Measure-Object -Sum).Sum
                                New-HTMLText -Alignment center {
                                    New-HTMLText -Text "$critXpoints pt(s)" -Alignment center -FontSize 22 -FontWeight bold
                                    '<span style="text-align:center;font-size:12px;color:#ffffff;padding: 2px;background-color:{0};border-radius:2px;">Criticity {1}</span>' -f $Colors."Level$i", $i
                                }
                            }
                        }
                    }

                    New-HTMLSection -HeaderText 'Scores per model' -HeaderBackgroundColor White -HeaderTextColor $Colors.Neutral -CanCollapse -Collapsed {
                        New-HTMLPanel {
                            $perModel = $currentReport.RiskRules.Model | Sort-Object -Unique | ForEach-Object {
                                $model = $_
                                [PSCustomObject]@{
                                    Category = ($currentReport.RiskRules | Where-Object { $_.Model -eq $model })[0].Category
                                    Model    = $model
                                    Points   = [int]($currentReport.RiskRules | Where-Object { $_.Model -eq $model } | Measure-Object -Sum -Property Points).Sum
                                    Count    = [int]($currentReport.RiskRules | Where-Object { $_.Model -eq $model } | Measure-Object).Count
                                }
                            }
                            $perModel = $perModel | Where-Object { $_.Points -ne 0 } | Sort-Object -Property Points -Descending
                            New-HTMLTable -Title 'Point distribution per model' -DataTable $perModel -PagingLength 10 -HideFooter -HideButtons -DisableSearch
                        }
                        New-HTMLPanel {
                            New-HTMLChart -Title 'Point distribution per model' {
                                $i = 0
                                New-ChartLegend -Name $perModel.Model -LegendPosition bottom
                                $perModel | ForEach-Object {
                                    if ($i -ge $Palette.Count) { $i = 0 }
                                    New-ChartPie -Value $_.Points -Name $_.Model -Color $Palette[$i]
                                    $i++
                                }
                            }
                        }
                    }
                }

                # Show PingCastle scores between 0 and 100 pts
                New-HTMLSection -HeaderText 'Scores' {
                    New-HTMLPanel {
                        New-HTMLGage -Label 'Anomalies' -Value $currentReport.Scores.Anomaly
                        New-HTMLText -Alignment center -TextBlock { 'Specific security control points' }
                    }
                    New-HTMLPanel {
                        New-HTMLGage -Label 'Privileged Accounts' -Value $currentReport.Scores.PrivilegiedGroup
                        New-HTMLText -Alignment center -TextBlock { 'Administrators of the Active Directory' }
                    }
                    New-HTMLPanel {
                        New-HTMLGage -Label 'Stale Objects' -Value $currentReport.Scores.StaleObjects
                        New-HTMLText -Alignment center -TextBlock { 'Operations related to user or computer objects' }
                    }
                    New-HTMLPanel {
                        New-HTMLGage -Label 'Trusts' -Value $currentReport.Scores.Trust
                        New-HTMLText -Alignment center -TextBlock { 'Connections between two Active Directories' }
                    }
                }

                # Show all risk rules
                New-HTMLSection -HeaderText 'All current risk rules' -Direction column {
                    New-HTMLTable -DataTable $currentReport.RiskRules -DefaultSortIndex 1 -DisablePaging
                }
            }
        }
    }
    
    # Footer
    New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\footer.html")) }
}

$newInlineCss = '<div data-panes="true" style="max-width: ' + $MaxWidth + 'px; margin: 0 auto;">'
$content = (Get-Content -Path "$OutputPath\dashboard_global.html") -replace '<div data-panes="true">', $newInlineCss
$content | Set-Content -Path "$OutputPath\dashboard_global.html" -Encoding utf8
if (!$DoNotShow.IsPresent) { Start-Process "$OutputPath\dashboard_global.html" }
