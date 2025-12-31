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
        IgnoredRiskRules = $null
    }
}

# Handle exceptions
$exceptions = Import-Csv -Path "$PSScriptRoot\data\exceptions.csv" -Delimiter ";" -Encoding UTF8
$reports | ForEach-Object {
    $domain = $_.Domain
    $domainExceptions = $exceptions | Where-Object { $_.Domain -eq $domain -or $_.Domain -eq '*' }
    $_.IgnoredRiskRules = $_.RiskRules | Where-Object { $_.RiskId -in $domainExceptions.RiskId }
    $_.RiskRules = $_.RiskRules | Where-Object { $_.RiskId -notin $domainExceptions.RiskId }
}

$reports = $reports | Sort-Object Date

# Create one dashboard foreach domain
$reports.Domain | Sort-Object -Unique | ForEach-Object {

    $domain = $_
    $domainReports = $reports | Where-Object { $_.Domain -eq $domain }
    $allRiskRules = $domainReports.RiskRules | Sort-Object -Unique -Property RiskId

    New-HTML -Name 'PingCastle dashboard' -FilePath "$OutputPath\dashboard_$domain.html" -Encoding UTF8 -Author $Author -DateFormat 'dd/MM/yyyy HH:mm:ss' {
        
        # Header
        New-HTMLHeader -HTMLContent { 
            $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\header.html"))
        }

        # Main
        New-HTMLMain {

            # Home tab
            New-HTMLTab -Name 'Home' -IconSolid home {

                $firstReport = $domainReports[0]
                $lastReport = $domainReports[-1]
                $comp = Compare-Object -ReferenceObject $firstReport.RiskRules.RiskId -DifferenceObject $lastReport.RiskRules.RiskId
                $riskSolvedSince = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                    $riskId = $_.InputObject
                    $lastAppearance = Get-Date ($domainReports | Where-Object { $_.RiskRules.RiskId -eq $riskId })[-1].Date -Format $DateFormat
                    $allRiskRules | Where-Object { $_.RiskId -eq $riskId } | Select-Object *, @{Name = 'LastAppearance'; Expression = { $lastAppearance } }
                }

                $scores = $domainReports | ForEach-Object {
                    [PSCustomObject]@{
                        Date                  = Get-Date $_.Date -Format $DateFormat
                        Maturity              = $_.Maturity
                        'Global score'        = $_.Scores.Global
                        'Total score'         = ($_.RiskRules.Points | Measure-Object -Sum).Sum
                        Anomalies             = (($_.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum
                        'Privileged Accounts' = (($_.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum
                        'Stale Objects'       = (($_.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum
                        Trusts                = (($_.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum
                    }
                }

                $riskRulesEvolution = $allRiskRules | Select-Object * -ExcludeProperty Points
                $domainReports | Sort-Object Date -Descending | ForEach-Object {
                    $name = (Get-Date $_.Date -Format $DateFormat)
                    $riskRulesEvolution | Add-Member -Name $name -MemberType NoteProperty -Value $null
                }
                $domainReports | ForEach-Object {
                    $report = $_
                    $member = (Get-Date $report.Date -Format $DateFormat)
                    $riskRulesEvolution | ForEach-Object {
                        $riskId = $_.RiskId
                        $_.$member = ($report.RiskRules | Where-Object { $_.RiskId -eq $riskId }).Points
                    }
                }

                $chartAxisX = $domainReports | ForEach-Object { Get-Date $_.date -Format $DateFormat }
                $chartLineTotal = $domainReports | ForEach-Object { ($_.RiskRules.Points | Measure-Object -Sum).Sum }
                $chartLineCriticity1 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 1 } | Measure-Object).Count }
                $chartLineCriticity2 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 2 } | Measure-Object).Count }
                $chartLineCriticity3 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 3 } | Measure-Object).Count }
                $chartLineCriticity4 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 4 } | Measure-Object).Count }
                $chartLineCriticity5 = $domainReports | ForEach-Object { ($_.RiskRules.Level | Where-Object { $_ -eq 5 } | Measure-Object).Count }
                $chartLineMaturity = $domainReports.Maturity
                $chartLineAnoma = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum }
                $chartLinePrivi = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum }
                $chartLineStale = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum }
                $chartLineTrust = $domainReports | ForEach-Object { (($_.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum }

                # Diagram for global score
                New-HTMLSection -HeaderText 'Evolution of global score and criticity rule matching' {
                    New-HTMLPanel {
                        New-HTMLSection -Invisible {
                            New-HTMLChart -Title 'Uncapped total of the four items score' {
                                New-ChartAxisX -Name $chartAxisX
                                if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                                New-ChartLine -Value $chartLineTotal -Name 'Point(s)'
                            }
                        }
                        New-HTMLSection -Invisible {
                            New-HTMLChart -Title 'Domain maturity' {
                                New-ChartAxisX -Name $chartAxisX
                                New-ChartAxisY -MinValue 1 -MaxValue 4 -Reversed
                                New-ChartLine -Value $chartLineMaturity -Curve stepline -Color $Colors.Level1
                            }
                            New-HTMLChart -Title 'Criticity level rule matching' {
                                New-ChartAxisX -Name $chartAxisX
                                if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                                New-ChartLine -Value $chartLineCriticity1 -Name 'Criticity 1' -Color $Colors.Level1
                                New-ChartLine -Value $chartLineCriticity2 -Name 'Criticity 2' -Color $Colors.Level2
                                New-ChartLine -Value $chartLineCriticity3 -Name 'Criticity 3' -Color $Colors.Level3
                                New-ChartLine -Value $chartLineCriticity4 -Name 'Criticity 4' -Color $Colors.Level4
                                New-ChartLine -Value $chartLineCriticity5 -Name 'Criticity 5' -Color $Colors.Level5
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
                            New-ChartLine -Value $chartLineAnoma -Name 'Point(s)'
                        }
                        New-HTMLChart -Title 'Privileged Accounts' {
                            New-ChartAxisX -Name $chartAxisX
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                            New-ChartLine -Value $chartLinePrivi -Name 'Point(s)'
                        }
                    }
                    New-HTMLPanel {
                        New-HTMLChart -Title 'Stale Objects' {
                            New-ChartAxisX -Name $chartAxisX
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                            New-ChartLine -Value $chartLineStale -Name 'Point(s)'
                        }
                        New-HTMLChart -Title 'Trusts' {
                            New-ChartAxisX -Name $chartAxisX
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Reversed }
                            New-ChartLine -Value $chartLineTrust -Name 'Point(s)'
                        }
                    }
                }

                # Remediations
                New-HTMLSection -HeaderText 'Remediations' {
                    New-HTMLTable -Title 'All risks solved' -DataTable $riskSolvedSince -DefaultSortIndex 1 -DisablePaging
                }

                # Scores
                New-HTMLSection -HeaderText 'Score & maturity evolution (uncapped)' {
                    New-HTMLTable -DataTable $scores -DefaultSortIndex 0 -DisablePaging {
                        # Maturity
                        New-HTMLTableCondition -Name 'Maturity' -ComparisonType number -Operator eq -Value 1 -BackgroundColor $Colors.Level1
                        New-HTMLTableCondition -Name 'Maturity' -ComparisonType number -Operator eq -Value 2 -BackgroundColor $Colors.Level2
                        New-HTMLTableCondition -Name 'Maturity' -ComparisonType number -Operator eq -Value 3 -BackgroundColor $Colors.Level3
                        New-HTMLTableCondition -Name 'Maturity' -ComparisonType number -Operator eq -Value 4 -BackgroundColor $Colors.Level4
                        New-HTMLTableCondition -Name 'Maturity' -ComparisonType number -Operator eq -Value 5 -BackgroundColor $Colors.Level5
                        # All scores
                        'Global score', 'Anomalies', 'Privileged Accounts', 'Stale Objects', 'Trusts' | ForEach-Object {
                            New-HTMLTableCondition -Name $_ -ComparisonType number -Operator gt -Value 0 -Color $Colors.Lowest -FontWeight bold
                            New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 25 -Color $Colors.Low -FontWeight bold
                            New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 50 -Color $Colors.Medium -FontWeight bold
                            New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 75 -Color $Colors.High -FontWeight bold
                            New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 100 -Color $Colors.Highest -FontWeight bold
                        }
                    }
                }

                # Risk rules evolution
                New-HTMLSection -HeaderText 'Risk rules evolution' {
                    New-HTMLTable -DataTable $riskRulesEvolution -DefaultSortIndex 0 -DisablePaging {
                        New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 1 -BackgroundColor $Colors.Level1
                        New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 2 -BackgroundColor $Colors.Level2
                        New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 3 -BackgroundColor $Colors.Level3
                        New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 4 -BackgroundColor $Colors.Level4
                        New-HTMLTableCondition -Name 'Level' -ComparisonType number -Operator eq -Value 5 -BackgroundColor $Colors.Level5
                        $domainReports.Date | ForEach-Object { Get-Date $_ -Format $DateFormat } | ForEach-Object {
                            New-HTMLTableCondition -Name $_ -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray'
                        }
                        # Grayed the row how have been resolved
                        New-HTMLTableCondition -Name (Get-Date $domainReports.Date[-1] -Format $DateFormat) -ComparisonType string -Operator eq -Value '' -Color 'darkgray' -Row
                    }
                }
            }

            # Create a new tab for all other reports
            $i = 0
            $domainReports | ForEach-Object {
    
                $currentReport = $_
                if ($i -gt 0) { 
                    $previousReport = $domainReports[$i - 1]

                    $comp = Compare-Object -ReferenceObject $previousReport.RiskRules.RiskId -DifferenceObject $currentReport.RiskRules.RiskId
                    $riskSolved = $comp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object {
                        $riskId = $_.InputObject
                        $allRiskRules | Where-Object { $_.RiskId -eq $riskId }
                    }
                    $riskNew = $comp | Where-Object { $_.SideIndicator -eq '=>' } | ForEach-Object {
                        $riskId = $_.InputObject
                        $allRiskRules | Where-Object { $_.RiskId -eq $riskId }
                    }
                }

                New-HTMLTab -Name (Get-Date $_.date -Format $DateFormat) {

                    # Show main informations about report and domain
                    New-HTMLSection -HeaderText 'Report and domain information' -Direction column {
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
                                New-HTMLImage -Source 'https://www.pingcastle.com/wp/wp-content/uploads/2024/07/PC_Logo_PNG-300x253.png' -Height 80
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
                                $prevTotalScore = ($previousReport.RiskRules.Points | Measure-Object -Sum).Sum
                                if ($previousReport) {
                                    if ($prevTotalScore -lt $totalScore) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px;"><path d="M384 160c-17.7 0-32-14.3-32-32s14.3-32 32-32H544c17.7 0 32 14.3 32 32V288c0 17.7-14.3 32-32 32s-32-14.3-32-32V205.3L342.6 374.6c-12.5 12.5-32.8 12.5-45.3 0L192 269.3 54.6 406.6c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3l160-160c12.5-12.5 32.8-12.5 45.3 0L320 306.7 466.7 160H384z"/></svg>' }
                                    if ($prevTotalScore -gt $totalScore) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px;"><path d="M384 352c-17.7 0-32 14.3-32 32s14.3 32 32 32H544c17.7 0 32-14.3 32-32V224c0-17.7-14.3-32-32-32s-32 14.3-32 32v82.7L342.6 137.4c-12.5-12.5-32.8-12.5-45.3 0L192 242.7 54.6 105.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0L320 205.3 466.7 352H384z"/></svg>' }
                                    if ($prevTotalScore -eq $totalScore) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="height: 22px;"><path d="M502.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-128-128c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L402.7 224 32 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l370.7 0-73.4 73.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l128-128z"/></svg>' }
                                }
                                else {
                                    $svg = $null
                                }
                                
                                @'
                                <div style="justify-content:center;align-items: center;width: 100%;display: flex;">
                                  {0}
                                  <div class="defaultText">
                                    <div align="center">
                                      <span style="font-weight:bold;text-align:center;font-size:22px">{1} pt(s)</span>
                                    </div>
                                  </div>
                                </div>
'@ -f $svg, $totalScore
                                New-HTMLText -Text "Total score" -Alignment center -FontSize 12
                            }
                            1..5 | ForEach-Object {
                                $i = $_
                                New-HTMLPanel {
                                    $critXpoints = (($currentReport.RiskRules | Where-Object { $_.Level -eq $i }).Points | Measure-Object -Sum).Sum
                                    $prevCritXpoints = (($previousReport.RiskRules | Where-Object { $_.Level -eq $i }).Points | Measure-Object -Sum).Sum
                                    if ($previousReport) {
                                        if ($prevCritXpoints -lt $critXpoints) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px;"><path d="M384 160c-17.7 0-32-14.3-32-32s14.3-32 32-32H544c17.7 0 32 14.3 32 32V288c0 17.7-14.3 32-32 32s-32-14.3-32-32V205.3L342.6 374.6c-12.5 12.5-32.8 12.5-45.3 0L192 269.3 54.6 406.6c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3l160-160c12.5-12.5 32.8-12.5 45.3 0L320 306.7 466.7 160H384z"/></svg>' }
                                        if ($prevCritXpoints -gt $critXpoints) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px;"><path d="M384 352c-17.7 0-32 14.3-32 32s14.3 32 32 32H544c17.7 0 32-14.3 32-32V224c0-17.7-14.3-32-32-32s-32 14.3-32 32v82.7L342.6 137.4c-12.5-12.5-32.8-12.5-45.3 0L192 242.7 54.6 105.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0L320 205.3 466.7 352H384z"/></svg>' }
                                        if ($prevCritXpoints -eq $critXpoints) { $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="height: 22px;"><path d="M502.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-128-128c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L402.7 224 32 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l370.7 0-73.4 73.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l128-128z"/></svg>' }
                                    }
                                    else {
                                        $svg = $null
                                    }
                                    
                                    @'
                                    <div style="justify-content:center;align-items: center;width: 100%;display: flex;">
                                      {0}
                                      <div class="defaultText">
                                        <div align="center">
                                          <span style="font-weight:bold;text-align:center;font-size:22px">{1} pt(s)</span>
                                        </div>
                                      </div>
                                    </div>
'@ -f $svg, $critXpoints
                                    New-HTMLText -Alignment center {
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
                                $otherThreshold = ($perModel.Points | Measure-Object -Sum).Sum * 0.05
                                $otherModel = [PSCustomObject]@{
                                    Model  = 'Other'
                                    Points = [int]($perModel | Where-Object { $_.Points -lt $otherThreshold } | Measure-Object -Sum -Property Points).Sum
                                }
                                $perModel = $perModel | Where-Object { $_.Points -ge $otherThreshold } | Select-Object Model, Points
                                $perModel += $otherModel
                                New-HTMLChart -Title 'Point distribution per model' {
                                    New-ChartLegend -Name $perModel.Model -LegendPosition bottom
                                    $perModel | ForEach-Object {
                                        New-ChartPie -Name $_.Model -Value $_.Points
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

                    # Show evolution per item between initial, previous and current report
                    New-HTMLSection -HeaderText 'Comparison with previous reports' -Direction column {
                        New-HTMLText -Text 'The evolution of the uncapped score in each category. The score may vary from the normal Ping Castle score (capped to 100), as some rules can be ignored using the exceptions.csv file. The list of ignored risk rules is available at the bottom of the page.'
                        New-HTMLSection -Invisible {
                            New-HTMLPanel {
                                New-HTMLChart -Title 'Anomalies' {
                                    New-ChartBarOptions -Vertical
                                    if ($i -gt 1) { New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum }
                                    if ($i -gt 0) { New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum }
                                    New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'Anomalies' }).Points | Measure-Object -Sum).Sum
                                }
                            }
                            New-HTMLPanel {
                                New-HTMLChart -Title 'Privileged Accounts' {
                                    New-ChartBarOptions -Vertical
                                    if ($i -gt 1) { New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum }
                                    if ($i -gt 0) { New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum }
                                    New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'PrivilegedAccounts' }).Points | Measure-Object -Sum).Sum
                                }
                            }
                            New-HTMLPanel {
                                New-HTMLChart -Title 'Stale Objects' {
                                    New-ChartBarOptions -Vertical
                                    if ($i -gt 1) { New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum }
                                    if ($i -gt 0) { New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum }
                                    New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'StaleObjects' }).Points | Measure-Object -Sum).Sum
                                }
                            }
                            New-HTMLPanel {
                                New-HTMLChart -Title 'Trusts' {
                                    New-ChartBarOptions -Vertical
                                    if ($i -gt 1) { New-ChartBar -Name 'Initial' -Value (($domainReports[0].RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum }
                                    if ($i -gt 0) { New-ChartBar -Name 'Previous' -Value (($previousReport.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum }
                                    New-ChartBar -Name 'Current' -Value (($currentReport.RiskRules | Where-Object { $_.Category -eq 'Trusts' }).Points | Measure-Object -Sum).Sum
                                }
                            }
                        }
                    }

                    # Comparison between previous and current report
                    New-HTMLSection -HeaderText 'Improvement & deterioration' {
                        New-HTMLPanel {
                            # The following risk rules have been resolved since the last report (improvement)
                            New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Positive {
                                New-HTMLHeading h2 -HeadingText 'Risk rules resolved'
                            }
                            New-HTMLTable -DataTable $riskSolved -DefaultSortIndex 1 -HideButtons
                        }
                        New-HTMLPanel {
                            # The following risk rules have been discovered since the last report (deterioration)
                            New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Negative {
                                New-HTMLHeading h2 -HeadingText 'New risk rules triggered'
                            }
                            New-HTMLTable -DataTable $riskNew -DefaultSortIndex 1 -HideButtons
                        }
                    }

                    # Show all risk rules
                    New-HTMLSection -HeaderText 'All current risk rules' -Direction column {
                        New-HTMLTable -DataTable $currentReport.RiskRules -DefaultSortIndex 1 -DisablePaging
                    }

                    # Show ignored risk rules
                    if ($currentReport.IgnoredRiskRules) {
                        New-HTMLSection -HeaderText 'Ignored risk rules' -Direction column {
                            New-HTMLText -Text 'The following rules have been excluded from the calculated scores using the "exceptions.csv" file.'
                            New-HTMLTable -DataTable $currentReport.IgnoredRiskRules -DefaultSortIndex 1 -DisablePaging
                        }
                    }
                }

                $i++
            }
        }

        # Footer
        New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\footer.html")) }
    }
}

$reports.Domain | Sort-Object -Unique | ForEach-Object {

    $domain = $_
    $newInlineCss = '<div data-panes="true" style="max-width: ' + $MaxWidth + 'px; margin: 0 auto;">'
    $content = (Get-Content -Path "$OutputPath\dashboard_$domain.html") -replace '<div data-panes="true">', $newInlineCss
    $content | Set-Content -Path "$OutputPath\dashboard_$domain.html" -Encoding utf8
    if (!$DoNotShow.IsPresent) { Start-Process "$OutputPath\dashboard_$domain.html" }

}