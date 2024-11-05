# PingCastleDashboard

Generates an HTML dashboard to review the evolution of PingCastle scores and metrics using the XML files generated by PingCastle.

**This script requires the module `PSWriteHTML`.** You can install the module through PSGallery using `Install-Module PSWriteHTML`. More information about this module here: <https://github.com/EvotecIT/PSWriteHTML>

## How to use it?

You simply have to gather the XML files generated by PingCastle in one folder and execute the `PingCastleDashboard.ps1` script. One report will be generated for each domain.

## Results

You can check an overview here (based on fictitious data): <https://www.labouabouate.fr/assets/files/dashboard_test.mysmartlogon.com>

## Parameters

### -XMLPath

Indicates the path to the folder containing PingCastle XML files.\
If you do not use this setting, a file explorer will be displayed.

### -OutputPath

Indicates path to destination folder for generated files (dashboard and JSON).\
Default value is `.\output`

### -DateFormat

Date format used in the dashboard.\
Default value is `yyyy-MM-dd`

## UpdateHCRules

To add criticity scores to the healthcheck rules, the script uses the CSV file `HCRules.csv` which can be outdated. To update the file, you simply have to run the script `UpdateHCRules.ps1`. The script will parse the GitHub repository of the PingCastle project to get the latest healthcheck rules.

## Handle exceptions

In some cases, PingCastle can be a little blind or too severe. If you wish, you can add some risk rules to the `data\exceptions.csv` file to ignore them in the dashboard.

If you wish to add the exception to each domain, you can use the wildcard character (*) in the "Domain" column.

All ID risk rules are available in the `HCRules.csv` file if you need a complete repository.

## Links

- [netwrix/pingcastle: PingCastle - Get Active Directory Security at 80% in 20% of the time](https://github.com/netwrix/pingcastle)
- [EvotecIT/PSWriteHTML: PSWriteHTML is PowerShell Module to generate beautiful HTML reports, pages, emails without any knowledge of HTML, CSS or JavaScript. To get started basics PowerShell knowledge is required.](https://github.com/EvotecIT/PSWriteHTML)
- [Dashimo (PSWriteHTML) - Charting, Icons and few other changes - Evotec](https://evotec.xyz/dashimo-pswritehtml-charting-icons-and-few-other-changes/)
- [Dashimo - Easy Table Conditional Formatting and more - Evotec](https://evotec.xyz/dashimo-easy-table-conditional-formatting-and-more/)

## Special thanks

This project wouldn't be possible without the help and/or the work of colleagues from [METSYS](https://blog.metsys.fr), so thank you to:

- [Thierry PLANTIVE](https://www.linkedin.com/in/thierry-plantive-764b5b93/) for the original idea, feedbacks and features requests
- [stefan-frenchies](https://github.com/stefan-frenchies) for the XML parse code
- [Vincent VUCCINO](https://www.linkedin.com/in/vincent-vuccino-7948762b/) for the HCRule code
- [mgirardi23](https://github.com/mgirardi23) for the exceptions list
