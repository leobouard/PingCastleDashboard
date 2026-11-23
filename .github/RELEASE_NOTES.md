New design of the PingCastle dashboard, with AD hardening help. Same data as the original dashboard: every table and chart value was checked against it.

**What's new in 2.1.1**
- Gauges, donuts and line charts now grow to fill their box, on every screen size and in expanded mode
- The PingCastle logo is visible again in the Midnight theme

**2.1.0**
- Every PingCastle rule (193) now has a fix, a PowerShell check, an effort level and a time estimate, in English and French. No more empty "How to fix" cells
- The checks now also cover AD CS templates and web enrollment, DNS zones, dSHeuristics, RODC, delegations, SID History and trusts

**Since 2.0.0**
- 3 themes (Midnight, Slate, Report), switch at the top right
- English and French, switch at the top left (`-Language EN|FR`)
- AD hardening plan, a **How to fix** column, and a **Hardening guide** tab with a fix, a PowerShell check and a source link for every rule
- **Quick wins**: low effort fixes that save the most points, with a time estimate and the scores you would get after them
- Animations, hover effects, expandable panels (charts redraw at their real size), sortable and searchable tables, CSV export
- No module needed anymore. The original PSWriteHTML dashboard is still there with `-Classic`
- `tests\Test-HardeningChecks.ps1` to run all the checks on a DC of a test domain

**Fixed from the original**
- The "Point distribution per model" pie was always empty
- The dashboard crashed in the browser when there was only one XML report

**Files**
- `PingCastleDashboard-<version>.zip`: the scripts, data files, sample XML reports and tests
- `sample-dashboard-en.html` / `sample-dashboard-fr.html`: the result on the sample reports. Download and open in a browser

Every release is built and tested by the CI pipeline: PSScriptAnalyzer, Windows PowerShell 5.1 and PowerShell 7 builds, the classic PSWriteHTML mode, all checks validated against the real ActiveDirectory, GroupPolicy and DnsServer modules, and a browser test of every tab, theme and language.
