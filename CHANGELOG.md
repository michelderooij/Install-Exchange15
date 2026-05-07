# Changelog

All notable changes to Install-Exchange15 are documented here.

## [4.26]
- Corrected documentation: Windows Server 2022 and Windows Server 2025 support both Desktop Experience and Server Core installations, not just Windows Server 2019
- Improved credential security: the AutoPilot password is no longer decrypted to plain text in memory longer than necessary during credential validation
- State file format changed from XML to JSON; any existing state file from a previous run will not be recognised — if you are resuming a multi-phase installation started with an older version, re-run the script from phase 0

## [4.25]
- Fixed Windows Server 2025 not being correctly identified, which could cause the wrong prerequisites or version checks to be applied
- Fixed `-EnableAMSI` switch being ignored after an AutoPilot reboot; AMSI is now correctly configured in all phases
- Dropped all Exchange 2013 references and dead code; Exchange 2013 has not been supported for installation since v4.0
- Updated help text and usage examples to reflect currently supported Exchange versions (2016 CU23, 2019 CU10–CU15, Exchange SE RTM)
- Added CHANGELOG.md; revision history removed from the script itself
- Updated README.md with current OS support matrix, parameter reference, and usage examples

## [4.24]
- Fixed autodiscover SCP configuration

## [4.23]
- Fixed Edge installation (no need checking for Ex2013 in AD)

## [4.22]
- Corrected download VC++2013 runtime URL due to shortcut being unavailable

## [4.21]
- Added disabling MSExchangeAutodiscoverAppPool during setup to prevent responding to requests during setup and postconfig

## [4.20]
- Clearing/setting SCP now background job during install to configure it asynchronous & ASAP

## [4.13]
- Fixed race issue when installing from ISO and restarting installation
- Tested with SW_DVD9_Exchange_Server_Subscription_64bit_MultiLang_Std_Ent_.iso_MLF_X24-08113.iso

## [4.12]
- Fixed feature installation (Web-W-Auth, should be Web-Windows-Auth)
- Using ADSI for Ex2013 detection

## [4.11]
- Fixed feature installation for WS2022/WS2025 Core

## [4.10]
- Added support for Exchange Server SE

## [4.01]
- Removed obsolete TLS13 setup detection

## [4.0]
- Added support for Exchange 2019 CU15
- Added support for Windows Server 2025 (Exchange 2019 CU15+)
- Removed Exchange 2013 support
- Removed Exchange 2016 CU1-22 support
- Removed Exchange 2019 RTM-CU9
- Removed Windows Server 2012 R2 support
- Added removal of obsolete MSMQ feature when installed
- Added EnableECC switch to configure Elliptic Curve Crypto support
- Added NoCBC switch to prevent configuring AES256-CBC-encrypted content support
- Added EnableAMSI switch to configure AMSI body scanning for ECP, EWS, OWA and PowerShell
- Added EnableTLS12 switch to configure TLS12
- Added EnableTLS13 switch to configure TLS13 on WS2022/WS2025 with EX2019CU15+
- Removed InstallMailbox, InstallCAS, InstallMultiRole switches
- Removed NoNet461, NoNet471, NoNet472 and NoNet48 switches
- Removed UseWMF3 switch
- Added Ex2013 detection as it cannot coexist with Ex2019CU15+
- Enabled loading Exchange module in postconf needed for possible override cmdlets
- Removed setup phase shown on wallpaper
- Set minimal required PS version to 5.1
- Code cleanup
- Functions now use approved verbs

## [3.9]
- Added support for Exchange 2019 CU14
- Added support for .NET Framework 4.8.1
- Added NONET481 switch to use .NET 4.8 instead of 4.8.1 for Exchange 2019 CU14+
- Added DoNotEnableEP and DoNotEnableEP_FEEWS switches for Exchange 2019 CU14+
- Added deploying AUG2023 SUs for Ex2019CU13/Ex2019CU12/Ex2016CU23 when IncludeFixes specified
- Changed example to show usage of iso as source
- Added descriptive message when specifying invalid SourcePath
- Fixed detection source path when iso already mounted without drive letter assignment

## [3.8]
- Added support for Exchange 2019 CU13

## [3.71]
- Updated recommended Defender AV inclusions/exclusions

## [3.7]
- Added support for Windows Server 2022
- Fixed logic for installing the IIS Rewrite module for Ex2016CU22+/Ex2019CU11+
- Fixed logic when to use the new /IAcceptExchangeServerLicenseTerms_DiagnosticData* switch

## [3.62]
- Added support for Exchange 2019 CU12
- Added support for Exchange 2016 CU23

## [3.61]
- Added mention of Exchange 2019

## [3.6]
- Added support for Exchange 2019 CU11
- Added support for Exchange 2016 CU22
- Added support for Exchange 2019 CU10
- Added support for Exchange 2019 CU9
- Added support for Exchange 2016 CU21
- Added support for Exchange 2016 CU20
- Added IIS URL Rewrite prereq for Ex2019CU11+ & Ex2016 CU22+
- Added support for KB2999226 on for WS2012R2
- Added DiagnosticData switch to set initial DataCollectionEnabled mode

## [3.5]
- Added support for Exchange 2019 CU8
- Added support for Exchange 2016 CU19
- Added support for KB5003435 for 2019CU8+9,2016CU19+20 and 2013CU23
- Added support for KB5000871 for 2019RTM-CU7, 2016CU8-CU18 and 2013CU21+22
- Added support for Interim Update installation & detection
- Updated .NET 4.8 download link
- Updated Visual C++ 2012 download link (latest release)
- Updated Visual C++ 2013 download link (latest release)
- Corrected not installing KB3206632 on WS2019
- Corrected disabling of Server Manager during setup
- Fixed setting High Performance Plan for recent Windows builds
- Textual corrections

## [3.4]
- Added support for Exchange 2019 CU8
- Added support for Exchange 2016 CU19
- Script allows non-static IP config with service Windows Azure Guest Agent, Network Agent or Telemetry Service present

## [3.3]
- Added support for Exchange 2019 CU7
- Added support for Exchange 2016 CU18

## [3.2.6]
- Added support for Exchange 2019 CU6
- Added support for Exchange 2016 CU17
- Added VC++ Runtime 2012 for Exchange 2019

## [3.2.5]
- Fixed typo in enumeration of Exchange build to report
- Fixed specified vs used MDBLogPath (would add unspecified `<DBNAME>\Log`)

## [3.2.4]
- Added support for Exchange 2019 CU4+CU5
- Added support for Exchange 2016 CU15+CU16

## [3.2.3]
- Fixed typo for Ex2019CU3 detection

## [3.2.2]
- Added support for Exchange 2019 CU3
- Added support for Exchange 2016 CU14

## [3.2.1]
- Updated Pagefile config for Exchange 2019 (25% mem.size)

## [3.2.0]
- Added support for Exchange 2019 CU2
- Added support for Exchange 2016 CU13
- Added support for Exchange 2013 CU23
- Added support for NET Framework 4.8
- Added NoNET48 switch
- Added disabling of Server Manager during installation
- Removed support for Windows Server 2008R2
- Removed support for Windows Server 2012
- Removed Switch UseWMF3

## [3.1.1]
- Fixed detection of Defender

## [3.1.0]
- Added support for Exchange 2019 CU1
- Added support for Exchange 2016 CU12
- Added support for Exchange 2013 CU22
- Fixed Hotfix KB3041832 url
- Fixed NoSetup Mode/EmptyRoles problem
- Added skip Health Monitor checks for InstallEdge
- Fixed potential Exchange version misreporting

## [3.0.4]
- Fixed bug in Install-MyPackage

## [3.0.3]
- Fixed typos in Join-Path constructs

## [3.0.2]
- Replaced filename constructs with Join-Path
- Fixed typo in installing KB4054530

## [3.0.1]
- Integrated Exchange 2019 RTM Cipher correction

## [3.0.0]
- Added Exchange 2019 support
- Rewritten VC++ detection

## [2.99.9]
- Added support for Exchange 2016 CU11
- Updated SourcePath parameter usage (ISO)
- Added .NET Framework 4.7.2 support
- Added Windows Defender presence check

## [2.99.82]
- Added reapplying KB2565063 (MS11-025) to IncludeFixes
- Changed downloading VC++ Package to filename indicating version
- Added post-setup Healthcheck / IIS Warmup

## [2.99.81]
- Fixed phase sequencing with reboot pending

## [2.99.8]
- Updated to Support Edge (Simon Poirier)

## [2.99.7]
- Updated location where hotfix are being published

## [2.99.6]
- Added Exchange 2019 Preview on Windows Server 2019 support (desktop & core)

## [2.99.5]
- Added setting desktop background during setup
- Some code cleanup

## [2.99.4]
- Fixed Recover mode not adding /InstallWindowsComponents
- Added SkipRolesCheck switch
- Added Exchange 2019 Public Preview support on Windows Server 2016

## [2.99.3]
- Fixed TargetPath-Recover parameter mutual exclusion

## [2.99.2]
- Fixed Recover Mode Phase
- Fixed InstallMDBDBPath location check
- Added support for Exchange 2016 CU10
- Added support for Exchange 2013 CU21
- Added Visual C++ 2013 Redistributable prereq (Ex2016CU10+/Ex2013CU21+)
- Fixed Exchange setup result detection
- Changed code to determine AD Configuration container
- Changed script to abort on non-static IP presence
- Removed InstallFilterPack switch (obsolete)
- Code cleanup and cosmetics

## [2.991]
- Fixed .NET blockade removal
- Fixed upgrade detection
- Minor bugs and cosmetics fixes

## [2.99]
- Added Windows Defender exclusions (Ex2016 on WS2016)

## [2.98]
- Added support for Exchange 2016 CU9
- Added support for Exchange 2013 CU20
- Added blocking of .NET Framework 4.7.2 (Preview)
- Added upgrade mode detection
- Added TargetPath usage for Recover mode

## [2.97]
- Added support for Exchange 2016 CU8
- Added support for Exchange 2013 CU19
- Added NONET471 switch

## [2.96]
- Added support for Exchange 2016 CU7
- Added support for Exchange 2013 CU18
- Added FFL 2008R2 checks for Exchange 2016 CU7
- Added blocking of .NET Framework 4.7.1
- Consolidated .NET Framework blocking routines
- Modified version comparison routine

## [2.95]
- Added support for Exchange 2016 CU6
- Added support for Exchange 2013 CU17

## [2.93]
- Added blocking .NET Framework 4.7

## [2.92]
- Cosmetics and code cleanup when installing on WS2016
- Output cosmetics when disabling RC4

## [2.91]
- Added support for Exchange 2016 CU5
- Added support for Exchange 2013 CU16

## [2.9]
- Added support for Exchange 2016 CU4
- Added support for Exchange 2013 CU15
- Added KB3206632 to Exchange 2016 @ WS2016 requirements

## [2.8]
- Added DisableRC4 to disable RC4 (kb2868725)
- Fixed DisableSSL3, removed disabling SSL3 as client
- Disables NIC Power Management during post config

## [2.7]
- Added support for Windows Server 2016 (Exchange Server 2016 CU3+ only)

## [2.6]
- Added support for Exchange 2013 CU14 and Exchange 2016 CU3
- Fixed 7318.DrainNGenQueue routine
- Some minor cosmetics

## [2.54]
- Fixed failing TargetPath check

## [2.53]
- Fixed NoSetup logic skipping NET 4.6.1 installation
- Added .NET framework optimization post-config (7318.DrainNGenQueue)

## [2.52]
- Script will abort when AD site can not be determined
- Fixed SCP parameter handling, use '-' to remove the SCP

## [2.51]
- Script will abort when ExSetup has non-0 exitcode
- Script will ignore package exit codes -2145124329 (SUS_E_NOT_APPLICABLE)

## [2.5]
- Added recommended hotfixes: KB3146717, KB2985459 (WS2012), KB3041832 (WS2012R2), KB3004383 (WS2008R2)
- Added logging of AD Site
- Added computername to filename of state file and log
- Changed credential prompting, will use current account
- Changed Power Plan setting to use InstanceID instead of textual match
- Fixed KeepAlive timeout setting
- Added checks for running as Enterprise & Schema admin
- Fixed NoSetup bug (would abort)
- Added check to see if Exchange server object already exists
- Added Recover switch for RecoverServer mode

## [2.42]
- Bug fix - Installation of KB2919442 only detectable after reboot; adjusted logic
- Added /f (forceAppsClose) for .MSU installations

## [2.41]
- Bug fix - Setup version of Exchange 2013 CU13 is .000, not .003

## [2.4]
- Added support up to Exchange 2013 CU13 / Exchange 2016 CU2
- Added support for .NET 4.6.1 (Exchange 2013 CU13+ / Exchange 2016 CU2+)
- Added NONET461 switch, to use .NET 4.5.2, and block .NET 4.6.1
- Added installation of .NET 4.6.1 OS-dependent required hotfixes
- Added recommended Keep-Alive and RPC timeout settings
- Added DisableSSL3 to disable SSL3 (KB187498)

## [2.31]
- Fixed output error messages

## [2.3]
- Added support up to Exchange 2013 CU12 / Exchange 2016 CU1
- Switched version detection to ExSetup, now follows Build

## [2.2]
- Added (temporary) blocking unsupported .NET Framework 4.6.1 (KB3133990)
- Added recommended updates KB2884597 & KB2894875 for WS2012
- Changes to output so all output/verbose/warning/error get logged
- Added check to Organization for invalid characters
- Fixed specifying an Organization name containing spaces

## [2.12]
- Fixed pre-CU7 .NET installation logic

## [2.11]
- Added Exchange 2016 RTM support
- Removed Exchange 2016 Preview support

## [2.1]
- Replaced ClearSCP with SCP param
- Added Lock switch to lock computer during installation
- Configures High Performance Power plan
- Added installing feature RSAT-Clustering-CmdInterface
- Added pagefile configuration when it's set to 'system managed'

## [2.03]
- Bug & typo fix

## [2.0]
- Renamed script to Install-Exchange15
- Added CU9 support
- Added Exchange Server 2016 Preview support
- Fixed registry checks for GPO error messages
- Added ClearSCP switch to clear Autodiscover SCP record post-setup
- Added Import-ExchangeModule() for post-configuration using EMS
- Bug fix .NET installation
- Modified AD checks to support multi-forest deployments
- Added access checks for Installation, MDB and Log locations
- Added checks for Exchange organization/Organization parameter

## [1.9]
- Added CU8 support
- Fixed CU6/CU7 detection
- Added (temporary) clearing of Execution Policy GPO value
- Added Forest Level check to throw warning when it can't read value
- Added KB2985459 for WS2012
- Using different service to detect installed version
- Installs WMF4/NET452 for supported Exchange versions
- Added UseWMF3 switch to use WMF3 on WS2008R2

## [1.8]
- Added CU7 support

## [1.73]
- Added CU6 support
- Added KB2997355 (Exchange Online mailboxes cannot be managed by using EAC)
- Added .NET Framework 4.52
- Removed DisableRetStructPinning (not required for .NET 4.52 or later)

## [1.72]
- Added CU5 support
- Added KB2971467 (CU5 Disable Shared Cache Service Managed Availability probes)

## [1.71]
- Uncommented RunOnce line - AutoPilot should work again
- Using strings for OS version comparisons (should fix issue w/localized OS)
- Fixed issue installing .NET 4.51 on WS2012
- Fixed inconsistency with .NET detection in WS2012

## [1.7]
- Added Exchange 2013 SP1 & WS2012R2 support
- Added installing .NET Framework 4.51 (2008 R2 & 2012 - 2012R2 has 4.51)
- Added DisableRetStructPinning for Mailbox roles
- Added KB2938053 (SP1 Transport Agent Fix)
- Added switch InstallFilterPack to install Office Filter Pack (OneNote & Publisher support)
- Fixed Exchange failed setup exit code anomaly

## [1.61]
- Fixed XML not found issue when specifying different InstallPath (Cory Wood)

## [1.6]
- Code cleanup (merged KB/QFE/package functions)
- Fixed Verbose setting not being restored when script continues after reboot
- Renamed InstallBoth to InstallMultiRole
- Added 'Yes to All' option to extract function to prevent overwrite popup
- Added detection of setup file version
- Added switch IncludeFixes, which will install recommended hotfixes

## [1.56]
- Changed logic of final cleanup

## [1.55]
- Feature installation bug fix on WS2012

## [1.54]
- Added Parameter InstallBoth to install CAS and Mailbox

## [1.53]
- Fix phase of Forest/Domain Level check

## [1.52]
- Fix .NET / PrepareAD order for WS2008R2, relocated RebootPending check

## [1.51]
- Rewrote Test-Credentials due to missing .NET 3.5 Out of the Box in WS2008R2
- Testing for proper loading of servermanager module in WS2008R2

## [1.5]
- Added support for WS2008R2 (prereqs NET45, WMF3), IEESC toggling
- Added InstallPath to AutoPilot set

## [1.1]
- When used for AD preparation, RSAT-ADDS-Tools won't be uninstalled
- Pending reboot detection; in AutoPilot, script will reboot and restart phase
- Installs Server-Media-Foundation feature (UCMA 4.0 requirement)
- Validates provided credentials for AutoPilot
- Check OS version as string (accommodates non-US OS)

## [1.03]
- Replaced installing most OS features in favor of /InstallWindowsComponents
- Removed installation of Office Filtering Pack

## [1.02]
- Fixed small typo in post-prepare AD function

## [1.01]
- Added logic to prepare AD when organization present
- Fixed checks and logic to prepare AD
- Added testing for domain mixed/native mode
- Added testing for forest functional level

## [1.0]
- Initial community release
