<#
    .SYNOPSIS
    Install-Exchange15

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 4.11, July 9th, 2025

    Thanks to Maarten Piederiet, Thomas Stensitzki, Brian Reid, Martin Sieber, Sebastiaan Brozius, Bobby West,
    Pavel Andreev, Rob Whaley, Simon Poirier, Brenle, Eric Vegter and everyone else who provided feedback
    or contributed in other ways.

    .DESCRIPTION
    This script can install Exchange 2016/2019 prerequisites, optionally create the Exchange
    organization (prepares Active Directory) and installs Exchange Server. When the AutoPilot switch is
    specified, it will do all the required rebooting and automatic logging on using provided credentials.
    To keep track of provided parameters and state, it uses an XML file; if this file is
    present, this information will be used to resume the process. Note that you can use a central
    location for Install (UNC path with proper permissions) to re-use additional downloads.

    .LINK
    http://eightwone.com

    .NOTES

    Requirements:
    - Supported Operating Systems
      - Windows Server 2016 (Exchange 2016 CU23)
      - Windows Server 2019 (Desktop or Core, Exchange 2019 only)
      - Windows Server 2022 (Exchange 2019 only)
      - Windows Server 2025 (Exchange 2019 CU15+ only)
    - Domain-joined system, except for Edge Server Role
    - "AutoPilot" mode requires account with elevated administrator privileges
    - When you let the script prepare AD, the account needs proper permissions

    .REVISIONS

    1.0     Initial community release
    1.01    Added logic to prepare AD when organization present
            Fixed checks and logic to prepare AD
            Added testing for domain mixed/native mode
            Added testing for forest functional level
    1.02    Fixed small typo in post-prepare AD function
    1.03    Replaced installing most OS features in favor of /InstallWindowsComponents
            Removed installation of Office Filtering Pack
    1.1     When used for AD preparation, RSAT-ADDS-Tools won't be uninstalled
            Pending reboot detection. In AutoPilot, script will reboot and restart phase.
            Installs Server-Media-Foundation feature (UCMA 4.0 requirement)
            Validates provided credentials for AutoPilot
            Check OS version as string (should accomodate non-US OS)
    1.5     Added support for WS2008R2 (i.e. added prereqs NET45, WMF3), IEESC toggling,
            KB974405, KB2619234, KB2758857 (supersedes KB2533623). Inserted phase for
            WS2008R2 to install hotfixes (+reboot); this phase is skipped for WS2012.
            Added InstallPath to AutoPilot set (or default won't be set).
    1.51    Rewrote Test-Credentials due to missing .NET 3.5 Out of the Box in WS2008R2.
            Testing for proper loading of servermanager module in WS2008R2.
    1.52    Fix .NET / PrepareAD order for WS2008R2, relocated RebootPending check
    1.53    Fix phase of Forest/Domain Level check
    1.54    Added Parameter InstallBoth to install CAS and Mailbox, workaround as PoSHv2
            can discriminate overlapping ParameterSets (resulting in AmbigiousParameterSet)
    1.55    Feature installation bug fix on WS2012
    1.56    Changed logic of final cleanup
    1.6     Code cleanup (merged KB/QFE/package functions)
            Fixed Verbose setting not being restored when script continues after reboot
            Renamed InstallBoth to InstallMultiRole
            Added 'Yes to All' option to extract function to prevent overwrite popup
            Added detection of setup file version
            Added switch IncludeFixes, which will install recommended hotfixes
            (2008R2:KB2803754,KB2862063 2012:KB2803755,KB2862064) and KB2880833 for CU2 & CU3.
    1.61    Fixed XML not found issue when specifying different InstallPath (Cory Wood)
    1.7     Added Exchange 2013 SP1 & WS2012R2 support
            Added installing .NET Framework 4.51 (2008 R2 & 2012 - 2012R2 has 4.51)
            Added DisableRetStructPinning for Mailbox roles
            Added KB2938053 (SP1 Transport Agent Fix)
            Added switch InstallFilterPack to install Office Filter Pack (OneNote & Publisher support)
            Fixed Exchange failed setup exit code anomaly
    1.71    Uncommented RunOnce line - AutoPilot should work again
            Using strings for OS version comparisons (should fix issue w/localized OS)
            Fixed issue installing .NET 4.51 on WS2012 ('all in one' kb2858728 contains/reports
            WS2008R2/kb958488 versus WS2012/kb2881468
            Fixed inconsistency with .NET detection in WS2012
    1.72    Added CU5 support
            Added KB2971467 (CU5 Disable Shared Cache Service Managed Availability probes)
    1.73    Added CU6 support
            Added KB2997355 (Exchange Online mailboxes cannot be managed by using EAC)
            Added .NET Framework 4.52
            Removed DisableRetStructPinning (not required for .NET 4.52 or later)
7    1.8     Added CU7 support
    1.9     Added CU8 support
            Fixed CU6/CU7 detection
            Added (temporary) clearing of Execution Policy GPO value
            Added Forest Level check to throw warning when it can't read value
            Added KB2985459 for WS2012
            Using different service to detect installed version
            Installs WMF4/NET452 for supported Exchange versions
            Added UseWMF3 switch to use WMF3 on WS2008R2
    2.0     Renamed script to Install-Exchange15
            Added CU9 support
            Added Exchange Server 2016 Preview support
            Fixed registry checks for GPO error messages
            Added ClearSCP switch to clear Autodiscover SCP record post-setup
            Added Import-ExchangeModule() for post-configuration using EMS
            Bug fix .NET installation
            Modified AD checks to support multi-forest deployments
            Added access checks for Installation, MDB and Log locations
            Added checks for Exchange organization/Organization parameter
    2.03    Bug & typo fix
    2.1     Replaced ClearSCP with SCP param
            Added Lock switch to lock computer during installation
            Configures High Performance Power plan
            Added installing feature RSAT-Clustering-CmdInterface
            Added pagefile configuration when it's set to 'system managed'
    2.11    Added Exchange 2016 RTM support
            Removed Exchange 2016 Preview support
    2.12    Fixed pre-CU7 .NET installation logic
    2.2     Added (temporary) blocking unsupported .NET Framework 4.6.1 (KB3133990)
            Added recommended updates KB2884597 & KB2894875 for WS2012
            Changes to output so all output/verbose/warning/error get logged
            Added check to Organization for invalid characters
            Fixed specifying an Organization name containing spaces
    2.3     Added support up to Exchange 2013 CU12 / Exchange 2016 CU1
            Switched version detection to ExSetup, now follows Build
    2.31    Fixed output error messages
    2.4     Added support up to Exchange 2013 CU13 / Exchange 2016 CU2
            Added support for .NET 4.6.1 (Exchange 2013 CU13+ / Exchange 2016 CU2+)
            Added NONET461 switch, to use .NET 4.5.2, and block .NET 4.6.1
            Added installation of .NET 4.6.1 OS-dependent required hotfixes:
            * KB2919442 and KB2919355 (~700MB!) for WS2012R2 (prerequisites).
            * KB3146716 for WS2008/WS2008R2, KB3146714 for WS2012, and KB3146715 for WS2012R2.
            Added recommended Keep-Alive and RPC timeout settings
            Added DisableSSL3 to disable SSL3 (KB187498)
    2.41    Bug fix - Setup version of Exchange 2013 CU13 is .000, not .003
    2.42    Bug fix - Installation of KB2919442 only detectable after reboot; adjusted logic
            Added /f (forceAppsClose) for .MSU installations
    2.5     Added recommended hotfixes:
            * KB3146717 (=offline version of 3146718)
            * KB2985459 (WS2012)
            * KB3041832 (WS2012R2)
            * KB3004383 (WS2008R2)
            Added logging of AD Site
            Added computername to filename of state file and log
            Changed credential prompting, will use current account
            Changed Power Plan setting to use InstanceID instead of textual match
            Fixed KeepAlive timeout setting
            Added checks for running as Enterpise & Schema admin
            Fixed NoSetup bug (would abort)
            Added check to see if Exchange server object already exists
            Added Recover switch for RecoverServer mode
    2.51    Script will abort when ExSetup has non-0 exitcode
            Script will ignore package exit codes -2145124329 (SUS_E_NOT_APPLICABLE)
    2.52    Script will abort when AD site can not be determined
            Fixed SCP parameter handling, use '-' to remove the SCP
    2.53    Fixed NoSetup logic skipping NET 4.6.1 installation
            Added .NET framework optimization post-config (7318.DrainNGenQueue)
    2.54    Fixed failing TargetPath check
    2.6     Added support for Exchange 2013 CU14 and Exchange 2016 CU3
            Fixed 7318.DrainNGenQueue routine
            Some minor cosmetics
    2.7     Added support for Windows Server 2016 (Exchange Server 2016 CU3+ only)
    2.8     Added DisableRC4 to disable RC4 (kb2868725)
            Fixed DisableSSL3, removed disabling SSL3 as client
            Disables NIC Power Management during post config
    2.9     Added support for Exchange 2016 CU4
            Added support for Exchange 2013 CU15
            Added KB3206632 to Exchange 2016 @ WS2016 requirements
    2.91    Added support for Exchange 2016 CU5
            Added support for Exchange 2013 CU16
    2.92    Cosmetics and code cleanup when installing on WS2016
            Output cosmetics when disabling RC4
    2.93    Added blocking .NET Framework 4.7
    2.95    Added support for Exchange 2016 CU6
            Added support for Exchange 2013 CU17
    2.96    Added support for Exchange 2016 CU7
            Added support for Exchange 2013 CU18
            Added FFL 2008R2 checks for Exchange 2016 CU7
            Added blocking of .NET Framework 4.7.1
            Consolidated .NET Framework blocking routines
            Modified version comparison routine
    2.97    Added support for Exchange 2016 CU8
            Added support for Exchange 2013 CU19
            Added NONET471 switch
    2.98    Added support for Exchange 2016 CU9
            Added support for Exchange 2013 CU20
            Added blocking of .NET Framework 4.7.2 (Preview)
            Added upgrade mode detection
            Added TargetPath usage for Recover mode
    2.99    Added Windows Defender exclusions (Ex2016 on WS2016)
    2.991   Fixed .NET blockade removal
            Fixed upgrade detection
            Minor bugs and cosmetics fixes
    2.99.2  Fixed Recover Mode Phase
            Fixed InstallMDBDBPath location check
            Added support for Exchange 2016 CU10
            Added support for Exchange 2013 CU21
            Added Visual C++ 2013 Redistributable prereq (Ex2016CU10+/Ex2013CU21+)
            Fixed Exchange setup result detection
            Changed code to determine AD Configuration container
            Changed script to abort on non-static IP presence
            Removed InstallFilterPack switch (obsolete)
            Code cleanup and cosmetics
    2.99.3  Fixed TargetPath-Recover parameter mutual exclusion
    2.99.4  Fixed Recover mode not adding /InstallWindowsComponents
            Added SkipRolesCheck switch
            Added Exchange 2019 Public Preview support on Windows Server 2016
    2.99.5  Added setting desktop background during setup
            Some code cleanup
    2.99.6  Added Exchange 2019 Preview on Windows Server 2019 support (desktop & core)
    2.99.7  Updated location where hotfix are being published
    2.99.8  Updated to Support Edge (Simon Poirier)
    2.99.81 Fixed phase sequencing with reboot pending
    2.99.82 Added reapplying KB2565063 (MS11-025) to IncludeFixes
            Changed downloading VC++ Package to filename indicating version
            Added post-setup Healthcheck / IIS Warmup
    2.99.9  Added support for Exchange 2016 CU11
            Updated SourcePath parameter usage (ISO)
            Added .NET Framework 4.7.2 support
            Added Windows Defender presence check
    3.0.0   Added Exchange 2019 support
            Rewritten VC++ detection
    3.0.1   Integrated Exchange 2019 RTM Cipher correction
    3.0.2   Replaced filename constructs with Join-Path
            Fixed typo in installing KB4054530
    3.0.3   Fixed typos in Join-Path constructs
    3.0.4   Fixed bug in Install-MyPackage
    3.1.0   Added support for Exchange 2019 CU1
            Added support for Exchange 2016 CU12
            Added support for Exchange 2013 CU22
            Fixed Hotfix KB3041832 url
            Fixed NoSetup Mode/EmptyRoles problem
            Added skip Health Monitor checks for InstallEdge
            Fixed potential Exchange version misreporting
    3.1.1   Fixed detection of Defender
    3.2.0   Added support for Exchange 2019 CU2
            Added support for Exchange 2016 CU13
            Added support for Exchange 2013 CU23
            Added support for NET Framework 4.8
            Added NoNET48 switch
            Added disabling of Server Manager during installation
            Removed support for Windows Server 2008R2
            Removed support for Windows Server 2012
            Removed Switch UseWMF3
    3.2.1   Updated Pagefile config for Exchange 2019 (25% mem.size)
    3.2.2   Added support for Exchange 2019 CU3
            Added support for Exchange 2016 CU14
    3.2.3   Fixed typo for Ex2019CU3 detection
    3.2.4   Added support for Exchange 2019 CU4+CU5
            Added support for Exchange 2016 CU15+CU16
    3.2.5   Fixed typo in enumeration of Exchange build to report
            Fixed specified vs used MDBLogPath (would add unspecified <DBNAME>\Log)
    3.2.6   Added support for Exchange 2019 CU6
            Added support for Exchange 2016 CU17
            Added VC++ Runtime 2012 for Exchange 2019
    3.3     Added support for Exchange 2019 CU7
            Added support for Exchange 2016 CU18
    3.4     Added support for Exchange 2019 CU8
            Added support for Exchange 2016 CU19
            Script allows non-static IP config with service Windows Azure Guest Agent, Network Agent or Telemetry Service present
    3.5     Added support for Exchange 2019 CU8
            Added support for Exchange 2016 CU19
            Added support for KB5003435 for 2019CU8+9,2016CU19+20 and 2013CU23
            Added support for KB5000871 for 2019RTM-CU7, 2016CU8-CU18 and 2013CU21+22
            Added support for Interim Update installation & detection
            Updated .NET 4.8 download link
            Updated Visual C++ 2012 download link (latest release)
            Updated Visual C++ 2013 download link (latest release)
            Corrected not installing KB3206632 on WS2019
            Corrected disabling of Server Manager during setup
            Fixed setting High Performance Plan for recent Windows builds
            Textual corrections
    3.6     Added support for Exchange 2019 CU11
            Added support for Exchange 2016 CU22
            Added support for Exchange 2019 CU10
            Added support for Exchange 2019 CU9
            Added support for Exchange 2016 CU21
            Added support for Exchange 2016 CU20
            Added IIS URL Rewrite prereq for Ex2019CU11+ & Ex2016 CU22+
            Added support for KB2999226 on for WS2012R2
            Added DiagnosticData switch to set initial DataCollectionEnabled mode
    3.61    Added mention of Exchange 2019
    3.62    Added support for Exchange 2019 CU12
            Added support for Exchange 2016 CU23
    3.7     Added support for Windows Server 2022
            Fixed logic for installing the IIS Rewrite module for Ex2016CU22+/Ex2019CU11+
            Fixed logic when to use the new /IAcceptExchangeServerLicenseTerms_DiagnosticData* switch
    3.71    Updated recommended Defender AV inclusions/exclusions
    3.8     Added support for Exchange 2019 CU13
    3.9     Added support for Exchange 2019 CU14
            Added support for .NET Framework 4.8.1
            Added NONET481 switch to use .NET 4.8 instead of 4.8.1 for Exchange 2019 CU14+
            Added DoNotEnableEP and DoNotEnableEP_FEEWS switches for Exchange 2019 CU14+
            Added deploying AUG2023 SUs for Ex2019CU13/Ex2019CU12/Ex2016CU23 when IncludeFixes specified
            Changed example to show usage of iso as source
            Added descriptive message when specifying invalid SourcePath
            Fixed detection source path when iso already mounted without drive letter assignment
    4.0     Added support for Exchange 2019 CU15
            Added support for Windows Server 2025 (Exchange 2019 CU15+)
            Removed Exchange 2013 support
            Removed Exchange 2016 CU1-22 support
            Removed Exchange 2019 RTM-CU9
            Removed Windows Server 2012 R2 support
            Added removal of obsolete MSMQ feature when installed
            Added EnableECC switch to configure Elliptic Curve Crypto support
            Added NoCBC switch to prevent configuring AES256-CBC-encrypted content support
            Added EnableAMSI switch to configure AMSI body scanning for ECP, EWS, OWA and PowerShell
            Added EnableTLS12 switch to configure TLS12
            Added EnableTLS13 switch to configure TLS13 on WS2022/WS2025 with EX2019CU15+
            Removed InstallMailbox, InstallCAS, InstallMultiRole switches
            Removed NoNet461, NoNet471, NoNet472 and NoNet48 switches
            Removed UseWMF3 switch
            Added Ex2013 detection as it cannot coexist with Ex2019CU15+
            Enabled loading Exchange module in postconf needed for possible override cmdlets
            Removed setup phase shown on wallpaper
            Set minimal required PS version to 5.1
            Code cleanup
            Functions now use approved verbs
    4.01    Removed obsolete TLS13 setup detection
    4.10    Added support for Exchange Server SE
    4.11    Fixed feature installation for WS2022/WS2025 Core

    .PARAMETER Organization
    Specifies name of the Exchange organization to create. When omitted, the step
    to prepare Active Directory (PrepareAD) will be skipped.

    .PARAMETER InstallEdge
    Specifies you want to install the Edge server role  (Exchange 2013/2016/2019).

    .PARAMETER EdgeDNSSuffix
    Specifies the DNS suffix you want to use on your EDGE

    .PARAMETER MDBName (optional)
    Specifies name of the initially created database.

    .PARAMETER MDBDBPath (optional)
    Specifies database path of the initially created database. Requires MDBName.

    .PARAMETER MDBLogPath (optional)
    Specifies log path of the initially created database. Requires MDBName.

    .PARAMETER InstallPath (optional)
    Specifies (temporary) location of where to store prerequisites files, log
    files, etc. Default location is C:\Install.

    .PARAMETER NoSetup (optional)
    Specifies you don't want to setup Exchange (prepare/prerequisites only). Note that you
    still need to specify the location of Exchange setup, which is used to determine
    its version and which prerequisites should be installed.

    .PARAMETER SourcePath
    Specifies location of the Exchange installation files (setup.exe) or the location of
    the Exchange installation ISO. This ISO will be mounted during installation.

    .PARAMETER TargetPath
    Specifies the location where to install the Exchange binaries.

    .PARAMETER AutoPilot (switch)
    Specifies you want to automatically restart and logon using Account specified. When
    not specified, you will need to restart, logon and start the script again manually.
    You also need to use the InstallPath parameter when used before, so the script knows where
    to pick up the state file.

    .PARAMETER Credentials
    Specifies credentials to use for automatic logon. Use DOMAIN\User or user@domain. When
    not specified, you will be prompted to enter credentials.

    .PARAMETER IncludeFixes (optional)
    Depending on operating system and detected Exchange version to install, will download
    and install additional recommended Exchange hotfixes.

    .PARAMETER SkipRolesCheck (optional)
    Instructs script not to check for Schema Admin and Enterprise Admin roles.

    .PARAMETER NONET481 (optional)
    Prevents installing .NET Framework 4.8.1 and uses 4.8 when deploying Exchange 2019 CU14+
    on supported Operating Systems (WS2016, WS2019). WS2022 only supports .NET Framework 4.8.1

    .PARAMETER DoNotEnableEP (optional)
    Do not enable Extended Protection on Exchange 2019 CU14+

    .PARAMETER DoNotEnableEP_FEEWS (optional)
    Do not enable Extended Protection on the Front-End EWS virtual directory on Exchange 2019 CU14+

    .PARAMETER DisableSSL3 (optional)
    Disables SSL3 after setup.

    .PARAMETER DisableRC4 (optional)
    Disables RC4 after setup.

    .PARAMETER EnableECC (optional)
    Configures Elliptic Curve Cryptography support after setup.

    .PARAMETER NoCBC (optional)
    Prevents configuring AES256-CBC-encrypted content support after setup.

    .PARAMETER EnableAMSI (optional)
    Configure AMSI body scanning for ECP, EWS, OWA and PowerShell (adjust as necessary in-code)

    .PARAMETER EnableTLS12 (optional)
    Enable or disable TLS12

    .PARAMETER EnableTLS13 (optional)
    Enable or disable TLS13 on WS2022/WS2025 for Exchange 2019 CU15+ (default: enable)

    .PARAMETER Recover
    Runs Exchange setup in RecoverServer mode.

    .PARAMETER SCP (optional)
    Reconfigures Autodiscover Service Connection Point record for this server post-setup, i.e.
    https://autodiscover.contoso.com/autodiscover/autodiscover.xml. If you want to remove
    the record, set it to '-'.

    .PARAMETER Lock (optional)
    Locks system when running script.

    .PARAMETER DiagnosticData (optional)
    Switch determines initial Data Collection mode for deploying Exchange 2019 CU11+ or Exchange 2016.

    .PARAMETER Phase
    Internal Use Only :)

    .EXAMPLE
    $Cred=Get-Credential
    .\Install-Exchange15.ps1 -Organization Fabrikam -InstallMailbox -MDBDBPath C:\MailboxData\MDB1\DB -MDBLogPath C:\MailboxData\MDB1\Log -MDBName MDB1 -InstallPath C:\Install -AutoPilot -Credentials $Cred -SourcePath '\\server\share\Exchange 2019\ExchangeServer2019-x64-cu14' -SCP https://autodiscover.fabrikam.com/autodiscover/autodiscover.xml -Verbose

    .EXAMPLE
    .\Install-Exchange15.ps1 -InstallMailbox -MDBName MDB3 -MDBDBPath C:\MailboxData\MDB3\DB\MDB3.edb -MDBLogPath C:\MailboxData\MDB3\Log -AutoPilot -SourcePath D:\Install\ExchangeServer2019-x64-CU14.ISO -Verbose

    .EXAMPLE
    $Cred=Get-Credential
    .\Install-Exchange15.ps1 -AutoPilot -Credentials $Cred

    .EXAMPLE
    .\Install-Exchange15.ps1 -Recover -Autopilot -Install -AutoPilot -SourcePath \\server1\sources\ex2016cu23

    .EXAMPLE
    .\Install-Exchange15.ps1 -NoSetup -Autopilot -InstallPath \\server1\exfiles\\server1\sources\ex2019cu14

#>
[cmdletbinding(DefaultParameterSetName='AutoPilot')]
param(
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[ValidatePattern(('(?# Organization Name can only consist of upper or lowercase A-Z, 0-9, spaces - not at beginning or end, hyphen or dash characters, up to 64 characters in length, and cannot be empty)^[a-zA-Z0-9\-\–\—][a-zA-Z0-9\-\–\—\ ]{1,62}[a-zA-Z0-9\-\–\—]$'))]
    [string]$Organization,
        [parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
    [switch]$InstallEdge,
        [parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
    [string]$EdgeDNSSuffix,
	[parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [switch]$Recover,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[string]$MDBName,
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[string]$MDBDBPath,
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[string]$MDBLogPath,
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='AutoPilot')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
	[string]$InstallPath= 'C:\Install',
	[parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
 	[parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
        [ValidateScript({ If((Test-Path -Path $_ -PathType Container) -or (Get-DiskImage -ImagePath $_)) { $true } Else { Throw ('Specified source path or image {0} not found or inaccessible' -f $_)} })]
	[string]$SourcePath,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[string]$TargetPath,
	[parameter( Mandatory=$true, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[switch]$NoSetup= $false,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
	[switch]$AutoPilot,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [System.Management.Automation.PsCredential]$Credentials,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$IncludeFixes,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$NoNet481,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$DoNotEnableEP,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$DoNotEnableEP_FEEWS,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$DisableSSL3,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$DisableRC4,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$EnableECC,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$NoCBC,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$EnableAMSI,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$EnableTLS12,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$EnableTLS13,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
    [ValidateScript({ ($_ -eq '') -or ($_ -eq '-') -or (([System.URI]$_).AbsoluteUri -ne $null)})]
    [String]$SCP='',
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$DiagnosticData,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$Lock,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [Switch]$SkipRolesCheck,
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='M')]
 	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='E')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='NoSetup')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='AutoPilot')]
	[parameter( Mandatory=$false, ValueFromPipelineByPropertyName=$false, ParameterSetName='Recover')]
    [ValidateRange(0,6)]
    [int]$Phase
)

process {

    $ScriptVersion                  = '4.11'

    $ERR_OK                         = 0
    $ERR_PROBLEMADPREPARE	    = 1001
    $ERR_UNEXPECTEDOS               = 1002
    $ERR_UNEXPTECTEDPHASE           = 1003
    $ERR_PROBLEMADDINGFEATURE	    = 1004
    $ERR_NOTDOMAINJOINED            = 1005
    $ERR_NOFIXEDIPADDRESS           = 1006
    $ERR_CANTCREATETEMPFOLDER       = 1007
    $ERR_UNKNOWNROLESSPECIFIED      = 1008
    $ERR_NOACCOUNTSPECIFIED         = 1009
    $ERR_RUNNINGNONADMINMODE        = 1010
    $ERR_AUTOPILOTNOSTATEFILE       = 1011
    $ERR_ADMIXEDMODE                = 1012
    $ERR_ADFORESTLEVEL              = 1013
    $ERR_INVALIDCREDENTIALS         = 1014
    $ERR_MDBDBLOGPATH               = 1016
    $ERR_MISSINGORGANIZATIONNAME    = 1017
    $ERR_ORGANIZATIONNAMEMISMATCH   = 1018
    $ERR_RUNNINGNONENTERPRISEADMIN  = 1019
    $ERR_RUNNINGNONSCHEMAADMIN      = 1020
    $ERR_COULDNOTDETERMINEADSITE    = 1021
    $ERR_PROBLEMPACKAGEDL           = 1120
    $ERR_PROBLEMPACKAGESETUP        = 1121
    $ERR_PROBLEMPACKAGEEXTRACT      = 1122
    $ERR_BADFORESTLEVEL             = 1151
    $ERR_BADDOMAINLEVEL             = 1152
    $ERR_MISSINGEXCHANGESETUP       = 1201
    $ERR_PROBLEMEXCHANGESETUP       = 1202
    $ERR_PROBLEMEXCHANGESERVEREXISTS= 1203
    $ERR_EX19EX2013COEXIST          = 1204
    $ERR_UNSUPPORTEDEX              = 1205

    $COUNTDOWN_TIMER                = 10
    $DOMAIN_MIXEDMODE               = 0
    $FOREST_LEVEL2012               = 5
    $FOREST_LEVEL2012R2             = 6

    # Minimum FFL/DFL levels
    $EX2016_MINFORESTLEVEL          = 15317
    $EX2016_MINDOMAINLEVEL          = 13236
    $EX2019_MINFORESTLEVEL          = 17000
    $EX2019_MINDOMAINLEVEL          = 13236

    # Exchange Versions
    $EX2016_MAJOR                   = '15.1'
    $EX2019_MAJOR                   = '15.2'

    # Exchange Install registry key
    $EXCHANGEINSTALLKEY             = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup"

    # Supported Exchange versions (setup.exe)
    $EX2016SETUPEXE_CU23            = '15.01.2507.006'
    $EX2019SETUPEXE_CU10            = '15.02.0922.007'
    $EX2019SETUPEXE_CU11            = '15.02.0986.005'
    $EX2019SETUPEXE_CU12            = '15.02.1118.007'
    $EX2019SETUPEXE_CU13            = '15.02.1258.012'
    $EX2019SETUPEXE_CU14            = '15.02.1544.004'
    $EX2019SETUPEXE_CU15            = '15.02.1748.008'
    $EXSESETUPEXE_RTM               = '15.02.2562.017'

    # Supported Operating Systems
    $WS2016_MAJOR                   = '10.0'
    $WS2019_PREFULL                 = '10.0.17709'
    $WS2022_PREFULL                 = '10.0.20348'
    $WS2025_PREFULL                 = '10.0.20348'

    # .NET Framework Versions
    $NETVERSION_48                  = 528040
    $NETVERSION_481                 = 533320

    # FFL
    $FFL_2003                       = 2
    $FFL_2008                       = 3
    $FFL_2008R2                     = 4
    $FFL_2012                       = 5
    $FFL_2012R2                     = 6
    $FFL_2016                       = 7
    $FFL_2025                       = 10

    Function Save-State( $State) {
        Write-MyVerbose "Saving state information to $StateFile"
        Export-Clixml -InputObject $State -Path $StateFile
    }

    Function Restore-State() {
        $State= @{}
        If(Test-Path $StateFile) {
            $State= Import-Clixml -Path $StateFile -ErrorAction SilentlyContinue
            Write-Verbose "State information loaded from $StateFile"
        }
        Else {
            Write-Verbose "No state file found at $StateFile"
        }
        Return $State
    }


    Function Get-SetupTextVersion( $FileVersion) {
      $Versions= @{
        $EX2016SETUPEXE_CU23= 'Exchange Server 2016 Cumulative Update 23';
        $EX2019SETUPEXE_CU10= 'Exchange Server 2019 CU10';
        $EX2019SETUPEXE_CU11= 'Exchange Server 2019 CU11';
        $EX2019SETUPEXE_CU12= 'Exchange Server 2019 CU12';
        $EX2019SETUPEXE_CU13= 'Exchange Server 2019 CU13';
        $EX2019SETUPEXE_CU14= 'Exchange Server 2019 CU14';
        $EX2019SETUPEXE_CU15= 'Exchange Server 2019 CU15';
        $EXSESETUPEXE_RTM= 'Exchange Server SE RTM';
      }
      $res= "Unsupported version (build $FileVersion)"
      $Versions.GetEnumerator() | Sort-Object -Property {[System.Version]$_.Name} | ForEach-Object {
          If( [System.Version]$FileVersion -ge [System.Version]$_.Name) {
              $res= '{0} (build {1})' -f $_.Value, $FileVersion
          }
      }
      return $res
    }

    Function Get-DetectedFileVersion( $File) {
        $res= 0
        If( Test-Path $File) {
            $res= (Get-Command $File).FileVersionInfo.ProductVersion
        }
        Else {
            $res= 0
        }
        return $res
    }

    Function Write-MyOutput( $Text) {
        Write-Output $Text
        $Location= Split-Path $State['TranscriptFile'] -Parent
        If( Test-Path $Location) {
            Write-Output "$(Get-Date -Format u): $Text" | Out-File $State['TranscriptFile'] -Append -ErrorAction SilentlyContinue
        }
    }

    Function Write-MyWarning( $Text) {
        Write-Warning $Text
        $Location= Split-Path $State['TranscriptFile'] -Parent
        If( Test-Path $Location) {
            Write-Output "$(Get-Date -Format u): [WARNING] $Text" | Out-File $State['TranscriptFile'] -Append -ErrorAction SilentlyContinue
        }
    }

    Function Write-MyError( $Text) {
        Write-Error $Text
        $Location= Split-Path $State['TranscriptFile'] -Parent
        If( Test-Path $Location) {
            Write-Output "$(Get-Date -Format u): [ERROR] $Text" | Out-File $State['TranscriptFile'] -Append -ErrorAction SilentlyContinue
        }
    }

    Function Write-MyVerbose( $Text) {
        Write-Verbose $Text
        $Location= Split-Path $State['TranscriptFile'] -Parent
        If( Test-Path $Location) {
            Write-Output "$(Get-Date -Format u): [VERBOSE] $Text" | Out-File $State['TranscriptFile'] -Append -ErrorAction SilentlyContinue
        }
    }

    Function Get-PSExecutionPolicy {
        $PSPolicyKey= Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell' -Name ExecutionPolicy -ErrorAction SilentlyContinue
        If( $PSPolicyKey) {
            Write-MyWarning "PowerShell Execution Policy is set to $($PSPolicyKey.ExecutionPolicy) through GPO"
        }
        Else {
            Write-MyVerbose 'PowerShell Execution Policy not configured through GPO'
        }
        return $PSPolicyKey
    }

    Function Get-MyPackage () {
        Param ( [String]$Package, [String]$URL, [String]$FileName, [String]$InstallPath)
        $res= $true
        If( !( Test-Path $(Join-Path $InstallPath $Filename))) {
            If( $URL) {
                Write-MyOutput "Package $Package not found, downloading to $FileName"
                Try{
                    Write-MyVerbose "Source: $URL"
                    Start-BitsTransfer -Source $URL -Destination $(Join-Path $InstallPath $Filename)
                }
                Catch{
                    Write-MyError 'Problem downloading package from URL'
                    $res= $false
                }
            }
            Else {
                Write-MyWarning "$FileName not present, skipping"
                $res= $false
            }
        }
        Else {
            Write-MyVerbose "Located $Package ($InstallPath\$FileName)"
        }
        Return $res
    }

    Function Get-CurrentUserName {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    Function Test-Admin {
        $currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
        return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    }

    Function Test-SchemaAdmin {
        $FRNC= Get-ForestRootNC
        $ADRootSID= ([ADSI]"LDAP://$FRNC").ObjectSID[0]
        $SID= (New-object System.Security.Principal.SecurityIdentifier ($ADRootSID, 0)).Value.toString()
        return [Security.Principal.WindowsIdentity]::GetCurrent().Groups | Where-Object {$_.Value -eq "$SID-518"}
    }

    Function Test-EnterpriseAdmin {
        $FRNC= Get-ForestRootNC
        $ADRootSID= ([ADSI]"LDAP://$FRNC").ObjectSID[0]
        $SID= (New-object System.Security.Principal.SecurityIdentifier ($ADRootSID, 0)).Value.toString()
        return [Security.Principal.WindowsIdentity]::GetCurrent().Groups | Where-Object {$_.Value -eq "$SID-519"}
    }

    Function Test-ServerCore {
        (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion' -Name 'InstallationType' -ErrorAction SilentlyContinue).InstallationType -eq 'Server Core'
    }

    Function Test-RebootPending {
        $Pending= $False
        If( Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue) {
            $Pending= $True
        }
        If( Test-Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ErrorAction SilentlyContinue) {
            $Pending= $True
        }
        Return $Pending
    }

    Function Enable-RunOnce {
        Write-MyOutput 'Set script to run once after reboot'
        $RunOnce= "$PSHome\powershell.exe -NoProfile -ExecutionPolicy Unrestricted -Command `"& `'$ScriptFullName`' -InstallPath `'$InstallPath`'`""
        Write-MyVerbose "RunOnce: $RunOnce"
        New-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce' -Name "$ScriptName"  -Value "$RunOnce" -ErrorAction SilentlyContinue| out-null
    }

    Function Disable-UAC {
        Write-MyVerbose 'Disabling User Account Control'
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 -ErrorAction SilentlyContinue| out-null
    }

    Function Enable-UAC {
        Write-MyVerbose 'Enabling User Account Control'
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 1 -ErrorAction SilentlyContinue| out-null
    }

    Function Disable-IEESC {
        Write-MyOutput 'Disabling IE Enhanced Security Configuration'
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | out-null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 0 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | out-null
        Set-ItemProperty -Path $UserKey -Name 'IsInstalled' -Value 0 -Force | Out-Null
        If( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    Function Enable-IEESC {
        Write-MyVerbose 'Enabling IE Enhanced Security Configuration'
        $AdminKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
        $UserKey = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'
        New-Item -Path (Split-Path $AdminKey -Parent) -Name (Split-Path $AdminKey -Leaf) -ErrorAction SilentlyContinue | out-null
        Set-ItemProperty -Path $AdminKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        New-Item -Path (Split-Path $UserKey -Parent) -Name (Split-Path $UserKey -Leaf) -ErrorAction SilentlyContinue | out-null
        Set-ItemProperty -Path $UserKey -Name 'IsInstalled' -Value 1 -Force | Out-Null
        If( Get-Process -Name explorer.exe -ErrorAction SilentlyContinue) {
            Stop-Process -Name Explorer
        }
    }

    Function get-FullDomainAccount {
        $PlainTextAccount= $State['AdminAccount']
        If( $PlainTextAccount.indexOf('\') -gt 0) {
            $Parts= $PlainTextAccount.split('\')
            $Domain = $Parts[0]
            $UserName= $Parts[1]
            Return "$Domain\$UserName"
        } Else {
            If( $PlainTextAccount.indexOf('@') -gt 0) {
                Return $PlainTextAccount
            }
            Else {
                $Domain = $env:USERDOMAIN
                $UserName= $PlainTextAccount
                Return "$Domain\$UserName"
            }
        }
    }

    #From https://gallery.technet.microsoft.com/scriptcenter/Verify-the-Local-User-1e365545
    function Test-LocalCredential {
        [CmdletBinding()]
        Param
        (
            [string]$UserName,
            [string]$ComputerName = $env:COMPUTERNAME,
            [string]$Password
        )
        if (!($UserName) -or !($Password)) {
            Write-Warning 'Test-LocalCredential: Please specify both user name and password'
        } else {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
            $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('machine',$ComputerName)
            $DS.ValidateCredentials($UserName, $Password)
        }
    }

    Function Test-Credentials {
        $PlainTextPassword= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( (ConvertTo-SecureString $State['AdminPassword']) ))
        $FullPlainTextAccount= get-FullDomainAccount
	    Try {
            If( $State['InstallEdge']) {
                $Username = $FullPlainTextAccount.split("\")[-1]
                Return $( Test-LocalCredential -UserName $Username -Password $PlainTextPassword)
            }else{
                $dc= New-Object DirectoryServices.DirectoryEntry( $Null, $FullPlainTextAccount, $PlainTextPassword)
                If($dc.Name) {
                    return $true
                }
                Else {
                    Return $false
                }
            }

	    }
	    Catch {
		    Return $false
	    }
	    Return $false
    }

    Function Enable-AutoLogon {
        Write-MyVerbose 'Enabling Automatic Logon'
        $PlainTextPassword= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( (ConvertTo-SecureString $State['AdminPassword']) ))
        $PlainTextAccount= $State['AdminAccount']
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 -ErrorAction SilentlyContinue| out-null
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $PlainTextAccount -ErrorAction SilentlyContinue| out-null
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $PlainTextPassword -ErrorAction SilentlyContinue| out-null
    }

    Function Disable-AutoLogon {
        Write-MyVerbose 'Disabling Automatic Logon'
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -ErrorAction SilentlyContinue| out-null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -ErrorAction SilentlyContinue| out-null
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -ErrorAction SilentlyContinue| out-null
    }

    Function Disable-OpenFileSecurityWarning {
        Write-MyVerbose 'Disabling File Security Warning dialog'
        New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -ErrorAction SilentlyContinue |out-null
        New-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -name 'LowRiskFileTypes' -value '.exe;.msp;.msu;.msi' -ErrorAction SilentlyContinue |out-null
        New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -ErrorAction SilentlyContinue |out-null
        New-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -name 'SaveZoneInformation' -value 1 -ErrorAction SilentlyContinue |out-null
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    Function Enable-OpenFileSecurityWarning {
        Write-MyVerbose 'Enabling File Security Warning dialog'
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Associations' -Name 'LowRiskFileTypes' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments' -Name 'SaveZoneInformation' -ErrorAction SilentlyContinue
    }

    Function Invoke-Extract ( $FilePath, $FileName) {
        Write-MyVerbose "Extracting $FilePath\$FileName to $FilePath"
        If( Test-Path $(Join-Path $FilePath $Filename)) {
            $TempNam= "$(Join-Path $FilePath $Filename).zip"
            Copy-Item $(Join-Path $FilePath $Filename) "$TempNam" -Force
            $shellApplication = new-object -com shell.application
            $zipPackage = $shellApplication.NameSpace( $TempNam)
            $destFolder = $shellApplication.NameSpace( $FilePath)
            $destFolder.CopyHere( $zipPackage.Items(), 0x10)
            Remove-Item $TempNam
        }
        Else {
            Write-MyWarning "$FilePath\$FileName not found"
        }
    }

    Function Invoke-Process ( $FilePath, $FileName, $ArgumentList) {
        $rval= 0
        $FullName= Join-Path $FilePath $FileName
        If( Test-Path $FullName) {
            Switch( ([io.fileinfo]$Filename).extension.ToUpper()) {
                '.MSU' {
                    $ArgumentList+= @( $FullName)
                    $ArgumentList+= @( '/f')
                    $Cmd= "$env:SystemRoot\System32\WUSA.EXE"
                }
                '.MSI' {
                    $ArgumentList+= @( '/i')
                    $ArgumentList+= @( $FullName)
                    $Cmd= "MSIEXEC.EXE"
                }
                '.MSP' {
                    $ArgumentList+= @( '/update')
                    $ArgumentList+= @( $FullName)
                    $Cmd= 'MSIEXEC.EXE'
                }
                default {
                    $Cmd= $FullName
                }
            }
          Write-MyVerbose "Executing $Cmd $($ArgumentList -Join ' ')"
          $rval=( Start-Process -FilePath $Cmd -ArgumentList $ArgumentList -NoNewWindow -PassThru -Wait).Exitcode
          Write-MyVerbose "Process exited with code $rval"
        }
        Else {
            Write-MyWarning "$FullName not found"
            $rval= -1
        }
        return $rval
    }
    Function Get-ForestRootNC {
        return ([ADSI]'LDAP://RootDSE').rootDomainNamingContext.toString()
    }
    Function Get-RootNC {
        return ([ADSI]'').distinguishedName.toString()
    }

    Function Get-ForestConfigurationNC {
        return ([ADSI]'LDAP://RootDSE').configurationNamingContext.toString()
    }

    Function Get-ForestFunctionalLevel {
        $CNC= Get-ForestConfigurationNC
        Try {
            $rval= ( ([ADSI]"LDAP://cn=partitions,$CNC").get('msDS-Behavior-Version') )
        }
        Catch {
            Write-MyError "Can't read Forest schema version, operator possibly not member of Schema Admin group"
        }
        return $rval
    }

    Function Test-DomainNativeMode {
        $NC= Get-RootNC
        return( ([ADSI]"LDAP://$NC").ntMixedDomain )
    }

    Function Get-ExchangeOrganization {
        $CNC= Get-ForestConfigurationNC
        Try {
            $ExOrgContainer= [ADSI]"LDAP://CN=Microsoft Exchange,CN=Services,$CNC"
            $rval= ($ExOrgContainer.PSBase.Children | Where-Object { $_.objectClass -eq 'msExchOrganizationContainer' }).Name
        }
        Catch {
            Write-MyVerbose "Can't find Exchange Organization object"
            $rval= $null
        }
        return $rval
    }

    Function Test-ExchangeOrganization( $Organization) {
        $CNC= Get-ForestConfigurationNC
        return( [ADSI]"LDAP://CN=$Organization,CN=Microsoft Exchange,CN=Services,$CNC")
    }

    Function Get-ExchangeForestLevel {
        $CNC= Get-ForestConfigurationNC
        return ( ([ADSI]"LDAP://CN=ms-Exch-Schema-Version-Pt,CN=Schema,$CNC").rangeUpper )
    }

    Function Get-ExchangeDomainLevel {
        $NC= Get-RootNC
        return( ([ADSI]"LDAP://CN=Microsoft Exchange System Objects,$NC").objectVersion )
    }

    Function Clear-AutodiscoverServiceConnectionPoint( [string]$Name) {
        $CNC= Get-ForestConfigurationNC
        $LDAPSearch= New-Object System.DirectoryServices.DirectorySearcher
        $LDAPSearch.SearchRoot= "LDAP://$CNC"
        $LDAPSearch.Filter= "(&(cn=$Name)(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))"
        $LDAPSearch.FindAll() | ForEach-Object {
            Write-MyVerbose "Removing object $($_.Path)"
            ([ADSI]($_.Path)).DeleteTree()
        }
    }

   Function Set-AutodiscoverServiceConnectionPoint( [string]$Name, [string]$ServiceBinding) {
        $CNC= Get-ForestConfigurationNC
        $LDAPSearch= New-Object System.DirectoryServices.DirectorySearcher
        $LDAPSearch.SearchRoot= "LDAP://$CNC"
        $LDAPSearch.Filter= "(&(cn=$Name)(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))"
        $LDAPSearch.FindAll() | ForEach-Object {
            Write-MyVerbose "Setting serviceBindingInformation on $($_.Path) to $ServiceBinding"
            Try {
                $SCPObj= $_.GetDirectoryEntry()
                $null = $SCPObj.Put( 'serviceBindingInformation', $ServiceBinding)
                $SCPObj.SetInfo()
            }
            Catch {
                Write-MyError "Problem setting serviceBindingInformation property: $($Error[0])"
            }
        }
    }

    Function Test-ExistingExchangeServer( [string]$Name) {
        $CNC= Get-ForestConfigurationNC
        $LDAPSearch= New-Object System.DirectoryServices.DirectorySearcher
        $LDAPSearch.SearchRoot = "LDAP://$CNC"
        $LDAPSearch.Filter = "(&(cn=$Name)(objectClass=msExchExchangeServer))"
        $Results = $LDAPSearch.FindAll()
        Return ($Results.Count -gt 0)
    }

    Function Get-LocalFQDNHostname {
        return ([System.Net.Dns]::GetHostByName(($env:computerName))).HostName
    }

    Function Get-ADSite {
        Try {
            return [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()
        }
        Catch {
            Return $null
        }
    }

    Function Get-ExchangeServerObjects {
        $CNC= Get-ForestConfigurationNC
        Get-ADObject -Filter "ObjectCategory -eq 'msExchExchangeServer'" -SearchBase $CNC -Properties msExchCurrentServerRoles, networkAddress, serialNumber
    }

    Function Set-EdgeDNSSuffix ([string]$DNSSuffix){
        Write-MyVerbose 'Setting Primary DNS Suffix'
        #https://technet.microsoft.com/library%28EXCHG.150%29/ms.exch.setupreadiness.FqdnMissing.aspx?f=255&MSPPError=-2147217396
        #Update primary DNS Suffix for FQDN
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\" -Name Domain -Value $DNSSuffix
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\" -Name "NV Domain" -Value $DNSSuffix

    }

    Function Import-ExchangeModule {
        Write-MyVerbose 'Loading Exchange PowerShell module'
        If( -not ( Get-Command Connect-ExchangeServer -ErrorAction SilentlyContinue)) {
            $SetupPath= (Get-ItemProperty -Path $EXCHANGEINSTALLKEY -Name MsiInstallPath -ErrorAction SilentlyContinue).MsiInstallPath
            If( ($State['InstallEdge'] -eq $true -and $SetupPath -and (Test-Path $(Join-Path $SetupPath "\bin\Exchange.ps1"))) -or ($State['InstallEdge'] -eq $false -and $SetupPath -and (Test-Path $(Join-Path $SetupPath "\bin\RemoteExchange.ps1")))) {
                If( $State['InstallEdge']) {
                    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
                    . "$SetupPath\bin\Exchange.ps1" | Out-Null
                }else{
                    . "$SetupPath\bin\RemoteExchange.ps1" | Out-Null
                    Try {
                        Connect-ExchangeServer (Get-LocalFQDNHostname)
                    }
                    Catch {
                        Write-MyError 'Problem loading Exchange module'
                    }
                }


            }
            Else {
                Write-MyWarning "Can't determine installation path to load Exchange module"
            }
        }
        Else {
            Write-MyWarning 'Exchange module already loaded'
        }
    }

    Function Install-Exchange15_ {
        $ver= $State['MajorSetupVersion']
        Write-MyOutput "Installing Microsoft Exchange Server ($ver)"
        $PresenceKey= 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{CD981244-E9B8-405A-9026-6AEB9DCEF1F1}'

        If( $State['Recover']) {
            Write-MyOutput 'Wil run Setup in recover mode'
            $Params= '/mode:RecoverServer', $State['IAcceptSwitch'], '/DoNotStartTransport', '/InstallWindowsComponents'
            If( $State['TargetPath']) {
                $Params+= "/TargetDir:`"$($State['TargetPath'])`""
            }
        }
        Else {
            If( $State['Upgrade']) {
                Write-MyOutput 'Wil run Setup in upgrade mode'
                $Params= '/mode:Upgrade', $State['IAcceptSwitch']
            }
            Else {
                $roles= @()
                If( $State['InstallEdge']) {
                    $roles = 'EdgeTransport'
                }
                Else
                {
                    $roles= 'Mailbox'
                }
	        $RolesParm= $roles -Join ','
                If([string]::IsNullOrEmpty( $RolesParam)) {
                    $RolesParam= 'Mailbox'
                }
                $Params= '/mode:install', "/roles:$RolesParm", $State['IAcceptSwitch'], '/DoNotStartTransport', '/InstallWindowsComponents'
                If( $State['InstallMailbox']) {
                    If( $State['InstallMDBName']) {
                        $Params+= "/MdbName:$($State['InstallMDBName'])"
                    }
                    If( $State['InstallMDBDBPath']) {
                        $Params+= "/DBFilePath:`"$($State['InstallMDBDBPath'])\$($State['InstallMDBName']).edb`""
                    }
                    If( $State['InstallMDBLogPath']) {
                        $Params+= "/LogFolderPath:`"$($State['InstallMDBLogPath'])`""
                    }
                }
                If( $State['TargetPath']) {
                    $Params+= "/TargetDir:`"$($State['TargetPath'])`""
                }
                If( $State['DoNotEnableEP']) {
                    $Params+= "/DoNotEnableEP"
                }
                If( $State['DoNotEnableEP_FEEWS']) {
                    $Params+= "/DoNotEnableEP_FEEWS"
                }
            }
        }

        $res= Invoke-Process $State['SourcePath'] 'setup.exe' $Params
        If( $res -ne 0 -or -not( Get-ItemProperty -Path $PresenceKey -Name InstallDate -ErrorAction SilentlyContinue)){
            Write-MyError 'Exchange Setup exited with non-zero value or Install info missing from registry: Please consult the Exchange setup log, i.e. C:\ExchangeSetupLogs\ExchangeSetup.log'
            Exit $ERR_PROBLEMEXCHANGESETUP
        }
    }

    Function Initialize-Exchange {
        If(!$State['InstallEdge']) {
            $params= @()
            Write-MyOutput 'Checking Exchange organization existence'
            If( $null -ne ( Test-ExchangeOrganization $State['OrganizationName'])) {
                $params+= '/PrepareAD', "/OrganizationName:`"$($State['OrganizationName'])`""
            }
            Else {
                Write-MyOutput 'Organization exist; checking Exchange Forest Schema and Domain versions'
                $forestlvl= Get-ExchangeForestLevel
                $domainlvl= Get-ExchangeDomainLevel
                Write-MyOutput "Exchange Forest Schema version: $forestlvl, Domain: $domainlvl)"
                $MinFFL= $EX2016_MINFORESTLEVEL
                $MinDFL= $EX2016_MINDOMAINLEVEL
                If(( $forestlvl -lt $MinFFL) -or ( $domainlvl -lt $MinDFL)) {
                    Write-MyOutput "Exchange Forest Schema or Domain needs updating (Required: $MinFFL/$MinDFL)"
                    $params+= '/PrepareAD'

                }
                Else {
                    Write-MyOutput 'Active Directory looks already updated'
                }
            }
        }
        If ($params.count -gt 0) {
            If(!$State['InstallEdge']) {
                Write-MyOutput "Preparing AD, Exchange organization will be $($State['OrganizationName'])"¨
            }
            $params+= $State['IAcceptSwitch']
            Invoke-Process $State['SourcePath'] 'setup.exe' $params
            If( ( $null -eq ( Test-ExchangeOrganization $State['OrganizationName'])) -or
                ( (Get-ExchangeForestLevel) -lt $MinFFL) -or
                ( (Get-ExchangeDomainLevel) -lt $MinDFL)) {
                Write-MyError 'Problem updating schema, domain or Exchange organization. Please consult the Exchange setup log, i.e. C:\ExchangeSetupLogs\ExchangeSetup.log'
                Exit $ERR_PROBLEMADPREPARE
            }
        }
        Else {
            Write-MyWarning "Exchange organization $($State['OrganizationName']) already exists, skipping this step"
        }
    }

    Function Install-WindowsFeatures( $MajorOSVersion) {
        Write-MyOutput 'Configuring Windows Features'

        If( $State['InstallEdge']) {
            $Feats= 'ADLDS'
        }
        Else {
            If( [System.Version]$WS2019_PREFULL -ge [System.Version]$MajorOSVersion) {

                # WS2019, WS2022, WS2025
                $Feats= 'Server-Media-Foundation', 'NET-Framework-45-Core', 'NET-Framework-45-ASPNET',
                    'NET-WCF-HTTP-Activation45', 'NET-WCF-Pipe-Activation45', 'NET-WCF-TCP-Activation45',
                    'NET-WCF-TCP-PortSharing45', 'RPC-over-HTTP-proxy', 'RSAT-Clustering',
                    'RSAT-Clustering-CmdInterface', 'RSAT-Clustering-PowerShell', 'WAS-Process-Model',
                    'Web-Asp-Net45', 'Web-Basic-Auth', 'Web-Client-Auth', 'Web-Digest-Auth',
                    'Web-Dir-Browsing', 'Web-Dyn-Compression', 'Web-Http-Errors', 'Web-Http-Logging',
                    'Web-Http-Redirect', 'Web-Http-Tracing', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter',
                    'Web-Metabase', 'Web-Mgmt-Service', 'Web-Net-Ext45', 'Web-Request-Monitor',
                    'Web-Server', 'Web-Stat-Compression', 'Web-Static-Content', 'Web-W-Auth',
                    'Web-WMI', 'RSAT-ADDS'

                If( !( Test-ServerCore)) {
                    $Feats+= 'RSAT-Clustering-Mgmt', 'Web-Mgmt-Console', 'Windows-Identity-Foundation'
                }
            }
            Else {
                # WS2016
                $Feats= 'NET-Framework-45-Core', 'NET-Framework-45-ASPNET', 'NET-WCF-HTTP-Activation45', 'NET-WCF-Pipe-Activation45', 'NET-WCF-TCP-Activation45', 'NET-WCF-TCP-PortSharing45', 'Server-Media-Foundation', 'RPC-over-HTTP-proxy', 'RSAT-Clustering', 'RSAT-Clustering-CmdInterface', 'RSAT-Clustering-Mgmt', 'RSAT-Clustering-PowerShell', 'WAS-Process-Model', 'Web-Asp-Net45', 'Web-Basic-Auth', 'Web-Client-Auth', 'Web-Digest-Auth', 'Web-Dir-Browsing', 'Web-Dyn-Compression', 'Web-Http-Errors', 'Web-Http-Logging', 'Web-Http-Redirect', 'Web-Http-Tracing', 'Web-ISAPI-Ext', 'Web-ISAPI-Filter', 'Web-Lgcy-Mgmt-Console', 'Web-Metabase', 'Web-Mgmt-Console', 'Web-Mgmt-Service', 'Web-Net-Ext45', 'Web-Request-Monitor', 'Web-Server', 'Web-Stat-Compression', 'Web-Static-Content', 'Web-Windows-Auth', 'Web-WMI', 'Windows-Identity-Foundation', 'RSAT-ADDS'
            }
        }
        $Feats+= 'Bits'

        Install-WindowsFeature $Feats | out-null

        ForEach( $Feat in $Feats) {
            If( !( Get-WindowsFeature ($Feat))) {
                Write-MyError "Feature $Feat appears not to be installed"
                Exit $ERR_PROBLEMADDINGFEATURE
            }
        }

        'NET-WCF-MSMQ-Activation45', 'MSMQ' | ForEach-Object {
            If( Get-WindowsFeature -Name $_) {
                Write-MyOutput ('Removing obsolete feature {0}' -f $_)
                Remove-WindowsFeature -Name $_
            }
        }
    }

    Function Test-MyPackage( $PackageID) {
        # Some packages are released using different GUIDs, specify more than 1 using '|'
        $PackageSet= $PackageID.split('|')
        $PresenceKey= $null
        ForEach( $ID in $PackageSet) {
            Write-MyVerbose "Checking if package $ID is installed .."
            $PresenceKey= (Get-WmiObject win32_quickfixengineering | Where-Object { $_.HotfixID -eq $ID }).HotfixID
            If( !( $PresenceKey)) {
                $PresenceKey= (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                If(!( $PresenceKey)) {
                    # Alternative (seen KB2803754, 2802063 register here)
                    $PresenceKey= (Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                    If( !( $PresenceKey)){
                        # Alternative (eg Office2010FilterPack SP1)
                        $PresenceKey= (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$ID" -Name 'DisplayName' -ErrorAction SilentlyContinue).DisplayName
                        If( !( $PresenceKey)){
                            # Check for installed Exchange IUs
                            Switch( $State["MajorSetupVersion"]) {
                                $EX2016_MAJOR {
                                    $IUPath= 'Exchange 2016'
                                }
                                default {
                                    $IUPath= 'Exchange 2019'
                                }
                            }
                            $PresenceKey= (Get-ItemProperty -Path ('HKLM:\Software\Microsoft\Updates\{0}\{1}' -f $IUPath, $ID) -Name 'PackageName' -ErrorAction SilentlyContinue).PackageName
                        }
                    }
                }
            }
        }
        return $PresenceKey
    }

    Function Install-MyPackage {
        Param ( [String]$PackageID, [string]$Package, [String]$FileName, [String]$OnlineURL, [array]$Arguments, [switch]$NoDownload)

        If( $PackageID) {
            Write-MyOutput "Processing $Package ($PackageID)"
            $PresenceKey= Test-MyPackage $PackageID
        }
        Else {
            # Just install, don't detect
            Write-MyOutput "Processing $Package"
            $PresenceKey= $false
        }
        $RunFrom= $State['InstallPath']
        If( !( $PresenceKey )){

            If( $FileName.contains('|')) {
                # Filename contains filename (dl) and package name (after extraction)
                $PackageFile= ($FileName.Split('|'))[1]
                $FileName= ($FileName.Split('|'))[0]
                If( !( Get-MyPackage $Package '' $FileName $RunFrom)) {
                    # Download & Extract
                    If( !( Get-MyPackage $Package $OnlineURL $PackageFile $RunFrom)) {
                        Write-MyError "Problem downloading/accessing $Package"
                        Exit $ERR_PROBLEMPACKAGEDL
                    }
                    Write-MyOutput "Extracting Hotfix Package $Package"
                    Invoke-Extract $RunFrom $PackageFile

                    If( !( Get-MyPackage $Package $OnlineURL $PackageFile $RunFrom)) {
                        Write-MyError "Problem downloading/accessing $Package"
                        Exit $ERR_PROBLEMPACKAGEEXTRACT
                    }
                }
            }
            Else {
                If( $NoDownload) {
                    $RunFrom= Split-Path -Path $OnlineURL -Parent
                    Write-MyVerbose "Will run $FileName straight from $RunFrom"
                }
                If( !( Get-MyPackage $Package $OnlineURL $FileName $RunFrom)) {
                    Write-MyError "Problem downloading/accessing $Package"
                    Exit $ERR_PROBLEMPACKAGEDL
                }
            }

            Write-MyOutput "Installing $Package from $RunFrom"
            $rval= Invoke-Process $RunFrom $FileName $Arguments

            If( $PackageID) {
                $PresenceKey= Test-MyPackage $PackageID
            }
            Else {
                # Don't check post-installation
                $PresenceKey= $true
            }
            If( ( @(3010,-2145124329) -contains $rval) -or $PresenceKey)  {
                switch ( $rval) {
                    3010: {
                        Write-MyVerbose "Installation $Package successful, reboot required"
                    }
                    -2145124329: {
                        Write-MyVerbose "$Package not applicable or blocked - ignoring"
                    }
                    default: {
                         Write-MyVerbose "Installation $Package successful"
                    }
                }
            }
            Else {
                Write-MyError "Problem installing $Package - For fixes, check $($ENV:WINDIR)\WindowsUpdate.log; For .NET Framework issues, check 'Microsoft .NET Framework 4 Setup' HTML document in $($ENV:TEMP)"
                Exit $ERR_PROBLEMPACKAGESETUP
            }
        }
        Else {
            Write-MyVerbose "$Package already installed"
        }
    }

    Function DisableSharedCacheServiceProbe {
        # Taken from DisableSharedCacheServiceProbe.ps1
        # Copyright (c) Microsoft Corporation. All rights reserved.
        Write-MyOutput "Applying DisableSharedCacheServiceProbe (KB2971467, 'Shared Cache Service Restart' Probe Fix)"
        $exchangeInstallPath = get-itemproperty -path $EXCHANGEINSTALLKEY -ErrorAction SilentlyContinue
        if ($null -ne $exchangeInstallPath -and (Test-Path $exchangeInstallPath.MsiInstallPath)) {
            $ProbeConfigFile= Join-Path ( $exchangeInstallPath.MsiInstallPath) ('Bin\Monitoring\Config\SharedCacheServiceTest.xml')
	        if (Test-Path $ProbeConfigFile) {
	            $date = get-date -format s
	            $ext = '.orig_' + $date.Replace(':', '-');
	            $backup = $ProbeConfigFile + $ext
	            $xmlBackup = [XML](Get-Content $ProbeConfigFile);
	            $xmlBackup.Save($backup);

	            $xmlDoc = [XML](Get-Content $ProbeConfigFile);
	            $definition = $xmlDoc.Definition.MaintenanceDefinition;

	            if($null -eq $definition) {
                    Write-MyError 'KB2971467: Expected XML node Definition.MaintenanceDefinition.ExtensionAttributes not found. Skipping.'
                }
                Else {
                    $modified = $false;
                    if ($null -ne $definition.Enabled -and $definition.Enabled -ne 'false') {
                        $definition.Enabled = 'false';
                        $modified = $true;
                    }
	                If($modified) {
                        $xmlDoc.Save($ProbeConfigFile);
                        Write-MyOutput "Finished KB2971467, Saved $ProbeConfigFile"
                    }
                    Else {
                        Write-MyOutput 'Finished KB2971467, No values modified.'
                    }
                }
            }
            Else {
	            Write-MyError "KB2971467: Did not find file in expected location, skipping $ProbeConfigFile"
	        }
        }
        Else {
            Write-MyError 'KB2971467: Unable to locate Exchange install path'
        }
    }

    Function Get-FFLText( $FFL= 0) {
        $FFLlevels= @{
            0='Unknown or unsupported'
            $FFL_2003='2003'
            $FFL_2008='2008'
            $FFL_2008R2='2008R2'
            $FFL_2012='2012'
            $FFL_2012R2='2012R2'
            $FFL_2016='2016'
            $FFL_2025='2025'
        }
        return ($FFLlevels.GetEnumerator() | Where-Object {$FFL -ge $_.Name} | Sort-Object Name -Descending | Select-Object -First 1).Value
    }

    Function Get-NetVersionText( $NetVersion= 0) {
        $NETversions= @{
            0='Unknown or unsupported'
            $NETVERSION_48='4.8'
            $NETVERSION_481='4.8.1'
        }
        return ($NetVersions.GetEnumerator() | Where-Object {$NetVersion -ge $_.Name} | Sort-Object Name -Descending | Select-Object -First 1).Value
    }

    Function Get-NETVersion {
        $NetVersion= (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
        return [int]$NetVersion
    }

    Function Set-NETFrameworkInstallBlock {
        Param ( [String]$Version, [String]$KB, [string]$Key)
        $RegKey= 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\WU'
        $RegName= ('BlockNetFramework{0}' -f $Key)
        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            If( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                Write-MyOutput ('Set installation blockade for .NET Framework {0} ({1})' -f $Version, $KB)
                New-Item -Path (Split-Path $RegKey -Parent) -Name (Split-Path $RegKey -Leaf) -ErrorAction SilentlyContinue | out-null
            }
        }
        New-ItemProperty -Path $RegKey -Name $RegName  -Value 1 -ErrorAction SilentlyContinue| out-null
        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            Write-MyError "Unable to set registry entry $RegKey\$RegName"
        }
    }

    Function Remove-NETFrameworkInstallBlock {
        Param ( [String]$Version, [String]$KB, [string]$Key)
        $RegKey= 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\WU'
        $RegName= ('BlockNetFramework{0}' -f $Key)
        If( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue) {
            Write-MyOutput ('Remove installation blockade for .NET Framework {0} ({1})' -f $Version, $KB)
            Remove-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue | out-null
        }
        If( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue) {
            Write-MyError "Unable to set registry entry $RegKey\$RegName"
        }
    }

    Function Test-Preflight {
        Write-MyOutput 'Performing preflight checks'

        $Computer= Get-LocalFQDNHostname
        If( $Computer) {
            Write-MyOutput "Computer name is $Computer"
        }

        Write-MyOutput 'Checking temporary installation folder'
        mkdir $State['InstallPath'] -ErrorAction SilentlyContinue |out-null
        If( !( Test-Path $State['InstallPath'])) {
            Write-MyError "Can't create temporary folder $($State['InstallPath'])"
            Exit $ERR_CANTCREATETEMPFOLDER
        }

        If( [System.Version]$MajorOSVersion -ge [System.Version]$WS2016_MAJOR ) {
            Write-MyOutput "Operating System is $($MajorOSVersion).$($MinorOSVersion)"
        }
        Else {
            Write-MyError 'The following Operating Systems are supported: Windows Server 2019, Windows Server 2022 (Exchange 2019) or Windows Server 2025 (Exchange 2019 CU15+)'
            Exit $ERR_UNEXPECTEDOS
        }
        Write-MyOutput ('Server core mode: {0}' -f (Test-ServerCore))

        $NetVersion= Get-NETVersion
        $NetVersionText= Get-NetVersionText $NetVersion
        Write-MyOutput ".NET Framework is $NetVersion ($NetVersionText)"

        If(! ( Test-Admin)) {
            Write-MyError 'Script requires running in elevated mode'
            Exit $ERR_RUNNINGNONADMINMODE
        }
        Else {
            Write-MyOutput 'Script running in elevated mode'
        }

        If( $State['AutoPilot']) {
            If( -not( $State['AdminAccount'] -and $State['AdminPassword'])) {
                Try {
		    $Script:Credentials= Get-Credential -UserName ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -Message 'Enter credentials to use'
                    $State['AdminAccount']= $Credentials.UserName
                    $State['AdminPassword']= ($Credentials.Password | ConvertFrom-SecureString)
                }
                Catch {
                    Write-MyError 'AutoPilot specified but no or improper credentials provided'
                    Exit $ERR_NOACCOUNTSPECIFIED
                }
	        }
            Write-MyOutput 'Checking provided credentials'
            If( Test-Credentials) {
                Write-MyOutput 'Credentials seem valid'
            }
            Else {
                Write-MyError "Provided credentials don't seem to be valid"
                Exit $ERR_INVALIDCREDENTIALS
            }
        }

	If( $State["SkipRolesCheck"] -or $State['InstallEdge']) {
            Write-MyOutput 'SkipRolesCheck: Skipping validation of Schema & Enterprise Administrators membership'
        }
        Else {
            If(! ( Test-SchemaAdmin)) {
                Write-MyError 'Current user is not member of Schema Administrators'
                Exit $ERR_RUNNINGNONSCHEMAADMIN
            }
            Else {
                Write-MyOutput 'User is member of Schema Administrators'
            }

            If(! ( Test-EnterpriseAdmin)) {
                Write-MyError 'User is not member of Enterprise Administrators'
                Exit $ERR_RUNNINGNONENTERPRISEADMIN
            }
            Else {
                Write-MyOutput 'User is member of Enterprise Administrators'
            }
        }
        if(!$State['InstallEdge']){
            $ADSite= Get-ADSite
            If( $ADSite) {
                Write-MyOutput "Computer is located in AD site $ADSite"
            }
            Else {
                Write-MyError 'Could not determine Active Directory site'
                Exit $ERR_COULDNOTDETERMINEADSITE
            }

            $ExOrg= Get-ExchangeOrganization
            If( $ExOrg) {
                If( $State['OrganizationName']) {
                    If( $State['OrganizationName'] -ne $ExOrg) {
                        Write-MyError "OrganizationName mismatches with discovered Exchange Organization name ($ExOrg, expected $($State['OrganizationName']))"
                        Exit $ERR_ORGANIZATIONNAMEMISMATCH
                    }
                }
                Write-MyOutput "Exchange Organization is: $ExOrg"
            }
            Else {
                If( $State['OrganizationName']) {
                    Write-MyOutput "Exchange Organization will be: $($State['OrganizationName'])"
                }
                Else {
                    Write-MyError 'OrganizationName not specified and no Exchange Organization discovered'
                    Exit $ERR_MISSINGORGANIZATIONNAME
                }
            }
        }
        Write-MyOutput 'Checking if we can access Exchange setup ..'

        If(! (Test-Path (Join-Path $State['SourcePath'] "setup.exe"))) {
            Write-MyError "Can't find Exchange setup at $($State['SourcePath'])"
            Exit $ERR_MISSINGEXCHANGESETUP
        }
        Else {
            Write-MyOutput "Exchange setup located at $(Join-Path $($State['SourcePath']) "setup.exe")"
        }

        $State['ExSetupVersion']= Get-DetectedFileVersion "$($State['SourcePath'])\Setup\ServerRoles\Common\ExSetup.exe"
        $SetupVersion= $State['ExSetupVersion']
	    $State['SetupVersionText']= Get-SetupTextVersion $SetupVersion
            Write-MyOutput ('ExSetup version: {0}' -f $State['SetupVersionText'])
            If( $SetupVersion) {
                $Num= $SetupVersion.split('.') | ForEach-Object { [string]([int]$_)
            }
            $MajorSetupVersion= [decimal]($num[0]+ '.'+ $num[1])
            $MinorSetupVersion= [decimal]($num[2]+ '.'+ $num[3])
        }
        Else {
            $MajorSetupVersion= 0
            $MinorSetupVersion= 0
        }
        $State['MajorSetupVersion'] = $MajorSetupVersion
        $State['MinorSetupVersion'] = $MinorSetupVersion

        If( ($MajorSetupVersion -eq $EX2019_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2019SETUPEXE_CU10) -or
            ($MajorSetupVersion -eq $EX2016_MAJOR -and [System.Version]$SetupVersion -lt [System.Version]$EX2016SETUPEXE_CU23) ) {
            Write-MyError 'Unsupported version of Exchange detected; only Exchange SE, Exchange 2019 CU10 or later, or Exchange 2016 CU23 are supported'
            Exit $ERR_UNSUPPORTEDEX
        }

        If( [System.Version]$SetupVersion -ge $EX2019SETUPEXE_CU15) {
            $Ex2013Exists= Get-ExchangeServerObjects | Where-Object {$_.serialNumber[0] -like 'Version 15.0*'}
            If( $Ex2013Exists) {
                Write-MyError ('Exchange 2013 detected: {0}. Exchange 2019 CU15 or later cannot co-exist with Exchange 2013' -f ($Ex2013Exists | Select-Object Name) -Join ',')
                Exit $ERR_EX19EX2013COEXIST
            }
        }

        If( [System.Version]$FullOSVersion -ge $WS2025_PREFULL -and [System.Version]$SetupVersion -lt $EX2019SETUPEXE_CU15) {
            Write-MyError 'Windows Server 2025 is only supported for Exchange 2019 CU15 or later.'
            Exit $ERR_UNEXPECTEDOS
        }

        If( [System.Version]$FullOSVersion -lt $WS2019_PREFULL -and [System.Version]$MajorSetupVersion -lt [System.Version]$EX2019_MAJOR) {
            Write-MyError 'Exchange 2019/SE is only supported on Windows Server 2019, Windows Server 2022 or Windows Server 2025 (CU15+)'
            Exit $ERR_UNEXPECTEDOS
        }

        If( [System.Version]$FullOSVersion -ge $WS2022_PREFULL -and [System.Version]$FullOSVersion -lt $WS2025_PREFULL -and [System.Version]$SetupVersion -lt $EX2019SETUPEXE_CU12) {
            Write-MyError 'Windows Server 2022 is only supported for Exchange Server 2019 CU12 or later'
            Exit $ERR_UNEXPECTEDOS
        }

        If( $State['NoSetup'] -or $State['Recover'] -or $State['Upgrade']) {
            Write-MyOutput 'Not checking roles (NoSetup, Recover or Upgrade mode)'
        }
        Else {
            Write-MyOutput 'Checking roles to install'
            If ( !( $State['InstallMailbox']) -and !($State['InstallEdge']) ) {
                Write-MyError 'No roles specified to install'
                Exit $ERR_UNKNOWNROLESSPECIFIED
            }
        }

        If( $State["MajorSetupVersion"] -eq $EX2019_MAJOR -and [System.Version]$State["SetupVersion"] -lt [System.Version]$EX2019SETUPEXE_CU14 ) {
            If( $State['DoNotEnableEP']) {
                Write-MyWarning 'DoNotEnableEP is not supported with this Exchange version, ignoring this switch'
                $State['DoNotEnableEP']= $false
            }
            If( $State['DoNotEnableEP_FEEWS']) {
                Write-MyWarning 'DoNotEnableEP_FEEWS is not supported with this Exchange version, ignoring this switch'
                $State['DoNotEnableEP_FEEWS']= $false
            }
        }

        If( ($State["MajorSetupVersion"] -eq $EX2019_MAJOR) -and [System.Version]$State["SetupVersion"] -ge [System.Version]$EX2019SETUPEXE_CU11 ) {
            If( $State['DiagnosticData']) {
                $State['IAcceptSwitch']= '/IAcceptExchangeServerLicenseTerms_DiagnosticDataON'
                Write-MyOutput 'Will deploy Exchange with Data Collection enabled'
            }
            Else {
                 $State['IAcceptSwitch']= '/IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
            }
        }
        Else {
             $State['IAcceptSwitch']= '/IAcceptExchangeServerLicenseTerms'
        }

        if( !($State['InstallEdge'])){
            If( ( Test-ExistingExchangeServer $env:computerName) -and ($State["InstallPhase"] -eq 1)) {
                If( $State['Recover']) {
                    Write-MyOutput 'Recovery mode specified, Exchange server object found'
                }
                Else {
                    If( Test-Path $EXCHANGEINSTALLKEY) {
                        Write-MyOutput 'Existing Exchange server object found in Active Directory, and installation seems present - switching to Upgrade mode'
                        $State['Upgrade']= $true
                    }
                    Else {
                        Write-MyError 'Existing Exchange server object found in Active Directory, but installation missing - please use Recover switch to recover a server'
                        Exit $ERR_PROBLEMEXCHANGESERVEREXISTS
                    }
                }
            }

            Write-MyOutput 'Checking domain membership status ..'
            If(! ( Get-WmiObject Win32_ComputerSystem).PartOfDomain) {
                Write-MyError 'System is not domain-joined'
                Exit $ERR_NOTDOMAINJOINED
            }
        }
        Write-MyOutput 'Checking NIC configuration ..'
        If(! (Get-WmiObject Win32_NetworkAdapterConfiguration -Filter {IPEnabled=True and DHCPEnabled=False})) {
            $AzureHosted= Get-Service | Where-Object {$_.Name -ieq 'Windows Azure Guest Agent' -or $_.Name -ieq 'Windows Azure Network Agent' -or $_.Name -ieq 'Windows Azure Telemetry Service'}
            If( $AzureHosted) {
                Write-MyError "System doesn't have a static IP addresses configured"
                Exit $ERR_NOFIXEDIPADDRESS
            }
            Else {
                Write-MyOutput 'Ignoring absence of static IP address assignment(s) as Azure service(s) are present.'
            }
        }
        Else {
            Write-MyVerbose 'Static IP address(es) assigned.'
        }
        If ( $State['TargetPath']) {
            $Location= Split-Path $State['TargetPath'] -Qualifier
            Write-MyOutput 'Checking installation path ..'
            If( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                Exit $ERR_MDBDBLOGPATH
            }
        }
        If ( $State['InstallMDBLogPath']) {
            $Location= Split-Path $State['InstallMDBLogPath'] -Qualifier
            Write-MyOutput 'Checking MDB log path ..'
            If( !(Test-Path $Location)) {
                Write-MyError "MDB log location unavailable: ($Location)"
                Exit $ERR_MDBDBLOGPATH
            }
        }
        If ( $State['InstallMDBDBPath']) {
            $Location= Split-Path $State['InstallMDBDBPath'] -Qualifier
            Write-MyOutput 'Checking MDB database path ..'
            If( !(Test-Path $Location)) {
                Write-MyError "MDB database location unavailable: ($Location)"
                Exit $ERR_MDBDBLOGPATH
            }
        }
        if( !($State['InstallEdge'])){
            Write-MyOutput 'Checking Exchange Forest Schema Version'
            If( $State['MajorSetupVersion'] -ge $EX2019_MAJOR) {
                $minFFL= $EX2019_MINFORESTLEVEL
                $minDFL= $EX2019_MINDOMAINLEVEL
            }
            Else {
                $minFFL= $EX2016_MINFORESTLEVEL
                $minDFL= $EX2016_MINDOMAINLEVEL
            }
            $EFL= Get-ExchangeForestLevel
            If( $EFL) {
                Write-MyOutput "Exchange Forest Schema Version is $EFL"
            }
            Else {
                Write-MyOutput 'Active Directory is not prepared'
            }
            If( $EFL -lt $minFFL) {
                If( $State['InstallPhase'] -eq 4) {
                    # Only check before starting setup
                    Write-MyError "Minimum required FFL version is $minFFL, aborting"
                    Exit $ERR_BADFORESTLEVEL
                }
            }

            Write-MyOutput 'Checking Exchange Domain Version'
            $EDV= Get-ExchangeDomainLevel
            If( $EDV) {
                Write-MyOutput "Exchange Domain Version is $EDV"
            }
            If( $EDV -lt $minDFL) {
                If( $State['InstallPhase'] -eq 4) {
                    # Only check before starting setup
                    Write-MyError "Minimum required DFL version is $minDFL, aborting"
                    Exit $ERR_BADDOMAINLEVEL
                }
            }

            Write-MyOutput 'Checking domain mode'
            If( Test-DomainNativeMode -eq $DOMAIN_MIXEDMODE) {
                Write-MyError 'Domain is in mixed mode, native mode is required'
                Exit $ERR_ADMIXEDMODE
            }
            Else {
                Write-MyOutput 'Domain is in native mode'
            }

            Write-MyOutput 'Checking Forest Functional Level'
            $FFL= Get-ForestFunctionalLevel
            If( $MajorVersion -eq $EX2019_MAJOR) {
                If( $FFL -lt $FOREST_LEVEL2012R2) {
                    Write-MyError ('Exchange Server 2019 or later requires Forest Functionality Level 2012R2 ({0}).' -f $FFL)
                    Exit $ERR_ADFORESTLEVEL
                }
                Else {
                    Write-MyOutput ('Forest Functional Level is {0} ({1})' -f $FFL, (Get-FFLText $FFL))
                }
            }
            Else {
                If( $FFL -lt $FOREST_LEVEL2012) {
                    Write-MyError ('Exchange Server 2016 or later requires Forest Functionality Level 2012 ({0}).' -f $FFL)
                    Exit $ERR_ADFORESTLEVEL
                }
                Else {
                    Write-MyOutput ('Forest Functional Level is OK ({0})' -f $FFL)
                }
            }
        }
        If( Get-PSExecutionPolicy) {
            # Referring to http://support.microsoft.com/kb/2810617/en
            Write-MyWarning 'PowerShell Execution Policy is configured through GPO and may prohibit Exchange Setup. Clearing entry.'
            Remove-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell -Name ExecutionPolicy -Force
        }
    }

    Function Cleanup {
        Write-MyOutput "Cleaning up .."
        If( Get-WindowsFeature Bits) {
            Write-MyOutput "Removing BITS feature"
            Remove-WindowsFeature Bits
        }
        Write-MyVerbose "Removing state file $Statefile"
        Remove-Item $Statefile
    }

    Function LockScreen {
        Write-MyVerbose 'Locking system'
        rundll32.exe user32.dll,LockWorkStation
    }

    Function Enable-HighPerformancePowerPlan {
        Write-MyVerbose 'Configuring Power Plan'
        $null= Start-Process -FilePath 'powercfg.exe' -ArgumentList ('/setactive','8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c') -NoNewWindow -PassThru -Wait
        $CurrentPlan = Get-WmiObject -Namespace root\cimv2\power -Class win32_PowerPlan | Where-Object { $_.IsActive }
        Write-MyOutput "Power Plan active: $($CurrentPlan.ElementName)"
    }

    Function Disable-NICPowerManagement {
        # http://support.microsoft.com/kb/2740020
        Write-MyVerbose 'Disabling Power Management on Network Adapters'
        # Find physical adapters that are OK and are not disabled
        $NICs = Get-WmiObject -ClassName Win32_NetworkAdapter | Where-Object { $_.AdapterTypeId -eq 0 -and $_.PhysicalAdapter -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22 }
        ForEach( $NIC in $NICs) {
                $PNPDeviceID= ($NIC.PNPDeviceID).ToUpper()
                $NICPowerMgt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi | Where-Object { $_.instancename -match [regex]::escape( $PNPDeviceID) }
                If ($NICPowerMgt.Enable) {
                    $NICPowerMgt.Enable = $false
                    $NICPowerMgt.psbase.Put() | Out-Null
                    If ($NICPowerMgt.Enable) {
                        Write-MyError "Problem disabling power management on $($NIC.Name) ($PNPDeviceID)"
                    } else {
                        Write-MyOutput "Disabled power management on $($NIC.Name) ($PNPDeviceID)"
                    }
                }
                Else {
                    Write-MyVerbose "Power management already disabled on $($NIC.Name) ($PNPDeviceID)"
                }
        }
    }

    Function Set-Pagefile {
        Write-MyVerbose 'Checking Pagefile Configuration'
        $CS = Get-WmiObject -Class Win32_ComputerSystem -EnableAllPrivileges
        If ($CS.AutomaticManagedPagefile) {
            Write-MyVerbose 'System configured to use Automatic Managed Pagefile, reconfiguring'
            Try {
                $CS.AutomaticManagedPagefile = $false
                $InstalledMem= $CS.TotalPhysicalMemory
                If( $State["MajorSetupVersion"] -ge $EX2019_MAJOR) {
                    # 25% of RAM
                    $DesiredSize= [int]($InstalledMem / 4 / 1MB)
                    Write-MyVerbose ('Configuring PageFile to 25% of Total Memory: {0}MB' -f $DesiredSize)
                }
                Else {
                    # RAM + 10 MB, with maximum of 32GB + 10MB
                    $DesiredSize= (($InstalledMem + 10MB), (32GB+10MB)| Measure-Object -Minimum).Minimum / 1MB
                    Write-MyVerbose ('Configuring PageFile Total Memory+10MB with maximum of 32GB+10MB: {0}MB' -f $DesiredSize)
                }
                $null= $CS.Put()
                $CPF= Get-WmiObject -Class Win32_PageFileSetting
                $CPF.InitialSize= $DesiredSize
                $CPF.MaximumSize= $DesiredSize
                $null= $CPF.Put()
            }
            Catch {
                Write-MyError "Problem reconfiguring pagefile: $($ERROR[0])"
            }
            $CPF= Get-WmiObject -Class Win32_PageFileSetting
            Write-MyOutput "Pagefile set to manual, initial/maximum size: $($CPF.InitialSize)MB / $($CPF.MaximumSize)MB"
        }
        Else {
            Write-MyVerbose 'Manually configured page file, skipping configuration'
        }
    }

    Function Set-TCPSettings {
        Write-MyVerbose 'Configuring RPC Timeout setting'
        $RegKey= "HKLM:\Software\Policies\Microsoft\Windows NT\RPC"
        $RegName= "MinimumConnectionTimeout"
        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            If( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                New-Item -Path (Split-Path $RegKey -Parent) -Name (Split-Path $RegKey -Leaf) -ErrorAction SilentlyContinue | out-null
            }
        }
        Write-MyOutput 'Setting RPC Timeout to 120 seconds'
        New-ItemProperty -Path $RegKey -Name $RegName  -Value 120 -ErrorAction SilentlyContinue| out-null

        Write-MyVerbose 'Configuring Keep-Alive Timeout setting'
        $RegKey= "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        $RegName= "KeepAliveTime"
        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            If( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                New-Item -Path (Split-Path $RegKey -Parent) -Name (Split-Path $RegKey -Leaf) -ErrorAction SilentlyContinue | out-null
            }
        }
        Write-MyOutput 'Setting Keep-Alive Timeout to 15 minutes'
        New-ItemProperty -Path $RegKey -Name $RegName  -Value 900000 -ErrorAction SilentlyContinue| out-null
    }

    Function Disable-SSL3 {
        # SSL3 disabling/Poodle, https://support.microsoft.com/en-us/kb/187498
        Write-MyVerbose 'Disabling SSL3 protocol for services'
        $RegKey= "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
        $RegName= "Enabled"
        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            If( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
                New-Item -Path (Split-Path $RegKey -Parent) -Name (Split-Path $RegKey -Leaf) -Force -ErrorAction SilentlyContinue | out-null
            }
        }
        Write-MyVerbose "Setting registry $RegKey\$RegName to 0"
        New-ItemProperty -Path $RegKey -Name $RegName  -Value 0 -ErrorAction SilentlyContinue -Force| out-null
    }

    Function Disable-RC4 {
        # https://support.microsoft.com/en-us/kb/2868725
        # Note: Can't use regular New-Item as registry path contains '/' (always interpreted as path splitter)
        Write-MyVerbose 'Disabling RC4 protocol for services'
        $RC4Keys= @('RC4 128/128', 'RC4 40/128', 'RC4 56/128')
        $RegKey= 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
        $RegName= "Enabled"
        ForEach( $RC4Key in $RC4Keys) {
            If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
                If( -not (Test-Path $RegKey -ErrorAction SilentlyContinue)) {
            		$RegHandle = (get-item 'HKLM:\').OpenSubKey( $RegKey, $true)
		            $RegHandle.CreateSubKey( $RC4Key) | out-null
		            $RegHandle.Close()
                }
            }
            Write-MyVerbose "Setting registry $RegKey\$RegName\RC4Key to 0"
            New-ItemProperty -Path (Join-Path (Join-Path 'HKLM:\' $RegKey) $RC4Key) -Name $RegName  -Value 0 -Force -ErrorAction SilentlyContinue| out-null
        }
    }

    Function Enable-ECC {
        # https://learn.microsoft.com/en-us/exchange/architecture/client-access/certificates?view=exchserver-2019#elliptic-curve-cryptography-certificates-support-in-exchange-server
        Write-MyVerbose 'Enabling Elliptic Curve Cryptography support'

        $RegKey= 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics'
        $RegName= 'EnableEccCertificateSupport'

        If( -not( Get-ItemProperty -Path $RegKey -Name $RegName -ErrorAction SilentlyContinue)) {
            Write-MyVerbose ('Setting {0}\{1} to 1' -f $RegKey, $RegName)
            New-ItemProperty -Path $RegKey -Name $RegName -Value 1 -Type String -Force -ErrorAction SilentlyContinue
        }

        # If overrides were configured, disable these (obsolete and not fully supporting ECC)
        $Override= Get-SettingOverride | Where-Object { ($_.SectionName -eq "ECCCertificateSupport") -and ($_.Parameters -eq "Enabled=true")}
        If( $Override) {
            Write-MyVerbose ('Override for ECC found, removing (obsolete)')
            $Override | Remove-SettingOverride
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh
            Restart-Service -Name W3SVC, WAS -Force
        }
        Else {
            Write-MyVerbose ('No override configuration for ECC found')
        }
    }

    Function Enable-CBC {
        # https://support.microsoft.com/en-us/topic/enable-support-for-aes256-cbc-encrypted-content-in-exchange-server-august-2023-su-add63652-ee17-4428-8928-ddc45339f99e
        Write-MyVerbose 'Enabling AES256-CBC mode of encryption support'

        $Override= Get-SettingOverride | Where-Object { ($_.SectionName -eq "EnableEncryptionAlgorithmCBC") -and ($_.Parameters -eq "Enabled=True")}
        If( $Override) {
            Write-MyVerbose ('Configuration for CBC already configured')
        }
        Else {
            New-SettingOverride -Name "EnableEncryptionAlgorithmCBC" -Parameters @("Enabled=True") -Component Encryption -Reason "Enable CBC encryption" -Section EnableEncryptionAlgorithmCBC
            Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh
            Restart-Service -Name W3SVC, WAS -Force
        }
    }

    Function Enable-AMSI {
        param(
            [string[]]$ConfigParam= @("EnabledEcp=True","EnabledEws=True","EnabledOwa=True","EnabledPowerShell=True")
        )
        # https://learn.microsoft.com/en-us/exchange/antispam-and-antimalware/amsi-integration-with-exchange?view=exchserver-2019#enable-exchange-server-amsi-body-scanning
        Write-MyVerbose 'Enabling AMSI body scanning for OWA, ECP, EWS and PowerShell'

        New-SettingOverride -Name "EnableAMSIBodyScan" -Component Cafe -Section AmsiRequestBodyScanning -Parameters $ConfigParam -Reason "Enabling AMSI body Scan"
        Get-ExchangeDiagnosticInfo -Process Microsoft.Exchange.Directory.TopologyService -Component VariantConfiguration -Argument Refresh
        Restart-Service -Name W3SVC, WAS -Force
    }

    Function Set-TLSSettings {

        param(
            [switch]$TLS12,
            [switch]$TLS13
        )

        If( $TLS12) {

            # configure the .NET Framework 4.x Schannel inheritance
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" -Name "SystemDefaultTlsVersions" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" -Name "SystemDefaultTlsVersions" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -Type DWord

            # Enable TLS 1.2 for client and server connections
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" -Name "TLS 1.2" -ErrorAction SilentlyContinue
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" -Name "Client" -ErrorAction SilentlyContinue
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" -Name "Server" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "DisabledByDefault" -Value 0 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "Enabled" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" -Name "DisabledByDefault" -Value 0 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" -Name "Enabled" -Value 1 -Type DWord
        }
        Else {

            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" -Name "TLS 1.2" -ErrorAction SilentlyContinue
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" -Name "Client" -ErrorAction SilentlyContinue
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2" -Name "Server" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "DisabledByDefault" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "Enabled" -Value 0 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" -Name "DisabledByDefault" -Value 1 -Type DWord
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server" -Name "Enabled" -Value 0 -Type DWord
        }

        If( [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL -and [System.Version]$SetupVersion -ge [System.Version]$EX2019SETUPEXE_CU15) {
            If( $TLS13) {

                # configure the .NET Framework 4.x Schannel inheritance
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" -Name "SystemDefaultTlsVersions" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" -Name "SystemDefaultTlsVersions" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -Type DWord

                # Enable TLS 1.3 for client and server connections
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" -Name "TLS 1.3" -ErrorAction SilentlyContinue
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3" -Name "Client" -ErrorAction SilentlyContinue
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3" -Name "Server" -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client" -Name "DisabledByDefault" -Value 0 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client" -Name "Enabled" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Name "DisabledByDefault" -Value 0 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Name "Enabled" -Value 1 -Type DWord

                # Configure the TLS 1.3 cipher suites
                Enable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -Position 0
                Enable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -Position 1
            }
            Else {

                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" -Name "TLS 1.3" -ErrorAction SilentlyContinue
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3" -Name "Client" -ErrorAction SilentlyContinue
                New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3" -Name "Server" -ErrorAction SilentlyContinue
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client" -Name "DisabledByDefault" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Client" -Name "Enabled" -Value 0 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Name "DisabledByDefault" -Value 1 -Type DWord
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Name "Enabled" -Value 0 -Type DWord

                Disable-TlsCipherSuite -Name TLS_AES_256_GCM_SHA384 -ErrorAction SilentlyContinue
                Disable-TlsCipherSuite -Name TLS_AES_128_GCM_SHA256 -ErrorAction SilentlyContinue
            }
        }
        Else {
            Write-MyWarning 'TLS13 configuration not supported for this OS or Exchange version'
        }

    }

    Function Enable-WindowsDefenderExclusions {

        If( Get-Command -Name Add-MpPreference -ErrorAction SilentlyContinue) {
            $SystemRoot= "$Env:SystemRoot"
            $SystemDrive= "$Env:SystemDrive"

            Write-MyOutput 'Configuring Windows Defender folder exclusions'
            If( $State['TargetPath']) {
                $InstallFolder= $State['TargetPath']
            }
            Else {
                # TargetPath not specified, using default location
                $InstallFolder= 'C:\Program Files\Microsoft\Exchange Server\V15'
            }

            $Locations= @(
                "$SystemRoot|Cluster",
                "$InstallFolder|ClientAccess\OAB,FIP-FS,GroupMetrics,Logging,Mailbox",
                "$InstallFolder\TransportRoles\Data|IpFilter,Queue,SenderReputation,Temp",
                "$InstallFolder\TransportRoles|Logs,Pickup,Replay",
                "$InstallFolder\UnifiedMessaging|Grammars,Prompts,Temp,VoiceMail",
                "$InstallFolder|Working\OleConverter",
                "$SystemDrive\InetPub\Temp|IIS Temporary Compressed Files",
                "$SystemDrive|Temp\OICE_*"
            )

            ForEach( $Location in $Locations) {
                $Location
                $Parts= $Location -split '\|'
                $Items= $Parts[1] -split ','
                ForEach( $Item in $Items) {
                    $ExcludeLocation= Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding location $ExcludeLocation"
                    try {
                        Add-MpPreference -ExclusionPath $ExcludeLocation -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            Write-MyOutput 'Configuring Windows Defender exclusions: NodeRunner process'
            $Processes= @(
                "$InstallFolder\Bin|ComplianceAuditService.exe,Microsoft.Exchange.Directory.TopologyService.exe,Microsoft.Exchange.EdgeSyncSvc.exe,Microsoft.Exchange.Notifications.Broker.exe,Microsoft.Exchange.ProtectedServiceHost.exe,Microsoft.Exchange.RPCClientAccess.Service.exe,Microsoft.Exchange.Search.Service.exe,Microsoft.Exchange.Store.Service.exe,Microsoft.Exchange.Store.Worker.exe,MSExchangeCompliance.exe,MSExchangeDagMgmt.exe,MSExchangeDelivery.exe,MSExchangeFrontendTransport.exe,MSExchangeMailboxAssistants.exe,MSExchangeMailboxReplication.exe,MSExchangeRepl.exe,MSExchangeSubmission.exe,MSExchangeThrottling.exe,OleConverter.exe,UmService.exe,UmWorkerProcess.exe,wsbexchange.exe,EdgeTransport.exe,Microsoft.Exchange.AntispamUpdateSvc.exe,Microsoft.Exchange.Diagnostics.Service.exe,Microsoft.Exchange.Servicehost.exe,MSExchangeHMHost.exe,MSExchangeHMWorker.exe,MSExchangeTransport.exe,MSExchangeTransportLogSearch.exe",
                "$InstallFolder\FIP-FS\Bin|fms.exe,ScanEngineTest.exe,ScanningProcess.exe,UpdateService.exe",
                "$InstallFolder|Bin\Search\Ceres|HostController\HostControllerService.exe,Runtime\1.0\Noderunner.exe,ParserServer\ParserServer.exe",
                "$InstallFolder|FrontEnd\PopImap|Microsoft.Exchange.Imap4.exe,Microsoft.Exchange.Pop3.exe",
                "$InstallFolder|ClientAccess\PopImap\Microsoft.Exchange.Imap4service.exe,Microsoft.Exchange.Pop3service.exe",
                "$InstallFolder|FrontEnd\CallRouter|Microsoft.Exchange.UM.CallRouter.exe",
                "$InstallFolder|TransportRoles\agents\Hygiene\Microsoft.Exchange.ContentFilter.Wrapper.exe"
            )

            ForEach( $Process in $Processes) {
                $Parts= $Process -split '\|'
                $Items= $Parts[1] -split ','
                ForEach( $Item in $Items) {
                    $ExcludeProcess= Join-Path -Path $Parts[0] -ChildPath $Item
                    Write-MyVerbose "WindowsDefender: Excluding process $ExcludeProcess"
                    try {
                        Add-MpPreference -ExclusionProcess $ExcludeProcess -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-MyWarning $_.Exception.Message
                    }
                }
            }

            $Extensions= 'dsc', 'txt', 'cfg', 'grxml', 'lzx', 'config', 'chk', 'edb', 'jfm', 'jrs', 'log', 'que'
            ForEach( $Extension in $Extensions) {
                $ExcludeExtension= '.{0}' -f $Extension
                Write-MyVerbose "WindowsDefender: Excluding extension $ExcludeExtension"
                try {
                    Add-MpPreference -ExclusionExtension $ExcludeExtension -ErrorAction SilentlyContinue
                }
                catch {
                    Write-MyWarning $_.Exception.Message
                }
            }
        }
        Else {
            Write-MyVerbose 'Windows Defender not installed'
        }
    }

    # Return location of mounted drive if ISO specified
    Function Resolve-SourcePath {
        Param (
            [String]$SourceImage
        )
        $disk= Get-DiskImage -ImagePath $SourceImage -ErrorAction SilentlyContinue
        If( $disk) {
            If( $disk.Attached) {
                $vol= $disk | Get-Volume -ErrorAction SilentlyContinue
                If( $vol) {
                    $Drive= $vol.DriveLetter
                }
                Else {
                    Write-Verbose ('{0} already attached but no drive letter - will mount again' -f $SourceImage)
                    $Drive= (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
                }
            }
            Else {
                $Drive= (Mount-DiskImage -ImagePath $SourceImage -PassThru | Get-Volume).DriveLetter
            }
            $SourcePath= '{0}:\' -f $Drive
            Write-Verbose ('Mounted {0} on drive {1}' -f $SourceImage, $SourcePath)
            return $SourcePath
        }
        Else {
            return $null
        }
    }

    Function Get-VCRuntime {
        Param (
            [String]$version
        )
        Write-MyVerbose ('Looking for presence of Visual C++ v{0} Runtime' -f $version)
        $RegPaths= @(
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\Runtimes\x64',
            'HKLM:\Software\WOW6432Node\Microsoft\VisualStudio\{0}\VC\VCRedist\x64',
            'HKLM:\Software\Microsoft\VisualStudio\{0}\VC\VCRedist\x64')
        $presence= $false
        ForEach( $RegPath in $RegPaths) {

            $Key= (Get-ItemProperty -Path ($RegPath -f $version) -Name Installed -ErrorAction SilentlyContinue).Installed
            If( $Key -eq 1) {
                $build= (Get-ItemProperty -Path ($RegPath -f $version) -Name Version -ErrorAction SilentlyContinue).Version
            }
        }
        If( $presence) {
            Write-MyVerbose ('Found Visual C++ Runtime v{0}, build {1}' -f $version, $build)
        }
        Else {
            Write-MyVerbose ('Could not find Visual C++ v{0} Runtime installed' -f $version)
        }
        return $presence
    }

    ########################################
    # MAIN
    ########################################

    #Requires -Version 5.1

    $ScriptFullName = $MyInvocation.MyCommand.Path
    $ScriptName = $ScriptFullName.Split("\")[-1]
    $ParameterString= $PSBoundParameters.getEnumerator() -join " "
    $MajorOSVersion= [string](Get-WmiObject Win32_OperatingSystem | Select-Object @{n="Major";e={($_.Version.Split(".")[0]+"."+$_.Version.Split(".")[1])}}).Major
    $MinorOSVersion= [string](Get-WmiObject Win32_OperatingSystem | Select-Object @{n="Minor";e={($_.Version.Split(".")[2])}}).Minor
    $FullOSVersion= ('{0}.{1}' -f $MajorOSVersion, $MinorOSVersion)

    $State=@{}
    $StateFile= "$InstallPath\$($env:computerName)_$($ScriptName)_state.xml"
    $State= Restore-State

    Write-Output "Script $ScriptFullName v$ScriptVersion called using $ParameterString"
    Write-Verbose "Using parameterSet $($PsCmdlet.ParameterSetName)"
    Write-Output ('Running on OS build {0}' -f $FullOSVersion)

    If(! $State.Count) {
        # No state, initialize settings from parameters
        If( $($PsCmdlet.ParameterSetName) -eq "AutoPilot") {
            Write-Error "Running in AutoPilot mode but no state file present"
            Exit $ERR_AUTOPILOTNOSTATEFILE
        }

        $State["InstallMailbox"]= $True
        $State["InstallEdge"]= $InstallEdge
        $State["InstallMDBDBPath"]= $MDBDBPath
        $State["InstallMDBLogPath"]= $MDBLogPath
        $State["InstallMDBName"]= $MDBName
        $State["InstallPhase"]= 0
        $State["OrganizationName"]= $Organization
        $State["AdminAccount"]= $Credentials.UserName
        $State["AdminPassword"]= ($Credentials.Password | ConvertFrom-SecureString -ErrorAction SilentlyContinue)
        If( Get-DiskImage -ImagePath $SourcePath -ErrorAction SilentlyContinue) {
            $State['SourceImage']= $SourcePath
            $State["SourcePath"]= Resolve-SourcePath -SourceImage $SourcePath
        }
        Else {
            $State['SourceImage']= $null
            $State["SourcePath"]= $SourcePath
        }

        $State["SetupVersion"]= ( Get-DetectedFileVersion "$($State["SourcePath"])\setup.exe")
        $State["TargetPath"]= $TargetPath
        $State["AutoPilot"]= $AutoPilot
        $State["IncludeFixes"]= $IncludeFixes
        $State["NoSetup"]= $NoSetup
        $State["Recover"]= $Recover
        $State["Upgrade"]= $false
        $State["Install481"]= $False
        $State["VCRedist2012"]= $False
        $State["VCRedist2013"]= $False
        $State["DisableSSL3"]= $DisableSSL3
        $State["DisableRC4"]= $DisableRC4
        $State["EnableECC"]= $EnableECC
        $State["EnableCBC"]= -not $NoCBC
        $State["EnableTLS12"]= $EnableTLS12
        $State["EnableTLS13"]= $EnableTLS13
        $State["DoNotEnableEP"]= $DoNotEnableEP
        $State["DoNotEnableEP_FEEWS"]= $DoNotEnableEP_FEEWS
        $State["SkipRolesCheck"]= $SkipRolesCheck
        $State["SCP"]= $SCP
        $State["DiagnosticData"]= $DiagnosticData
        $State["Lock"]= $Lock
        $State["EdgeDNSSuffix"]= $EdgeDNSSuffix
        $State["InstallPath"]= $InstallPath
        $State["TranscriptFile"]= "$($State["InstallPath"])\$($env:computerName)_$($ScriptName)_$(Get-Date -format "yyyyMMddHHmmss").log"

        # Store Server Manager state
        $State['DoNotOpenServerManagerAtLogon']= (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\ServerManager' -Name DoNotOpenServerManagerAtLogon -ErrorAction SilentlyContinue).DoNotOpenServerManagerAtLogon

        $State["Verbose"]= $VerbosePreference

    }
    Else {
        # Run from saved parameters
        If( $State['SourceImage']) {
            # Mount ISO image, and set SourcePath to actual mounted location to anticipate drive letter changes
            $State["SourcePath"]= Resolve-SourcePath -SourceImage $State['SourceImage']
        }
    }

    If( $State["Lock"] ) {
        LockScreen
    }

    If( $State.containsKey("LastSuccessfulPhase")) {
	Write-MyVerbose "Continuing from last successful phase $($State["InstallPhase"])"
        $State["InstallPhase"]= $State["LastSuccessfulPhase"]
    }
    If( $PSBoundParameters.ContainsKey('Phase')) {
	Write-MyVerbose "Phase manually set to $Phase"
        $State["InstallPhase"]= $Phase
    }
    Else {
        $State["InstallPhase"]++
    }

    # (Re)activate verbose setting (so settings becomes effective after reboot)
    If( $State["Verbose"].Value) {
        $VerbosePreference= $State["Verbose"].Value.ToString()
    }

    # When skipping setup, limit no. of steps
    If( $State["NoSetup"]) {
        $MAX_PHASE = 3
    }
    Else {
        $MAX_PHASE = 6
    }

    If( $AutoPilot -and $State["InstallPhase"] -gt 1) {
        # Wait a little before proceeding
        Write-MyOutput "Will continue unattended installation of Exchange in $COUNTDOWN_TIMER seconds .."
        Start-Sleep -Seconds $COUNTDOWN_TIMER
    }

    Test-Preflight

    Write-MyVerbose "Logging to $($State["TranscriptFile"])"

    # Get rid of the security dialog when spawning exe's etc.
    Disable-OpenFileSecurityWarning

    # Always disable autologon allowing you to "fix" things and reboot intermediately
    Disable-AutoLogon

    Write-MyOutput "Checking for pending reboot .."
    If( Test-RebootPending ) {
        $State["InstallPhase"]--
        If( $State["AutoPilot"]) {
            Write-MyWarning "Reboot pending, will reboot system and rerun phase"
        }
        Else {
            Write-MyError "Reboot pending, please reboot system and restart script (parameters will be saved)"v
        }
    }
    Else {

      Write-MyVerbose "Current phase is $($State["InstallPhase"]) of $MAX_PHASE"

      Write-MyVerbose 'Disabling Server Manager at logon'
      New-ItemProperty -Path 'HKCU:\Software\Microsoft\ServerManager' -Name DoNotOpenServerManagerAtLogon -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue

      Switch ($State["InstallPhase"]) {
        1 {

            If( [System.Version]$FullOSVersion -ge [System.Version]$WS2016_MAJOR) {

                Write-MyOutput ('Exchange setup version {0} detected' -f $State['SetupVersion'])
                Write-MyOutput ('Operating System version {0} detected' -f $FullOSVersion)

                If( $State["NoNet481"]) {
                    Write-MyOutput "NoNet481 specified, will not install .NET Framework 4.8.1"
                    $State["Install481"]= $False
                }
                Else {
                    If([System.Version]$FullOSVersion -lt [System.Version]$WS2022_PREFULL ) {
                        Write-MyOutput "Will install .NET Framework 4.8 as default for this OS"
                        $State["Install481"]= $False
                    }
                    Else {
                        Write-MyOutput "Will install .NET Framework 4.8.1 as default for this OS"
                        $State["Install481"]= $True
                    }
                }

                Write-MyOutput "Will install Visual C++ 2012 Runtime"
                $State["VCRedist2012"]= $True

                Write-MyOutput "Will install Visual C++ 2013 Runtime"
                $State["VCRedist2013"]= $True

            }
            Else {
                Write-MyError ('Operating System version {0} not supported' -f $FullOSVersion)
                Exit $ERR_UNEXPECTEDOS
            }
            Write-MyOutput "Installing Operating System prerequisites"
            Install-WindowsFeatures $MajorOSVersion
        }

        2 {
            Write-MyOutput "Installing BITS module"
            Import-Module BITSTransfer

            # Check .NET FrameWork 4.8.1 needs to be installed
            If( $State["Install481"]) {

                Remove-NETFrameworkInstallBlock '4.8.1' '-' '481'
                If( (Get-NETVersion) -lt $NETVERSION_481) {
                    Install-MyPackage "-" "Microsoft .NET Framework 4.8.1" "NDP481-x86-x64-AllOS-ENU.exe" "https://download.microsoft.com/download/4/b/2/cd00d4ed-ebdd-49ee-8a33-eabc3d1030e3/NDP481-x86-x64-AllOS-ENU.exe" ("/q", "/norestart")
                }
                Else {
                    Write-MyOutput ".NET Framework 4.8.1 or later detected"
                }
            }
            Else {
                Write-MyOutput ('Keeping current .NET Framework ({0})' -f (Get-NETVersion))
                Set-NETFrameworkInstallBlock '4.8.1' '-' '481'
            }

            # OS specific hotfixes

            If( [System.Version]$FullOSVersion -ge [System.Version]$WS2016_MAJOR -and [System.Version]$FullOSVersion -lt [System.Version]$WS2019_PREFULL) {
                # WS2016
                Install-MyPackage "KB3206632" "Cumulative Update for Windows Server 2016 for x64-based Systems" "windows10.0-kb3206632-x64_b2e20b7e1aa65288007de21e88cd21c3ffb05110.msu" "http://download.windowsupdate.com/d/msdownload/update/software/secu/2016/12/windows10.0-kb3206632-x64_b2e20b7e1aa65288007de21e88cd21c3ffb05110.msu" ("/quiet", "/norestart")
            }
            If( [System.Version]$FullOSVersion -ge [System.Version]$WS2019_PREFULL -and [System.Version]$FullOSVersion -lt [System.Version]$WS2022_PREFULL) {
                # WS2019
            }
            If( [System.Version]$FullOSVersion -ge [System.Version]$WS2022_PREFULL -and [System.Version]$FullOSVersion -lt [System.Version]$WS2025_PREFULL) {
                # WS2022
            }
            If( [System.Version]$FullOSVersion -ge [System.Version]$WS2025_PREFULL) {
                # WS2025
            }

            # Check if need to install VC++ Runtimes
            if( ($State['InstallEdge'])){
                If( -not (Get-VCRuntime -version '11.0') -and $State["VCRedist2012"] ) {
                    Install-MyPackage "" "Visual C++ 2012 Redistributable" "vcredist_x64_2012.exe" "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe" ("/install", "/quiet", "/norestart")
                }
            }

            If( -not (Get-VCRuntime -version '12.0') -and $State["VCRedist2013"] ) {
                Install-MyPackage "" "Visual C++ 2013 Redistributable" "vcredist_x64_2013.exe" "https://aka.ms/highdpimfc2013x64enu" ("/install", "/quiet", "/norestart")
            }

            # URL Rewrite module
            Install-MyPackage "{9BCA2118-F753-4A1E-BCF3-5A820729965C}" "URL Rewrite Module 2.1" "rewrite_amd64_en-US.msi" "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi" ("/quiet", "/norestart")

        }

        3 {
            if( !($State['InstallEdge'])){
                Write-MyOutput "Installing Exchange prerequisites (continued)"
                If( [System.Version]$FullOSVersion -ge [System.Version]$WS2019_PREFULL -and (Test-ServerCore) ) {
                    Install-MyPackage "{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" "Unified Communications Managed API 4.0 Runtime (Core)" "Setup.exe" (Join-Path -Path $State['SourcePath'] -ChildPath 'UcmaRedist\Setup.exe') ("/passive", "/norestart") -NoDownload
                }
                Else {
                    Install-MyPackage "{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" "Unified Communications Managed API 4.0 Runtime" "UcmaRuntimeSetup.exe" "https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe" ("/passive", "/norestart")
                }
            } else {
                Write-MyOutput 'Setting Primary DNS Suffix'
                Set-EdgeDNSSuffix -DNSSuffix $State['EdgeDNSSuffix']
            }
            If ($State["OrganizationName"]) {
                Write-MyOutput "Preparing Active Directory"
                Initialize-Exchange
            }
        }

        4 {
            Write-MyOutput "Installing Exchange"
            Install-Exchange15_
            switch( $State["SCP"]) {
                '' {
                        # Do nothing
                }
                '-' {
                    Write-MyOutput 'Removing Service Connection Point record'
                    Clear-AutodiscoverServiceConnectionPoint $ENV:COMPUTERNAME
                }
                default {
                    Write-MyOutput "Configuring Service Connection Point record as $($State['SCP'])"
                    Set-AutodiscoverServiceConnectionPoint $ENV:COMPUTERNAME $State['SCP']
                }
            }
            If( Get-Service MSExchangeTransport -ErrorAction SilentlyContinue) {
                Write-MyOutput "Configuring MSExchangeTransport startup to Manual"
                Set-Service MSExchangeTransport -StartupType Manual
            }
            If( Get-Service MSExchangeFrontEndTransport -ErrorAction SilentlyContinue) {
                Write-MyOutput "Configuring MSExchangeFrontEndTransport startup to Manual"
                Set-Service MSExchangeFrontEndTransport -StartupType Manual
            }
        }

        5 {
            Write-MyOutput "Post-configuring"

            Enable-WindowsDefenderExclusions
            Enable-HighPerformancePowerPlan
            Disable-NICPowerManagement
            Set-Pagefile
            Set-TCPSettings
            If( $State["DisableSSL3"]) {
                Disable-SSL3
            }
            If( $State["DisableRC4"]) {
                Disable-RC4
            }

            Set-TLSSettings -TLS12 $State["EnableTLS12"] -TLS13 $State["EnableTLS13"]

            Import-ExchangeModule

            If( $State["EnableECC"]) {
                Enable-ECC
            }
            If( $State["EnableCBC"]) {
                Enable-CBC
            }
            If( $State["EnableAMSI"]) {
                Enable-AMSI
            }

            If( $State["InstallMailbox"] ) {
                # Insert your own Mailbox Server code here
            }
            If( $State["InstallEdge"]) {
                # Insert your own Edge Server code here
            }
            # Insert your own generic customizations here

            If( $State["IncludeFixes"]) {
              Write-MyOutput "Installing additional recommended hotfixes and security updates for Exchange"

              $ImagePathVersion= Get-DetectedFileVersion ( (Get-WMIObject -Query 'select * from win32_service where name="MSExchangeServiceHost"').PathName.Trim('"') )
              Write-MyVerbose ('Installed Exchange MSExchangeIS version {0}' -f $ImagePathVersion)

              Switch( $State['ExSetupVersion']) {
                $EX2019SETUPEXE_CU14 {
                    Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2019 CU14 SU3 V2' 'Exchange2019-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/8/0/b/80b356e4-f7b1-4e11-9586-d3132a7a2fc3/Exchange2019-KB5049233-x64-en.exe' ('/passive')
                }
                $EX2019SETUPEXE_CU13 {
                    Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2019 CU13 SU7 V2' 'Exchange2019-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/4/e/5/4e5cbbcc-5894-457d-88c4-c0b2ff7f208f/Exchange2019-KB5049233-x64-en.exe' ('/passive')
                }
                $EX2016SETUPEXE_CU23 {
                    Install-MyPackage 'KB5049233' 'Security Update For Exchange Server 2016 CU23 SU14 V2' 'Exchange2016-KB5049233-x64-en.exe' 'https://download.microsoft.com/download/0/9/9/0998c26c-8eb6-403a-b97a-ae44c4db5e20/Exchange2016-KB5049233-x64-en.exe' ('/passive')
                }
                default {

                }
              }

            }
        }

        6 {
            If( Get-Service MSExchangeTransport -ErrorAction SilentlyContinue) {
                Write-MyOutput "Configuring MSExchangeTransport startup to Automatic"
                Set-Service MSExchangeTransport -StartupType Automatic
            }
            If( Get-Service MSExchangeFrontEndTransport -ErrorAction SilentlyContinue) {
                Write-MyOutput "Configuring MSExchangeFrontEndTransport startup to Automatic"
                Set-Service MSExchangeFrontEndTransport -StartupType Automatic
            }

            Write-MyVerbose 'Restoring Server Manager startup configuration'
            If( $State['DoNotOpenServerManagerAtLogon']) {
                New-ItemProperty -Path 'HKCU:\Software\Microsoft\ServerManager' -Name DoNotOpenServerManagerAtLogon -Value $State['DoNotOpenServerManagerAtLogon'] -Force -ErrorAction SilentlyContinue | Out-Null
            }

            if( !($State['InstallEdge'])){
                Write-MyVerbose 'Performing Health Monitor checks..'
                # Warmup IIS
                $web = New-Object Net.WebClient
                # To ignore self-signed cert warnings
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
                'OWA', 'ECP', 'EWS', 'Autodiscover', 'Microsoft-Server-ActiveSync', 'OAB', 'mapi', 'rpc' | ForEach-Object {
                    $url = 'https://localhost/{0}/healthcheck.htm' -f $_
                    Try {
                        $output = $web.DownloadString($url)
                        Write-MyOutput ('Healthcheck {0}: {1}' -f $url, ($output -split '<')[0])
                    }
                    Catch {
                        Write-MyWarning ('Healthcheck {0}: {1}' -f $url, 'ERR')
                    }
                }
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
            }
            Else {
                Write-MyVerbose 'InstallEdge Mode, skipping IIS health monitor checks'
            }

            Enable-UAC
            Enable-IEESC
            Write-MyOutput "Setup finished - We're good to go."
        }

        default {
            Write-MyError "Unknown phase ($($State["InstallPhase"]))"
            Exit $ERR_UNEXPTECTEDPHASE
        }
      }
    }
    $State["LastSuccessfulPhase"]= $State["InstallPhase"]
    Enable-OpenFileSecurityWarning
    Save-State $State
    If( $State['SourceImage']) {
        Dismount-DiskImage -ImagePath $State['SourceImage']
    }

    If( $State["AutoPilot"]) {
        If( $State["InstallPhase"] -lt $MAX_PHASE) {
        	Write-MyVerbose "Preparing system for next phase"
	        Disable-UAC
            Disable-IEESC
            Enable-AutoLogon
            Enable-RunOnce
        }
        Else {
            Cleanup
        }
        Write-MyOutput "Rebooting in $COUNTDOWN_TIMER seconds .."
        Start-Sleep -Seconds $COUNTDOWN_TIMER
        Restart-Computer -Force
    }
    Exit $ERR_OK

} #Process
