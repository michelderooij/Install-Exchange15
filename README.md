# Install-Exchange15

## Getting Started

This script can install Exchange 2013/2016 prerequisites, optionally create the Exchange
organization (prepares Active Directory) and installs Exchange Server. When the AutoPilot switch is
specified, it will do all the required rebooting and automatic logging on using provided credentials.

To keep track of provided parameters and state, it uses an XML file; if this file is
present, this information will be used to resume the process. Note that you can use a central
location for Install (UNC path with proper permissions) to re-use additional downloads.

### Requirements

* Windows Server 2008 R2 SP1, Windows Server 2012, Windows Server 2012 R2 or Windows Server 2016 (Exchange 2016 CU3+ only).
* Domain-joined system.
* "AutoPilot" mode requires account with elevated administrator privileges.
* When you let the script prepare AD, the account needs proper permissions.
* Edge role not yet supported.

### Usage

Syntax:
Install-Exchange15.ps1 -[InstallCAS|InstallMailbox|InstallMultiRole|Recover|NoSetup] -SourcePath  [-Organization ] [-MDBName ] [-MDBDBPath ] [-MDBLogPath ] [-InstallPath ] [-TargetPath ] [-AutoPilot] [-Credentials ] [-IncludeFixes] [-SCP] [-UseWMF3] [-DisableSSL3] [-Lock]

Examples:

```
$Cred=Get-Credentials
.\Install-Exchange15.ps1 -Organization Fabrikam -InstallMailbox -MDBDBPath C:\MailboxData\MDB1\DB -MDBLogPath C:\MailboxData\MDB1\Log -MDBName MDB1 -InstallPath C:\Install -AutoPilot -Credentials $Cred -SourcePath '\\server\share\Exchange 2013\mu_exchange_server_2013_x64_dvd_1112105' -SCP https://autodiscover.fabrikam.com/autodiscover/autodiscover.xml -Verbose
```
Perform an installation, creating Exchange organization Fabrikam (if it not already exists), using the specified name and location for the initial mailbox database, using provided credentials and
sources at provided location. After setup, alter the SCP value for this server.

```
.\Install-Exchange15.ps1 -Recover -Autopilot -Install -AutoPilot -SourcePath \\server1\sources\ex2016cu2
```
Perform a recovery installation.

### About

For more information on this script, as well as usage and examples, see
the related blog article, [Exchange v15 Unattended Setup](https://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 