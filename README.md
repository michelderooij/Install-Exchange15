# Install-Exchange15

## Overview

This script automates the unattended installation of Microsoft Exchange Server 2016, 2019, and Exchange Server SE on Windows Server 2016 through 2025.

It handles the full installation lifecycle: Windows features, prerequisites (.NET, VC++ runtimes, IIS components), Active Directory preparation, Exchange setup, and post-configuration hardening. With the `-AutoPilot` switch, the script manages automatic reboots and logon cycles, tracking progress in a JSON state file so it can resume exactly where it left off.

Point `-InstallPath` at a central UNC share to share prerequisites and downloads across multiple servers.

## Supported Versions

| Exchange Version | Minimum OS | Maximum OS |
|---|---|---|
| Exchange 2016 CU23 | Windows Server 2016 | Windows Server 2019 |
| Exchange 2019 CU10–CU14 | Windows Server 2019 (Desktop or Core) | Windows Server 2022 (Desktop or Core) |
| Exchange 2019 CU15 | Windows Server 2019 (Desktop or Core) | Windows Server 2025 (Desktop or Core) |
| Exchange Server SE RTM | Windows Server 2019 (Desktop or Core) | Windows Server 2025 (Desktop or Core) |

## Requirements

- PowerShell 5.1 or later
- Domain-joined system (Edge Server role is the exception)
- An account with local administrator rights
- When using `-AutoPilot`: the account must be able to configure auto-logon
- When creating a new Exchange organization (`-Organization`): Schema Admin and Enterprise Admin rights
- Static IP address (Azure-hosted VMs with the Azure Guest Agent are exempt)

## Key Parameters

| Parameter | Description |
|---|---|
| `-SourcePath` | Path to Exchange setup EXE folder or ISO file |
| `-Organization` | Exchange organization name to create. Omit to skip AD preparation. |
| `-InstallEdge` | Install the Edge Transport server role instead of Mailbox |
| `-AutoPilot` | Fully automated mode — handles reboots and resumes automatically |
| `-Credentials` | Credentials AutoPilot uses for automatic logon after each reboot |
| `-InstallPath` | Working folder for state file, logs, and downloaded prerequisites (default: `C:\Install`) |
| `-MDBName` | Name of the initial mailbox database |
| `-MDBDBPath` | Path for the mailbox database file |
| `-MDBLogPath` | Path for the mailbox database transaction logs |
| `-TargetPath` | Exchange binaries installation path (default: `C:\Program Files\Microsoft\Exchange Server\V15`) |
| `-SCP` | Autodiscover Service Connection Point URL to set after installation. Use `-` to clear. |
| `-IncludeFixes` | Install additional recommended hotfixes and security updates |
| `-DisableSSL3` | Disable SSL 3.0 |
| `-DisableRC4` | Disable the RC4 cipher suite |
| `-EnableECC` | Configure Elliptic Curve Cryptography |
| `-EnableTLS12` | Configure TLS 1.2 |
| `-EnableTLS13` | Configure TLS 1.3 (WS2022/WS2025 with Exchange 2019 CU15+) |
| `-EnableAMSI` | Enable AMSI body scanning for ECP, EWS, OWA, and PowerShell virtual directories |
| `-DisableTLS10` | Disable TLS 1.0 |
| `-DisableTLS11` | Disable TLS 1.1 |
| `-DisableInsecureRenegotiation` | Disallow insecure TLS renegotiation (`AllowInsecureRenegoClients` and `AllowInsecureRenegoServers` set to 0) |
| `-DisableWeakCiphers` | Disable weak SCHANNEL ciphers: NULL, DES 56/56, RC4 40/128, RC4 56/128, RC4 64/128, RC4 128/128, Triple DES 168 |
| `-DisableWeakHashAlgorithms` | Disable weak SCHANNEL hash algorithms: MD5 and SHA-1 |
| `-DisableNonForwardSecretKeyExchange` | Disable non-forward-secret key exchange (PKCS/static RSA) |
| `-DisableCredentialGuard` | Disable Credential Guard (`LsaCfgFlags` and `EnableVirtualizationBasedSecurity` set to 0) |
| `-NoSetup` | Install prerequisites only; skip Exchange setup |
| `-Recover` | Run in RecoverServer mode |
| `-NoNet481` | Use .NET 4.8 instead of 4.8.1 |
| `-DoNotEnableEP` | Skip enabling Extended Protection (Exchange 2019 CU14+) |
| `-Lock` | Lock the workstation screen during installation |
| `-DiagnosticData` | Set the initial diagnostic data collection mode |

## Usage Examples

**Basic AutoPilot install with a new Exchange organization (using splatting):**
```powershell
$Cred = Get-Credential

$Params = @{
    Organization                   = 'Fabrikam'
    SourcePath                     = '\\server\share\ExchangeServer2019-x64-CU15'
    InstallPath                    = 'C:\Install'
    Credentials                    = $Cred
    MDBName                        = 'MDB1'
    MDBDBPath                      = 'C:\MailboxData\MDB1\DB'
    MDBLogPath                     = 'C:\MailboxData\MDB1\Log'
    SCP                            = 'https://autodiscover.fabrikam.com/autodiscover/autodiscover.xml'
    AutoPilot                      = $true
    DisableSSL3                    = $true
    DisableRC4                     = $true
    DisableTLS10                   = $true
    DisableTLS11                   = $true
    DisableInsecureRenegotiation   = $true
    DisableWeakCiphers             = $true
    DisableWeakHashAlgorithms      = $true
    DisableNonForwardSecretKeyExchange = $true
    EnableTLS12                    = $true
    EnableECC                      = $true
    EnableAMSI                     = $true
    Verbose                        = $true
}

.\Install-Exchange15.ps1 @Params
```

**Install from ISO with an additional mailbox database:**
```powershell
.\Install-Exchange15.ps1 -MDBName MDB3 `
    -MDBDBPath C:\MailboxData\MDB3\DB\MDB3.edb -MDBLogPath C:\MailboxData\MDB3\Log `
    -AutoPilot -SourcePath D:\Install\ExchangeServer2019-x64-CU15.ISO -Verbose
```

**Resume AutoPilot using saved credentials (state file already present):**
```powershell
$Cred = Get-Credential
.\Install-Exchange15.ps1 -AutoPilot -Credentials $Cred
```

**RecoverServer mode:**
```powershell
.\Install-Exchange15.ps1 -Recover -AutoPilot -SourcePath \\server1\sources\ExchangeServerSE-RTM
```

**Prerequisites only (no setup):**
```powershell
.\Install-Exchange15.ps1 -NoSetup -AutoPilot -InstallPath \\server1\exfiles -SourcePath \\server1\sources\ExchangeServerSE-RTM
```

## Installation Phases

The script runs in six sequential phases. In AutoPilot mode, the system reboots after each phase before the next one begins.

| Phase | Name | What the script does |
|---|---|---|
| 1 | Windows prerequisites | Detects the OS and Exchange version; installs all required Windows roles and features (IIS, RSAT, clustering tools, WCF activation, etc.). Reboots when complete. |
| 2 | Runtime prerequisites | Installs .NET Framework 4.8 or 4.8.1 (based on OS); installs OS-specific hotfixes (e.g. KB3206632 on WS2016); installs the Visual C++ 2012 and 2013 runtimes; installs the IIS URL Rewrite module. Reboots when complete. |
| 3 | Exchange prerequisites & AD preparation | Installs Unified Communications Managed API 4.0 (UCMA); when you specify `-Organization`, runs Exchange AD preparation (`/PrepareSchema`, `/PrepareAD`, `/PrepareDomain`). For Edge servers: sets the primary DNS suffix. Reboots when complete. |
| 4 | Exchange setup | Runs `setup.exe` to install Exchange. When you specify `-SCP`, the script pre-configures or clears the Autodiscover Service Connection Point so the server does not receive client traffic until setup completes. Reboots when complete. |
| 5 | Post-installation configuration | Configures Windows Defender exclusions; sets the high-performance power plan; disables NIC power management; tunes the pagefile and TCP settings; applies the requested TLS configuration (TLS 1.2, TLS 1.3, disable SSL 3.0, disable RC4, disable TLS 1.0, disable TLS 1.1); applies optional SCHANNEL hardening (insecure renegotiation, weak ciphers, weak hash algorithms, non-forward-secret key exchange); disables Credential Guard when requested; enables ECC and/or AMSI when requested; installs additional recommended security updates when you specify `-IncludeFixes`. Reboots when complete. |
| 6 | Finalization | Re-enables Exchange transport services and the Autodiscover app pool; performs IIS health checks for OWA, ECP, EWS, Autodiscover, ActiveSync, OAB, MAPI, and RPC endpoints; re-enables UAC and IE Enhanced Security Configuration; removes auto-logon; locks the workstation when you specify `-Lock`. |

## Required Permissions

| Scenario | Required permissions |
|---|---|
| All scenarios | Run the script in an **elevated PowerShell session** (Run as Administrator) |
| Domain-joined Mailbox installation | **Schema Admins** and **Enterprise Admins** group membership (the script checks at startup; use `-SkipRolesCheck` to bypass) |
| New Exchange organization (`-Organization`) | Schema Admins and Enterprise Admins (as above) |
| AutoPilot mode | The script verifies the supplied credentials before the first reboot; they must be valid domain credentials. The account must also have rights to configure the local auto-logon registry keys. |
| AD preparation only | Run the script under an account that is a member of Schema Admins and Enterprise Admins |
| Edge Server installation | Local administrator only — Edge servers are not domain-joined; the script skips the Schema/Enterprise Admin check automatically |
| Shared install path (`-InstallPath \\server\share`) | The account running the script needs read/write access to the UNC share |

## How AutoPilot Works

1. The script configures auto-logon for the supplied credentials and registers itself in RunOnce.
2. After each phase that requires a reboot, the system restarts and the script resumes automatically.
3. The script tracks progress in a JSON state file named `<ComputerName>_Install-Exchange15.ps1_state.json` in `InstallPath`.
4. The script writes a transcript log to `<ComputerName>_Install-Exchange15.ps1_<timestamp>.log` in `InstallPath`.
5. On completion, the script removes auto-logon and locks the workstation if you specified `-Lock`.

To monitor progress, tail the transcript log or watch the desktop wallpaper, which shows the current phase.

## License

This project is licensed under the MIT License — see [LICENSE.md](LICENSE.md) for details.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

## More Information

Original blog article: [Exchange v15 Unattended Setup](https://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script/)
