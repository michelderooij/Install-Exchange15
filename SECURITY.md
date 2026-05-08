# Security Policy

## Supported Versions

Security fixes are applied to the **latest version** of `Install-Exchange15.ps1` only. Older snapshots or forks are not maintained.

| Exchange Version          | Supported          |
|---------------------------|--------------------|
| Exchange Server SE (RTM+) | ✅                 |
| Exchange 2019 CU15        | ⚠️ Best-effort     |
| Exchange 2019 CU10–CU14   | ⚠️ Best-effort    |
| Exchange 2016 CU23        | ⚠️ Best-effort    |
| Earlier versions          | ❌                 |

## Reporting a Vulnerability

**Do not open a public GitHub issue for security vulnerabilities.**

Please report them privately via one of:

- **GitHub private vulnerability reporting**: [Security Advisories](../../security/advisories/new)
- **Email**: *security@eightwone.com*

Please include as much of the following information as possible to help understand the nature and scope of the issue:

- **Type of issue** (e.g. credential exposure, arbitrary code execution, privilege escalation, insecure download, etc.)
- **Full path(s) of the source file(s)** related to the manifestation of the issue
- **Location of the affected source code** (tag, branch, commit, or direct URL)
- **Any special configuration required to reproduce the issue**
- **Step-by-step instructions to reproduce the issue**
- **Proof-of-concept or exploit code** (if possible)
- **Impact of the issue**, including how an attacker might exploit it

This information will help triage your report more quickly.

You can expect an acknowledgement within **5 business days** and a resolution or status update within **30 days**.

## Security Considerations for Users

This script runs with elevated privileges (local Administrator, and optionally Schema Admin / Enterprise Admin). Review the following before use:

- **Credentials**: When using `-AutoPilot`, credentials are used to configure auto-logon. Ensure the account is a dedicated service account with the minimum required rights, and that auto-logon is disabled after installation completes.
- **InstallPath / SourcePath**: If pointing to a UNC share, ensure the share is access-controlled. A compromised share could supply malicious binaries.
- **Downloaded prerequisites**: The script may download .NET runtimes, VC++ redistributables, and hotfixes. Verify hashes where possible and use an internal/trusted source via `-InstallPath`.
- **Script integrity**: Always download this script from the official repository. Verify the commit signature or hash before running in production.
- **Hardening parameters**: Use `-DisableSSL3`, `-DisableRC4`, `-EnableTLS12`, `-EnableTLS13`, and `-EnableAMSI` in all production deployments.

## Scope

The following are **in scope** for vulnerability reports:

- Insecure credential handling or storage
- Arbitrary code execution via script parameters or state file
- Insecure download behavior (missing integrity checks)
- Privilege escalation beyond what is documented

The following are **out of scope**:

- Vulnerabilities in Exchange Server itself
- Vulnerabilities in Windows Server, .NET or IIS
- Issues that require physical access to the machine

## Preferred Languages

We prefer all communications to be in English.
