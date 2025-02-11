# Install-Exchange15

## Getting Started

This script can install Exchange 2016/2019 prerequisites, optionally create the Exchange
organization (prepares Active Directory) and installs Exchange Server. When the AutoPilot switch is
specified, it will do all the required rebooting and automatic logging on using provided credentials.
To keep track of provided parameters and state, it uses an XML file; if this file is
present, this information will be used to resume the process. Note that you can use a central
location for Install (UNC path with proper permissions) to re-use additional downloads.

For more information on this script, as well as usage and examples, see
the original blog article, [Exchange v15 Unattended Setup](https://eightwone.com/2013/02/18/exchange-2013-unattended-installation-script/).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 
