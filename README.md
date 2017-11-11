# Remove-OrphanedMailboxAccounts.ps1

This script removes HealthMailbox or SystemMailbox accounts in MESO container that are lacking a mailbox database attribute.

## Description

This script removes Active Directory objects for HealthMailboxes or SystemMailboxes in the Microsoft Exchange System Objects (MESO) container that do not have a homeMDB attribute set.

Run the script with -WhatIf parameter to check objects first.

## Parameters

### HealthMailbox

Remove Boxes that have an empty database attribute

### SystemMailbox

Remove SystemBoxes that have an empty database attribute

## Requirements

This script utilizes the GlobalFunctions PowerShell module. This module needs to be installed first.

Read more here: [http://bit.ly/GlobalFunctions](http://bit.ly/GlobalFunctions)

## Examples

``` PowerShell
.\Remove-OrphanedMailboxAccounts.ps1 -SystemMailbox -WhatIf
```

Perform a WhatIf run in preparation to removing SystemMailboxes having an empty database attribute

``` PowerShell
.\Remove-OrphanedMailboxAccounts.ps1 -HealthMailbox
```

Remove HealthMailbox(es) having an empty database attribute

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Download and vote at TechNet Gallery

* [https://gallery.technet.microsoft.com/Remove-orphaned-HealthMailb-9f33ccbe](https://gallery.technet.microsoft.com/Remove-orphaned-HealthMailb-9f33ccbe)

## Credits

Written by: Thomas Stensitzki

Stay connected:

* My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
* Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
* LinkedIn:	[http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
* Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
* Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
* Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)