<# 
    .SYNOPSIS 
    This script removes HealthMailbox or SystemMailbox accounts in MESO container that are lacking a mailbox database attribute.

    Version 1.1, 2018-08-27 

    Author: Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Please send ideas, comments and suggestions to support@granikos.eu

    .LINK 
    http://scripts.granikos.eu

    .DESCRIPTION 
    This script removes Active Directory objects for HealthMailboxes or SystemMailboxes in the Microsoft Exchange System Objects (MESO) container that do not have a homeMDB attribute set.

    Call the script with -WhatIf to check objects first  

    .NOTES 
    
    Requirements 
    - Windows Server 2012 R2 or newer
    - Exchange Server 2013 or newer
    - Utilizes global functions library --> https://www.granikos.eu/en/justcantgetenough/PostId/210/globalfunctions-shared-powershell-library 
    - Delete permission to MESO container objects
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 
    1.1 Updated Examples

    .PARAMETER Domain
    The default Active Directory domain path to query

    .PARAMETER HealthMailbox
    Remove Boxes that have an empty database attribute  

    .PARAMETER SystemMailbox
    Remove SystemBoxes that have an empty database attribute

    .EXAMPLE 
    Perform a WhatIf run in preparation to removing SystemMailboxes having an empty database attribute, using the default domain set in the param section
    .\Remove-OrphanedMailboxAccounts.ps1 -SystemMailbox -WhatIf
    
    .EXAMPLE 
    Remove HealthMailbox(es) having an empty database attribute, using the default domain set in the param section.
    .\Remove-OrphanedMailboxAccounts.ps1 -HealthMailbox

    .EXAMPLE 
    Remove HealthMailbox(es) having an empty database attribute in the AD domain varunagroup.de
    .\Remove-OrphanedMailboxAccounts.ps1 -Domain 'DC=VARUNAGROUP,DC=DE' -HealthMailbox
#>  

[cmdletbinding(SupportsShouldProcess=$True)]
Param(
  [parameter()]
  [string]$Domain = 'DC=MCSMEMAIL,DC=DE',
  [parameter(ParameterSetName='Health')]
  [switch]$HealthMailbox,
  [parameter(ParameterSetName='System')]
  [switch]$SystemMailbox
)

# Import central logging functions and create a new logger
try {
  Import-Module -Name GlobalFunctions
}
catch {
  Write-Host 'Unable to load GlobalFunctions PowerShell module.'
  Write-Host 'Please check http://bit.ly/GlobalFunctions for further instructions'
}

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14

# Purge log files depending on LogFileRetention
$logger.Purge()

function Remove-Mailboxes
{
  if($HealthMailbox) {
    # Search for Monitoring mailboxes
    $SearchBase = ('CN=Monitoring Mailboxes,CN=Microsoft Exchange System Objects,{0}' -f $Domain.ToUpper())
    $objectClass = 'user'
    $Action = 'Cleaning HealthMailboxes'
  }
  if($SystemMailbox) {
    # Search for System Mailboxes in MESO container
    $SearchBase = ('CN=Microsoft Exchange System Objects,{0}' -f $Domain.ToUpper())
    $objectClass = 'msExchSystemMailbox'
    $Action = 'Cleaning SystemMailboxes'
  }

  $logger.Write(('WhatIf Preference: {0}' -f $WhatIfPreference))

  $orphanedAccounts = $null

  try {
    # Fetch user object from selected container having no homeMDB attribute set
    $orphanedAccounts = Get-ADObject -LDAPFilter "(&(!(homeMDB=*))(objectClass=$($objectClass)))" -SearchBase $SearchBase 
  }
  catch {
    # Ooops, something happend
    $logger.Write(('Error querying AD objects using SearchBase {0}. Are you using the correct domain?' -f $SearchBase), 1)
        
    Write-Error -Exception $_.Exception -Message 'Error querying AD objects! Are you using the the correct domain?'
    
    break    
  }
  
  if($orphanedAccounts -ne $null) { 

    Write-Output -InputObject ('Objects found: {0}' -f ($orphanedAccounts | Measure-Object).Count)
    $logger.Write(('{0} | {1} objects found' -f $Action, ($orphanedAccounts | Measure-Object).Count))

    foreach($object in $orphanedAccounts) {

      $logger.Write(('{0} | Delete {1}' -f $Action, $object.DistinguishedName))

      # Remove th -Whatif switch, after verifying the scritp for your production environment
      Remove-ADObject -Identity $object.DistinguishedName -Recursive -Confirm:$false 

    }

  }
  else {
    Write-Output -InputObject 'No objects found!'
  }

}

## MAIN ################################################
Set-ADServerSettings -ViewEntireForest $true

$logger.Write('Script started')

Remove-Mailboxes

$logger.Write('Script finished')