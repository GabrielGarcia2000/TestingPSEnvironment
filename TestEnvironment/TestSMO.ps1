#***********************************************************************
# PowerShell : TestSMO.ps1                                             *
#   Function : Test Microsoft Sql Server Management Object (SMO)       *
#            :                                                         *
#***********************************************************************
#                 M O D I F I C A T I O N S                            *
# -- Date -- ---- Name ---- --------- Description -------------------- *
# 11/06/2009 Gabriel Garcia Created.                                   *
#                                                                      *
#***********************************************************************

###
### Example 
### PS C:\> C:\Scripts\TestSMO.ps1
###

# Parameters
$instanceName = "BCIT-PDB-SV01"

# This script gets SQL Server database information using PowerShell
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null

# Create an SMO connection to the instance
$srv = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $instanceName

Write-Host "Instance name: " $srv.Name
Write-Host "Computer Name Physical NetBIOS: " $srv.ComputerNamePhysicalNetBIOS
Write-Host "Edition: " $srv.Edition
Write-Host "Version: " $srv.VersionString

# Disconnect from the SQL Server database
$srv.ConnectionContext.Disconnect()
