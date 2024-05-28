# Script Use: Collect Microsoft Teams Rooms (Android & Windows) health data via the Microsoft Graph and output to .CSV
# Version 1.0
# Date: May 2024
#
# This script is provided as-is and no warranties or support can be provided. It is intended for Proof of Conecept environment, use at your own risk.
# Test the script in an isolated environment first.

# Welcome
    Clear-Host
    Write-Host "This PowerShell script will collect all Microsoft Teams Rooms health information from your Microsoft Office 365 Tenant and output to MicrosoftTeamsRoomsReport.CSV."
    Read-Host "Please press <Enter> to continue or press <CTRL+C> to quit."
    Write-Host ""

#########################################################################################################

#======================= Test for Admin Rights ==========================================================
# Test if Windows Powershell is running in Admin mode
function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        }
    if (!(Test-IsAdmin))
        {
            throw "ERROR: Please run Windows Powershell in Admin Mode. Please close this window and restart PowerShell in Admin Mode - Script aborted."
        }
    else 
    {
        Write-Host "SUCCESS: Powershell is running in Admin Mode, proceeding with the deployment..." -Foreground Green
        Write-Host ""   
    }

# Test for at least PowerShell 5.1
    if([Version]'5.1.00000.000' -GT $PSVersionTable.PSVersion)
    {
        Write-Error "ERROR: You must first update to at least PowerShell 5.1 - Script aborted." -ErrorAction Stop
    }
    else {
        Write-Host "You are running at least PowerShell 5.1, continuing..."
        Write-Host ""
    }

### Check if required Powershell Modules are installed, if not, install the Powershell Modules ###
#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force

# Check & Install Microsoft Graph Beta for Teams Powershell Module
Write-Host "Connecting to Microsoft Graph Beta for Teams" -Foreground Green
if (Get-Module -ListAvailable -Name Microsoft.Graph.Beta.Teams) {
    Import-Module Microsoft.Graph.Beta.Teams -ErrorAction Stop
    Connect-MgGraph -ErrorAction Stop
}
else {
    Write-Host "WARNING: The Microsoft Graph Beta Powershell Module for Teams is not installed, installing now..." -Foreground Yellow
    Install-Module Microsoft.Graph.Beta.Teams -ErrorAction Stop
    Import-Module Microsoft.Graph.Beta.Teams -ErrorAction Stop
    Connect-MicrosoftTeams -ErrorAction Stop
}

Write-Host "SUCCESS: The Microsoft Graph Beta for Teams Powershell Module are installed and connected, proceeding to next steps..." -Foreground Green

#======================= Lines 1-60 Added by Tim ==========================================================

#try 
#{
#    Write-Host "Loading Teams Beta PowerShell Module ...." -ForegroundColor "yellow"
#    Import-Module Microsoft.Graph.Beta.Teams
#}
#catch {
#    Write-Host "!****!!! Failed to Load Teams Beta PowerShell Module !****!!!" -ForegroundColor "Red"
#    Write-host "Run as administrator: Install-module Microsoft.Graph.Beta.Teams"
#}

#======================= Lines 74-82 hashed by Tim ==========================================================

$Report = [System.Collections.Generic.List[Object]]::new()
$i = 0

try {
Write-Host "Connecting to Graph....." -ForegroundColor "yellow"
Connect-MgGraph -Scopes "TeamworkDevice.Read.All , Directory.Read.All" -NoWelcome:$true	
}
catch {
    Write-Host "Connection Failed ....." -ForegroundColor "Red"
}

Write-Host "👌.....Connected....." -ForegroundColor "Green"
<#
Alternative: Use Application ID and Secured Password for authentication (you could also pass a certificate thumbnail)
$ApplicationId = "<applicationId>"
$SecuredPassword = "<securedPassword>"
$tenantID = "<tenantId>"

$SecuredPasswordPassword = ConvertTo-SecureString -String $SecuredPassword -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential
#>

[datetime]$RunDate = Get-Date
[string]$ReportRunDate = Get-Date ($RunDate) -format 'dd-MMM-yyyy HH:mm'
$Version = "0.1"
$CSVOutputFile = "MicrosoftTeamsRoomsReport.CSV"
$HtmlReportFile = "MicrosoftTeamsRoomsReport.html"
$JSONReportFile = "MicrosoftTeamsRoomsReport.json"
$uri = "https://graph.microsoft.com/beta/teamwork/devices?$filter=deviceType eq 'TeamsRoom'"

Write-Host "Getting a List of  MTR Devices" -ForegroundColor "Yellow"

# The possible values are DeviceType: 0 /unknown, 1/ipPhone, 2/teamsRoom, 3/surfaceHub,  4/collaborationBar, 5/teamsDisplay, 6/touchConsole, 7/lowCostPhone, 8/teamsPanel, 9/sip, 10/sipAnalog, 11/unknownFutureValue.

$allTeamsRoomDevices = Get-MgBetaTeamworkDevice
if ($allTeamsRoomDevices.count -ge 1){
 $colour="green"
}
elseif ($allTeamsRoomDevices.count -le 1) {
    $colour="Red"
    
}
Write-Host " Number of Found Devices (Count) :" $allTeamsRoomDevices.Count -ForegroundColor $colour

foreach($room in $allTeamsRoomDevices){

   $roomdetails= Get-MgBetaTeamworkDeviceHealth -TeamworkDeviceId $room.Id

    $ReportLine = [PSCustomObject][Ordered]@{  
        RoomName                   = $room.currentUser.displayName
        HealthStatus               = $room.HealthStatus
        activityState              = $room.activityState
        deviceType                 = $room.deviceType
        TeamworkDeviceId           = $room.id
        companyAssetTag            = $room.companyAssetTag
        ConnectionStatus           = $roomdetails.connection.connectionstatus
        loginStatus                = $roomdetails.loginStatus
        peripheralsHealth          = $roomdetails.peripheralsHealth
        softwareUpdateHealth       = $roomdetails.softwareUpdateHealth
        hardwareHealth             = $roomdetails.hardwareHealth
      }

    $Report.Add($ReportLine)
 $i++
}

# Value Explanations
#HealthStatus The possible values are: 0/unknown, 1/offline, 2/critical, 3/nonUrgent, 4/healthy, 5/unknownFutureValue.
#ActivityState The possible values are: 0/unknown,1/ busy, 2/idle, 3/unavailable, 4/unknownFutureValue.

Write-Host "Exporting Reports..."
$Report | Export-CSV -NoTypeInformation -path $CSVOutputFile -Encoding UTF8
$Report | ConvertTo-Html -Property RoomName,HealthStatus,activityState,deviceType,TeamworkDeviceId,companyAssetTag,ConnectionStatus,loginStatus,peripheralsHealth,softwareUpdateHealth,hardwareHealth | Out-File $HtmlReportFile
Start-Process $HtmlReportFile
$Report | ConvertTo-Json | Out-File "MicrosoftTeamsRoomsReport.json"
Write-Host ""
Write-Host "All done. Output files are in the chosen directory" $CSVOutputFile