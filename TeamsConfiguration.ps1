try 
{
    Write-Host "Loading Teams Beta PowerShell Module ...." -ForegroundColor "yellow"
    Import-Module Microsoft.Graph.Beta.Teams
}
catch {
    Write-Host "!****!!! Failed to Load Teams Beta PowerShell Module !****!!!" -ForegroundColor "Red"
    Write-host "Run as administrator: Install-module Microsoft.Graph.Beta.Teams"
}


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
$CSVOutputFile = "c:\temp\TeamsRooms\MicrosoftTeamsRoomsReport.CSV"
$HtmlReportFile = "c:\temp\TeamsRooms\MicrosoftTeamsRoomsReport.html"

Write-Host "Getting a List of  MTR Devices" -ForegroundColor "Yellow"

#The possible values are DeviceType: 0 /unknown, 1/ipPhone, 2/teamsRoom, 3/surfaceHub,  4/collaborationBar, 5/teamsDisplay, 6/touchConsole, 7/lowCostPhone, 8/teamsPanel, 9/sip, 10/sipAnalog, 11/unknownFutureValue.

$allTeamsRoomDevices = Get-MgBetaTeamworkDevice #-filter "deviceType eq 'TeamsRoom'" 
if ($allTeamsRoomDevices.count -ge 1){
 $colour="green"
}
elseif ($allTeamsRoomDevices.count -le 1) {
    $colour="Red"
    
}
Write-Host " Number of Found Devices (Count) :" $allTeamsRoomDevices.Count -ForegroundColor $colour

foreach($room in $allTeamsRoomDevices){

   $roomdetails= Get-MgBetaTeamworkDeviceConfiguration -TeamworkDeviceId $room.Id

    $ReportLine = [PSCustomObject][Ordered]@{  
        TeamworkDeviceId                        = $room.id
        RoomName                                = $room.currentUser.displayName
        AdminAgentCurrentVersion                = $roomdetails.softwareVersions.adminAgentSoftwareVersion
        FirmwareSoftwareCurrentVersion          = $roomdetails.softwareVersions.partnerAgentSoftwareVersion
        FirmwareSoftwareAvailableVersion        = $roomdetails.softwareVersions.firmwareSoftwareVersion
        OperatingSystemSoftwareFreshness        = $roomdetails.softwareVersions.operatingSystemSoftwareVersion
        TeamsClientCurrentVersion               = $roomdetails.softwareVersions.teamsClientSoftwareVersion
        DisplayCount                            = $roomdetails.displayConfiguration.DisplayCount
        InBuiltDisplayScreenConfiguration       = $roomdetails.displayConfiguration.InBuiltDisplayScreenConfiguration
        IsDualDisplayModeEnabled                = $roomdetails.displayConfiguration.IsDualDisplayModeEnabled
        isContentDuplicationAllowed             = $roomdetails.displayConfiguration.isContentDuplicationAllowed
        CamerasDisplayName                      = $roomdetails.cameraConfiguration.Cameras.DisplayName
        CamerasCount                            = $roomdetails.cameraConfiguration.Cameras.Count
        IsContentCameraInverted                 = $roomdetails.CameraConfiguration.ContentCameraConfiguration.IsContentCameraInverted
        IsContentCameraOptional                 = $roomdetails.CameraConfiguration.ContentCameraConfiguration.IsContentCameraOptional
        IsContentEnhancementEnabled             = $roomdetails.CameraConfiguration.ContentCameraConfiguration.IsContentEnhancementEnabled
        IsLoggingEnabled                        = $roomdetails.SystemConfiguration.IsLoggingEnabled
        IsPowerSavingEnabled                    = $roomdetails.SystemConfiguration.IsPowerSavingEnabled
        IsScreenCaptureEnabled                  = $roomdetails.SystemConfiguration.IsScreenCaptureEnabled
        IsDeviceLockEnabled                     = $roomdetails.SystemConfiguration.IsDeviceLockEnabled
        IsAutoScreenShareEnabled                = $roomdetails.TeamsClientConfiguration.FeaturesConfiguration.IsAutoScreenShareEnabled
        IsBluetoothBeaconingEnabled             = $roomdetails.TeamsClientConfiguration.FeaturesConfiguration.IsBluetoothBeaconingEnabled
        IsHideMeetingNamesEnabled               = $roomdetails.TeamsClientConfiguration.FeaturesConfiguration.IsHideMeetingNamesEnabled
        IsSendLogsAndFeedbackEnabled            = $roomdetails.TeamsClientConfiguration.FeaturesConfiguration.IsSendLogsAndFeedbackEnabled
        EmailToSendLogsAndFeedback              = $roomdetails.TeamsClientConfiguration.FeaturesConfiguration.EmailToSendLogsAndFeedback
        SupportedClient                         = $roomdetails.TeamsClientConfiguration.AccountConfiguration.SupportedClient
        Domain                                  = $roomdetails.TeamsClientConfiguration.AccountConfiguration.OnPremisesCalendarSyncConfiguration.Domain
        DomainUserName                          = $roomdetails.TeamsClientConfiguration.AccountConfiguration.OnPremisesCalendarSyncConfiguration.DomainUserName
        SmtpAddress                             = $roomdetails.TeamsClientConfiguration.AccountConfiguration.OnPremisesCalendarSyncConfiguration.SmtpAddress
    }

    $Report.Add($ReportLine)
 $i++
}

# Value Explanations
#HealthStatus The possible values are: 0/unknown, 1/offline, 2/critical, 3/nonUrgent, 4/healthy, 5/unknownFutureValue.
#ActivityState The possible values are: 0/unknown,1/ busy, 2/idle, 3/unavailable, 4/unknownFutureValue.



Write-Host "Exporting Reports..."
$Report | Export-CSV -NoTypeInformation $CSVOutputFile -Encoding UTF8
Write-Host ""
Write-Host "All done. Output files are in the chosen directory" $CSVOutputFile