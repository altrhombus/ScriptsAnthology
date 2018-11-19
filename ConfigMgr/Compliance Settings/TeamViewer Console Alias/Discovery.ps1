$header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$header.Add("authorization", 'Bearer YOUR-TEAM-VIEWER-API-KEY')

$pingTest = Invoke-RestMethod -Uri "https://webapi.teamviewer.com/api/v1/ping" -Method Get -Headers $header

if ($pingTest.token_valid -eq "True")
{
    $deviceList = Invoke-RestMethod -Uri "https://webapi.teamviewer.com/api/v1/devices" -Method Get -Headers $header
    $remoteControlId = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\TeamViewer" -Name "ClientID"
    $userAffinity = (Get-CimInstance -ClassName "CCM_UserAffinity" -Namespace "root\ccm\Policy\Machine\ActualConfig" -Property "ConsoleUser")
    $localAlias = "$env:COMPUTERNAME ($($userAffinity.ConsoleUser))"
    $thisDevice = $deviceList.devices | Where-Object { $_.remotecontrol_id -eq "r$remoteControlId"}
    $tvAlias = $thisDevice.alias

    if ($tvAlias -eq $localAlias) { $compliance = "true" }
    if ($tvAlias -ine $localAlias) { $compliance = "false" }

    $compliance
}