#####################################################################################
#
#File: 
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  9/19/2016   1.0       Wind        First version.
##                                    
##  6/19/2017   4.5       Wind        Add check preqrequisite check with deployment activity.
##
##  8/27/2018   4.6       Wind        Add comparison between Gallain and DProd. Report error
##                                    if Gallatin GM Orch are latter than DProd more.
#####################################################################################


Param (
    [Parameter(Mandatory=$true)]
    [String] $xmlConfig,
    [String[]] $Module,
    [String[]] $Helper,
    [System.Management.Automation.PSCredential] $Credential
    
)

#Pre-loading
##==============================================================================================================
# Import module
if ($Module) {
    Import-Module $Module
}

#Load helper
$helper | %{ . $_ }

#Load xml configuraion file
$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load($XmlConfig)

If ($Xml.HasChildNodes -eq $false) {
    Write-Host "Can not load config file `"$XmlConfig`"!"
    Break
}

$separator = "=" * $Host.UI.RawUI.WindowSize.Width

# Initialize HTML table
$HtmlBody = "<TABLE Class='GMORCH' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

Write-Host $separator
[Int] $OffsetDays = $Xml.DailyCheck.GMOrch.DayOffset

$HtmlBody += "<TR style=background-color:$($CommonColor['Blue']);font-weight:bold;font-size:17px>"`
                + "<TD colspan='6' align='center' style=color:$($CommonColor['White'])>Centeral Manager Orchestration</TD></TR>"

#Check GM Version
$GMFarms =Get-CenteralFarm -Role CenteralManager
$PRFarm = $GMFarms | ? RecoveryFarmId -NE 0
$gmversion=$PRFarm.version
$HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>CenteralManagerVersion : $gmversion</TD></TR>"

# 4.4 update: Check UpgradeGM prerequisite
$GMUpgradeBakeTime = Get-CenteralRelationship -Role GMUpgradeBakeTime

#GM upgrade checking
$date = (get-date).AddDays($OffsetDays)
$UpgradeGMActivities = Get-CenteralDeploymentActivity -ActivityType UpgradeGM -Environment Gallatin -CreateTimeLowerBound $date | sort BuildVersion -Descending

## 4.6 update: Add comparison with DProd activities.
$DProdGMActivities = Get-CenteralDeploymentActivity -ActivityType UpgradeGM -Environment DProd -CreateTimeLowerBound $date | sort BuildVersion -Descending

# compare with DProd. Report error if their BuildVersions are different.
if ($UpgradeGMActivities) {
    $LastGallatinActivity = $UpgradeGMActivities[0]
    $LastDProdActivity = $DProdGMActivities[0]

    # Gallatin version fell behind the Dprod version.
    if($LastGallatinActivity.BuildVersion -lt $LastDProdActivity.BuildVersion) {
        $ReportError = "Gallatin UpgradeGM [version=$($LastGallatinActivity.BuildVersion), State=$($LastGallatinActivity.ActivityState)] `
            behind the Dprod [version=$($LastDProdActivity.BuildVersion), State=$($LastDProdActivity.ActivityState)]!"
    }

    # Gallatin is advance to DProd. This situation is not normal and it's very rare.
    if($LastGallatinActivity.BuildVersion -gt $LastDProdActivity.BuildVersion) {
        $ReportWarning = "Gallatin UpgradeGM [version=$($LastGallatinActivity.BuildVersion), State=$($LastGallatinActivity.ActivityState)] `
            fell behind the Dprod [version=$($LastDProdActivity.BuildVersion), State=$($LastDProdActivity.ActivityState)]!"
    }

}
else {
    $ReportError = "No any UpgradeGM orch in $([Math]::Abs($OffsetDays)) days!"
}

# 4.5 update: check deployment activity prerequisite if Centeral relationship failed.
if ($GMUpgradeBakeTime) {
    $Prerequisite = $GMUpgradeBakeTime.State
}
else {
    $ActivityPrereq = Get-CenteralPrerequisiteDeploymentActivity -ActivityType UpgradeGM -Environment Gallatin -BuildVersion $UpgradeGMActivities[0].BuildVersion
    $Prerequisite = $ActivityPrereq.Environment
}

$HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>UpgradeGM prerequiste is: "
If ($Prerequisite -eq "ProdBubble") {
    $HtmlBody += "$($Prerequisite)</TD></TR>"
}
Else {$HtmlBody += "<span style=background-color:Yellow>$($Prerequisite)</span></TD></TR>"}


# $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='left'>CenteralManager Upgrade Status</td></TR>"
$HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold><TD>BuildVersion</TD><TD>Environment</TD><TD>ActivityType</TD><TD>CreateTime</TD><TD>ControlState</TD><TD>ActivityState</TD></TR>"


foreach ($GMactivity in $UpgradeGMActivities)
{
    $GMAversion=$GMactivity.BuildVersion
    $GMAEnvironment = $GMactivity.Environment
    $GMAActivityType =$GMactivity.ActivityType
    $GMACreateTime=$GMactivity.CreateTime
    $GMAControlState=$GMactivity.ControlState
    $GMAActivityState=$GMactivity.ActivityState

    if($GMAActivityState -eq 'Succeeded') {
        $HtmlBody += "<TR><TD>$GMAversion</TD><TD>$GMAEnvironment</TD><TD>$GMAActivityType</TD><TD>$GMACreateTime</TD><TD>$GMAControlState</TD><TD style=background-color:#00A600;font-weight:bold;color:white>$GMAActivityState</TD></TR>"
    }
    else {
        $HtmlBody += "<TR><TD>$GMAversion</TD><TD>$GMAEnvironment</TD><TD>$GMAActivityType</TD><TD>$GMACreateTime</TD><TD>$GMAControlState</TD><TD style=background-color:#EA0000;font-weight:bold;color:white>$GMAActivityState</TD></TR>" 
    }
}

if ($ReportError) {
    $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);font-weight:bold;font-size:17px><TD colspan='6'" `
                + "align='left' style=color:$($CommonColor['White'])>$ReportError</TD></TR>"
}

if ($ReportWarning) {
    $HtmlBody += "<TR style=background-color:$($CommonColor['Yellow']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>" `
                + "$ReportWarning</TD></TR>"
}

Write-Host "Checking for 'GMOrch' done." -ForegroundColor Green
Write-Host $separator

# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================