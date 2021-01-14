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
$HtmlBody = "<TABLE Class='DEPLOYBUILDORCH' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

Write-Host $separator
[Int] $OffsetDays = $Xml.DailyCheck.DeployBuildOrch.DayOffset

$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><TD colspan='6' align='center'>Deploy Build Orchestration</TD></TR>"
$HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold><TD>BuildVersion</TD><TD>Environment</TD><TD>ActivityType</TD><TD>CreateTime</TD><TD>ControlState</TD><TD>ActivityState</TD></TR>"

$date = (get-date).AddDays($OffsetDays)
$DeployBuildActivities = Get-CenteralDeploymentActivity -ActivityType deploybuild -Environment Gallatin -CreateTimeLowerBound $date | sort BuildVersion -Descending

foreach($a in $DeployBuildActivities)
{
    $version=$a.BuildVersion
    $Environment = $a.Environment
    $ActivityType =$a.ActivityType
    $CreateTime=$a.CreateTime
    $ControlState=$a.ControlState
    $ActivityState=$a.ActivityState

    if($ActivityState -ne 'Succeeded') {
        $HtmlBody += "<TR><TD>$version</TD><TD>$Environment</TD><TD>$ActivityType</TD><TD>$CreateTime</TD><TD>$ControlState</TD><TD style=background-color:#EA0000;font-weight:bold;color:white>$ActivityState</TD></TR>"
    }
    else {
        $HtmlBody += "<TR><TD>$version</TD><TD>$Environment</TD><TD>$ActivityType</TD><TD>$CreateTime</TD><TD>$ControlState</TD><TD style=background-color:#00A600;font-weight:bold;color:white>$ActivityState</TD></TR>"
    }
}

Write-Host "Checking for 'DeployBuildOrch' done." -ForegroundColor Green
Write-Host $separator



# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================