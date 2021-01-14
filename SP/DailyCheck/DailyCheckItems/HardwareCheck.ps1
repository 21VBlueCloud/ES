#####################################################################################
#
#File: HardwareCheck.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  11/29/2017   1.1       Wind       Add unreachable PMs inspection.
##
##  5/9/2018     1.2       JK         Add send mail to ES pooling mailbox.
##
##  8/11/2019    1.3       Wind       Change mail sending block to comply if no mailto
##                                    in configure file.
##                                    
##  9/9/2019     1.4       Wind       Add credential to query DSE hosted PMs.
##
##  12/5/2019    1.5       Wind       Exclude new capacity check for zone 78 & 80, because
##                                    PMs in there are not HP.
#####################################################################################

Param (
    [Parameter(Mandatory=$true)]
    [String] $XmlConfig,
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
$HtmlBody = "<TABLE Class='HARDWARECHECK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

# Functions

# Main code

# HTML header
$TableHeader = "Hardware Check"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD colspan='5' align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"
<#
Procedure:
1. Get all online PMs from Centeral.
2. Ping test for getting pingable PMs and record non-pingable PMs.
3. Retrieve System health state summary by Get-HPComputerSystem.
4. Pick unhealth PMs and get details error by Get-HWHealthStatus.
5. Outputs non-pingable PMs and unhealth PMs.
#>


# Step 1: Get PMs
## Update 1.5: Exclude zone 78 & 80.
$Zones = 6, 7, 12, 13 | Get-CenteralZone
$PMs = $Zones | %{Get-CenteralPMachine -Zone $_} | ? State -NE "Decommissioned" | select -ExpandProperty Name

## Update 1.4: Destinguish DSE hosts.
$Networks = $Zones | %{Get-CenteralNetwork -Zone $_}
$DSE = $Networks | %{Get-CenteralVM -Role DSE -State Running -Network $_}
$DSEPMs = $DSE | Group-Object PmachineId -NoElement | %{Get-CenteralPMachine $_.Name} | select -ExpandProperty Name

$PMs = Compare-Object $PMs $DSEPMs | select -ExpandProperty InputObject

# Step 2: Ping test (Deprecated because Test-MachineConnectivity performance is low)
<#
$PingResult = Test-MachineConnectivity -ComputerName $PMs
$UnpingPMs = @($PingResult | ? Ping -EQ $false | select -ExpandProperty ComputerName)
$PingablePMs = @($PingResult | ? Ping -EQ $true | select -ExpandProperty ComputerName)
#>
$PingablePMs = $PMs

# Step 3: Retrieve system health summary
if ($PingablePMs) {
    $HealthSummary = @(Get-HPComputerSystem -ComputerName $PingablePMs)
    $HealthSummary += Get-HPComputerSystem -ComputerName $DSEPMs -Credential $Credential
    # check health state
        #    0  (Unknown)
        #    5  (OK)
        #    10 (Degraded)
        #    20 (Major Failure)
    $BadPMs = @($HealthSummary | ? HealthState -NE 5 | select -ExpandProperty PSComputerName)

}

# Update 1.1: Add a comparison between PMs and Health summary to get which PMs are not pingable.
$UnpingPMs = Compare-Object $PMs $HealthSummary.PSComputerName | ? SideIndicator -EQ "<=" | select -ExpandProperty InputObject


#Step 4: Get details for bad PMs
if ($BadPMs) {
    $Str = Convert-ArrayAsString $BadPMs
    Write-Host "bad PMs:"
    Write-Host $Str
    Write-Host "Run 'Get-HWHealthStatus -ComputerName $Str' to get details." -ForegroundColor Yellow

    ## update 1.4: Distinguish DSE hosts.
    $BadDSEPMs = @(Compare-Object $BadPMs $DSEPMs -ExcludeDifferent -IncludeEqual | Select-Object -ExpandProperty InputObject)
    $BadPMs = @(Compare-Object $BadPMs $BadDSEPMs | Select-Object -ExpandProperty InputObject)

    $BadInfo = @(Get-HWHealthStatus -ComputerName $BadPMs -PassThru | sort ComputerName)
    if ($BadDSEPMs.Count -gt 0) {
        $BadInfo = Get-HWHealthStatus -ComputerName $BadDSEPMs -PassThru -Credential $Credential | sort ComputerName
    }
}

# Step 5: Outputs
if ($BadInfo) {
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold;font-size:17px>`
                    <TD align='center'>ComputerName</TD><TD align='center'>SerialNumber</TD><TD align='center'>Parts</TD>`
                    <TD align='center'>Name</TD><TD align='center'>HealthState</TD></TR>"
    foreach ($info in $BadInfo) {
        $AlertColor = "Yellow"
        if ($info.HealthState -in "Major Failure","Error") { $AlertColor = "Red" }
        $HtmlBody += "<TR style=background-color:$($CommonColor[$AlertColor]);font-size:17px>`
                        <TD>$($info.ComputerName)</TD><TD>$($info.SerialNumber)</TD><TD>$($info.Parts)</TD>`
                        <TD>$($info.Name)</TD><TD>$($info.HealthState)</TD></TR>"
    }

}
else {
    # No errors
    $HtmlBody += "<TR style=font-weight:bold;font-size:17px><TD colspan='5' align='left' style=color:$($CommonColor['Green'])>`
                    None</TD></TR>"
}

# Update 1.1: Add unpingable PMs in mail content.
if ($UnpingPMs) {
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold;font-size:17px>`
                    <TD colspan='5' align='left'>Below PMs cannot connect:</TD></TR>"

    $HtmlBody += "<TR style=font-weight:bold;font-size:17px><TD colspan='5' align='left' style=color:$($CommonColor['Red'])>`
                    $(Convert-ArrayAsString $UnpingPMs)</TD></TR>"
}

Write-Host "Checking for '$TableHeader' done." -ForegroundColor Green

#Emial Properties
$MailServer = $Xml.DailyCheck.Common.SmtpServer

$MailSubject = "SPO Daily Check V6 `"HardwareCheck`" on $((Get-Date).ToString("yyyyMMddhhmmss")) by $env:username"
$MailSender = $Xml.DailyCheck.Common.MailSender
$MailTo = $Xml.DailyCheck.HardwareCheck.MailTo

## Update 1.3:
if ($MailTo) {
    Send-Email -To $MailTo -mailsubject $MailSubject -From $MailSender -mailbody $HtmlBody -SmtpServer $MailServer -BodyAsHtml
}

Write-Host $separator
# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================