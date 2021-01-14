#####################################################################################
#
#File: DailyChangeCheck.ps1
#Author: JianKun XU
#Version: 1.0
#
##  Revision History:
##  Date        Version    Alias       Reason for change
##  --------    -------   --------    ---------------------------------------
##  2/19/2019    1.0        JK          First version.
##     
#####################################################################################


Param (
    [Parameter(Mandatory=$true)]
    [String] $MailTo      
)

$separator = "="

# Initialize HTML table
$HtmlBody = "<TABLE Class='DailyChangeList' border='1' cellpadding='0'cellspacing='0' style='Width:1200px'>"

##===============================================================================================================

Write-Host $separator
Write-Host "Start to check Centeral Changes for last 48 hours of Content & Search." -ForegroundColor Green


$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='center'>DailyChangeList</td></TR>"
$HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold><TD colspan='2' align='center'>ChangeID</TD><TD align='center'>Starttime</TD><TD align='center'>Description</TD><TD align='center'>Count</TD></TR>"

$date = Get-Date
$checkpoint = $date.adddays(-2)

$farms = Get-CenteralFarm -Role content
$farms += Get-CenteralFarm -Role FederatedSearch
$farms += Get-CenteralFarm -Role SQL | ? state -ne 'Deleting'

$changes = $farms | % {Get-CenteralChange -Farm $_ -StartTimeLowerBound $checkpoint} | `
    ? description -NotLike '\\CHN.SPONETWORK.COM\*' | ? description -NotLike 'JobType=*SqlVm, Method=machinemgr*' | `
    ? description -NotLike 'Activity Lock*' | ? description -NotLike 'Patching Deployment Train*' | `
    ? description -NotLike 'DOU for Farm*' | ? description -NotLike 'InPlaceUpgrade for Farm*' | ? description -NotLike 'Tripwire MonitoringTargetFarm for Farm*'
$markedchanges = $changes |group description

foreach ($markedchange in $markedchanges)
{
    $ChangeID = ($markedchange | select -ExpandProperty group | select -First 1).Centeralchangeid
    $Starttime = ($markedchange | select -ExpandProperty group | select -First 1).starttime
    $Description = ($markedchange | select -ExpandProperty group | select -First 1).Description
    $Count = $markedchange.Count
   
    $HtmlBody += "<TR><TD colspan='2' align='center'>$ChangeID</TD><TD align='center'>$Starttime</TD><TD>$Description</TD><TD align='center'>$Count</TD></TR>"
}


#EmialProperty
$markdate=get-date -format MM/dd/yyyy
$MailSubject="SPO change check $markdate"


Send-Email -To $MailTo -mailbody $HtmlBody -mailsubject $MailSubject -BodyAsHtml -From "SPODailyCheck@21vianet.com"