#.Synopsis
# The script is used to monitor the change, consists of Deployment, patch payload, Rack, Tenant, flight and so on
# Authored by Eric Yuan

#Load management session#
#& 'C:\LocalFiles\Exchange\Datacenter\New-ManagementSession-V2.ps1' -Target "bjb" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

&'C:\Torus\TorusClient.ps1' -Mode connect -ServiceInstance Gallatin -UserType GallatinOperator -DMS 

$CommonColor = @{
    Red = "#EA0000"
    Cyan = "#00FFFF"
    Blue = "#0066CC"
    purple = "#800080"
    Yellow = "#FFFF00"
    Lime = "#00FF00"
    Magenta = "#FF00FF"
    White = "#FFFFFF"
    Silver = "#C0C0C0"
    Gray = "#808080"
    Black = "#000000"
    Orange = "#FFA500"
    Brown = "#A52A2A"
    Maroom = "#800000"
    Green = "#00A600"
    Olive = "#808000"
    Pink = "#FFC0CB"
    DarkBlue = "#00008B"
    DarkCyan = "#008B8B"
    DarkGray = "#A9A9A9"
    DarkGreen = "#006400"
    DarkMagenta = "#8B008B"
    DarkOrange = "#FF8C00"
    DarkRed = "#8B0000"
    DeepPink = "#FF1493"
    LightBlue = "#ADD8E6"
    LightCyan = "#E0FFFF"
    LightGray = "#D3D3D3"
    LightGreen = "#90EE90"
    LightYellow = "#FFFFE0"
    MediumBlue = "#0000CD"
    MediumPurple = "#9370DB"
    YellowGreen = "#9ACD32"
    DailmondBlue = "#AFEEEE"
}

$date=(get-date).AddDays(-7)
$date_month = (get-date).AddDays(-30)
$date_day=(get-date).AddDays(-1)
$d=(Get-Date).Date
$date3=(get-date).AddDays(-3)

#$v2Workflow= $v2Workflow_Month |?{$_.Time -gt $date}

########Get Data from Management Session##########
$DeployPayload = Get-CentralAdminDropBoxPayload * |?{$_.name -match 'DeploymentTrainSweeper_Gallatin'}|sort CreateTime  -Descending
$SiteMMHistory = Get-SiteMaintenanceHistory |?{$_.WhenCreated -gt $date}|sort WhenCreated  -Descending
$SiteMMHistory_Month = Get-SiteMaintenanceHistory |?{$_.WhenCreated -gt $date_month}|sort WhenCreated  -Descending
$PatchPayload = Get-CentralAdminDropBoxPayload * |?{$_.name -match 'Patch'}|sort CreateTime  -Descending
$Build = (Get-CentralAdminMachine *MB* | group actualversion |sort Name -Descending | select -First 1).name
$SiteMMIn3days = $SiteMMHistory |?{$_.WhenCreated -gt $date3}
$ActiveDeployPayload= $DeployPayload |?{$_.PayloadStatus -eq 'Started'}
$SiteMMRecent = $SiteMMHistory


$OldFEs = (get-centraladminrack |?{$_.type -match 'PreRackMTCapacityFrontEnd'}).name
$NewFEs = (get-centraladminrack |?{$_.type -match 'MTCapacityFE_C'}).name
# $a.name.Contains('SHA20F02C01-AN320-L')
#Load capacity session#
Get-PSSession |Remove-PSSession 
#&'C:\LocalFiles\Exchange\Datacenter\ConnectRPStoManandCap.ps1'  -Forest CHNPR01.prod.partner.outlook.cn -ErrorAction Continue -Version 15
<#
#Tenant check
$DualWrite_Tenants = @()
$Tenants = @()
$result = ""
$Tenants= Get-Organization -Filter "WhenCreated -gt '$($date_day)'" -AccountPartition 'CHNPR01a001.prod.partner.outlook.cn'|?{$_.Name -match 'onmschina'}
$DualWrite_Tenants = $Tenants |?{$_.IsDualWriteEnabled -eq 'True'}
If($DualWrite_Tenants -eq $null -and $Tenants -ne $null)
{$result = 'Dual-Wirte is disabled'}
Elseif($DualWrite_Tenants -ne $null -and $Tenants -ne $null)
{$result = 'Warning:Dual-Wirte feature of below tenants is enabled'}

#Flight#
$FlightChange =  Get-SettingOverride |?{$_.WhenCreated -gt $date} |sort WhenCreated -Descending
#$FlightChange|Select-Object ID,ModifiedBy,Reason,WhenCreated,ComponentName,SectionName,MinVersion,FixVersion,MaxVersion|Export-Csv 1.csv
$FlightChange_month = Get-SettingOverride |?{$_.WhenCreated -gt $date_month} |sort WhenCreated -Descending
#>
$MailBody="<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:800px'>"
$MailBody+="<TR style=background-color:$($CommonColor['Blue']);color:$($CommonColor['White']);font-weight:bold;font-size:20px><TD colspan='9' align='center'>Change Monitor</TD></TR>"
<#
$MailBody+="<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='9' align='left'>New Tenant Monitor ($(($Tenants| measure).count) Tenant was provisioned in last 24 hours. $($result))</TD></TR>"
#$MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='2' align='center'>TenantName</TD><TD colspan='2' align='center'>IsDualWriteEnabled</TD><TD colspan='5' align='center'>WhenCreated</TD></TR>"
If($DualWrite_Tenants -ne $null)
{
 $MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='2' align='center'>TenantName</TD><TD colspan='2' align='center'>IsDualWriteEnabled</TD><TD colspan='5' align='center'>WhenCreated</TD></TR>"
 foreach($Tenant in $DualWrite_Tenants)
  {
   $MailBody+="<TR style=font-size:15px><TD colspan='2' align='Left'>$($Tenant.Name) </TD><TD colspan='2' align='center'>$($Tenant.IsDualWriteEnabled)</TD><TD colspan='5' align='center'>$($Tenant.WhenCreated)</TD></TR>"
  }
}
<#
Else
{
   $MailBody+="<TR style=font-size:15px><TD colspan='9' align='Left'>No new Tenants in 24 hours</TD></TR>"
}
#>

$MailBody+="<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='9' align='left'>Deployment Payload Monitor ($($DeployPayload.count) payloads in a week)(Latest Build version:$($Build))</TD></TR>"
If($ActiveDeployPayload -ne $null)
{
$MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='2' align='center'>Deployment</TD><TD colspan='2' align='center'>Payload ID</TD><TD colspan='2' align='center'>Payload State</TD><TD colspan='3' align='center'>WhenCreated</TD></TR>"
 foreach($payload in $ActiveDeployPayload)
  {
   $MailBody+="<TR style=font-size:15px><TD colspan='2' align='Left'>Active Deployment payload </TD><TD colspan='2' align='center'>$($ActiveDeployPayload.Name)</TD><TD colspan='2' align='center'>$($ActiveDeployPayload.PayloadStatus)</TD><TD colspan='3' align='center'>$($ActiveDeployPayload.LastProcessTime)</TD></TR>"
  }
}
Else
{
   $MailBody+="<TR style=font-size:15px><TD colspan='2' align='Left'>No Active Depploy Payload. The latest payload is </TD><TD colspan='2' align='center'>$($DeployPayload[0].Name)</TD><TD colspan='2' align='center'>$($DeployPayload[0].PayloadStatus)</TD><TD colspan='3' align='center'>$($DeployPayload[0].LastProcessTime)</TD></TR>"
}
<#
$MailBody+="<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='9' align='Left'>Flight Change Monitor ($(($FlightChange| measure).count) in a week and $($FlightChange_month.count) in a month)</TD></TR>"
If($FlightChange -ne $null)
{
 $MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='1' align='center'>Flight ID</TD><TD colspan='1' align='center'>Modified By</TD><TD colspan='1' align='center'>Reason</TD><TD colspan='1' align='center'>Time</TD><TD colspan='1' align='center'>ComponentName</TD><TD colspan='1' align='center'>SectionName</TD><TD colspan='1' align='center'>MinVersion</TD><TD colspan='1' align='center'>FixVersion</TD><TD colspan='1' align='center'>MaxVersion</TD></TR>"
 Foreach($Flight in $FlightChange)
 {
 $MailBody+="<TR style=font-size:15px><TD colspan='1' align='center'>$($Flight.Id)</TD><TD colspan='1' align='center'>$($Flight.ModifiedBy)</TD><TD colspan='1' align='center'>$($Flight.Reason)</TD><TD colspan='1' align='center'>$($Flight.WhenCreated)</TD><TD colspan='1' align='center'>$($Flight.ComponentName)</TD><TD colspan='1' align='center'>$($Flight.SectionName)</TD><TD colspan='1' align='center'>$($Flight.MinVersion)</TD><TD colspan='1' align='center'>$($Flight.FixVersion)</TD><TD colspan='1' align='center'>$($Flight.MaxVersion)</TD></TR>"
 }
}
#>
#PatchPayload
$MailBody+="<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='9' align='left'>Patch Payload Monitor ($($PatchPayload.count) payloads in a week)</TD></TR>"
If($PatchPayload -ne $null)
{
 $MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='2' align='center'>Payload ID</TD><TD colspan='2' align='center'>UserCreated</TD><TD colspan='2' align='center'>EscalationTeam</TD><TD colspan='3' align='center'>WhenCreated</TD></TR>"
 Foreach($Patch in $PatchPayload)
 {
   $MailBody+="<TR style=font-size:15px><TD colspan='2' align='Left'>$($Patch.Name) </TD><TD colspan='2' align='center'>$($Patch.UserCreated)</TD><TD colspan='2' align='center'>$($Patch.EscalationTeam)</TD><TD colspan='3' align='center'>$($Patch.CreateTime)</TD></TR>"
 }
}

$MailBody+="<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='4' align='left'>Rack Change Monitor($($SiteMMHistory.count) in a week and $($SiteMMHistory_Month.count) in a month)</TD></TR>"
$MailBody+="<TR style=font-weight:bold;font-size:15px><TD colspan='2' align='center'>Rack(3 days)</TD><TD colspan='2' align='center'>Requester</TD><TD colspan='2' align='center'>Workflow Type</TD><TD colspan='3' align='center'>WhenChanged</TD></TR>"


Foreach($SiteMM in $SiteMMIn3days)
{
$Rack = (($SiteMM.FromSite).Split('\'))[1]
If($OldFEs.Contains($Rack))
{
$MailBody+="<TR style=background-color:$($CommonColor['Red']);font-size:15px><TD colspan='2' align='Left'>VIP_Rack: $($Rack)</TD><TD colspan='2' align='center'>$($SiteMM.Requester)</TD><TD colspan='2' align='center'>$($SiteMM.Workflow)</TD><TD colspan='3' align='center'>$($SiteMM.WhenChanged)</TD></TR>"
}
Elseif($NewFEs.Contains($Rack))
{
$MailBody+="<TR style=background-color:$($CommonColor['Yellow']);font-size:15px><TD colspan='2' align='Left'>New_VIP_Rack: $($Rack)</TD><TD colspan='2' align='center'>$($SiteMM.Requester)</TD><TD colspan='2' align='center'>$($SiteMM.Workflow)</TD><TD colspan='3' align='center'>$($SiteMM.WhenChanged)</TD></TR>"
}
Else
{
$MailBody+="<TR style=font-size:15px><TD colspan='2' align='Left'>$($SiteMM.FromSite)</TD><TD colspan='2' align='center'>$($SiteMM.Requester)</TD><TD colspan='2' align='center'>$($SiteMM.Workflow)</TD><TD colspan='3' align='center'>$($SiteMM.WhenChanged)</TD></TR>"
}

}

$MailBody+="</TABLE>"


$mbCred = new-object Management.Automation.PSCredential “EXO-EOP-TeamBox@oe.21vianet.com”, (“Oe.20202020.4” | ConvertTo-SecureString -AsPlainText -Force)

$SmtpServer="mail.21vianet.com"
$from = "EXO-EOP-TeamBox@oe.21vianet.com"
$to = "O365_EXO@oe.21vianet.com"

Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject "Change Monitor Report" -BodyAsHtml -Credential $mbCred 
<#
$credentials = new-object Management.Automation.PSCredential “Admin@exotest.partner.onmschina.cn”, (“Abcd1234” | ConvertTo-SecureString -AsPlainText -Force)
#$recipient="O365_EXO@oe.21vianet.com"
$recipient="yuan.jinliang@oe.21vianet.com"

Send-mailmessage -Body $MailBody -BodyAsHtml -smtpServer smtp.partner.outlook.cn -to $recipient -from "Admin@exotest.partner.onmschina.cn" -subject "Change Monitor Report" -credential $credentials -UseSsl
#>
