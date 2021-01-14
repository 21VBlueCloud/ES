#Load　module

#$MGT03_DC= Get-CentralAdminMachine *mgt03*DC* |?{$_.ProvisioningState -eq 'Provisioned'}
$PR01_DC= Get-CentralAdminMachine *pr01DC* |?{$_.ProvisioningState -eq 'Provisioned'}
#$PR01A001_DC= Get-CentralAdminMachine *pr01A001DC* |?{$_.ProvisioningState -eq 'Provisioned'}


$Problematic_DC=@()
$All_DC=@()
<#
foreach($MGT in $MGT03_DC)
{

$output= Invoke-Command -ComputerName $MGT.name -ScriptBlock {w32tm /query /status}
$Msg= $output|Out-String
$startIndex = $Msg.Indexof("Delay: ")
$LastIndex = $Msg.Indexof("Root Dispersion")
$delay_time= $Msg.Substring($startIndex+7,$LastIndex-$startIndex-7)

  $DC = get-centraladminmachine $MGT.Name | select name
  $DC| Add-Member -MemberType NoteProperty -Name 'Delay_Time' -Value $delay_time 
  <#
  If($DC.Delay_Time -ge 1)
  {
 $Problematic_DC+=$DC
  }
  
  $All_DC+=$DC
}
#>
foreach($PR01 in $PR01_DC)
{
$Name= $PR01.name + ".CHNPR01.prod.partner.outlook.cn"
$output= Invoke-Command -ComputerName $Name -ScriptBlock {w32tm /query /status}
$Msg= $output|Out-String
$startIndex = $Msg.Indexof("Delay: ")
$LastIndex = $Msg.Indexof("Root Dispersion")
$delay_time= $Msg.Substring($startIndex+7,$LastIndex-$startIndex-7)

  $DC = get-centraladminmachine $PR01.Name | select name
  $DC| Add-Member -MemberType NoteProperty -Name 'Delay_Time' -Value $delay_time 
  If($DC.Delay_Time -ge 1)
  {
  $Problematic_DC+=$DC
  }
 $All_DC+=$DC

}

foreach($PR01A001 in $PR01A001_DC)
{
$Name= $PR01A001.name + ".CHNPR01A001.prod.partner.outlook.cn"
$output= Invoke-Command -ComputerName $Name -ScriptBlock {w32tm /query /status}
$Msg= $output|Out-String
$startIndex = $Msg.Indexof("Delay: ")
$LastIndex = $Msg.Indexof("Root Dispersion")
$delay_time= $Msg.Substring($startIndex+7,$LastIndex-$startIndex-7)

  $DC = get-centraladminmachine $PR01.Name | select name
  $DC| Add-Member -MemberType NoteProperty -Name 'Delay_Time' -Value $delay_time 
  If($DC.Delay_Time -ge 1)
  {
  $Problematic_DC+=$DC
  }
 $All_DC+=$DC

}

$MailBody="<TABLE border='0' cellpadding='0'cellspacing='0' style='Width:800px'>"

<#
$Problematic_DC=@()
$DC = get-centraladminmachine SHAPR01DC003 | select name
$delay_time = 0.9
$DC| Add-Member -MemberType NoteProperty -Name 'Delay_Time' -Value $delay_time 
If($DC.Delay_Time -ge 1)
{
$Problematic_DC+=$DC
}
#>

$MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='Center'>Scanned $($MGT03_DC.Count) CHNMGT03 DC;$($PR01_DC.Count) CHNPR01 DC and $($PR01A001_DC.Count) CHNPR01A001 DC</TD></TR>"
$MailBody+="</Table>"
$MailBody+="<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:800px'>"
If($Problematic_DC.count -eq 0)
{
  $MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='1' align='Left'>DC Time Sync Status</TD><TD colspan='1' bgcolor='#00FF00' align='center'>Healthy</TD></TR>"

}
Else
{
  $MailBody+="<TR style=background-color:Orange;font-weight:bold;font-size:17px><TD colspan='1' align='Left'>DC Time Sync Status</TD><TD colspan='1' align='center'>Unhealthy</TD></TR>"
  $MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='Left'>The below are the DCs with the sync time delay above 1s</TD></TR>"
  foreach ($DC_Server in $Problematic_DC)
  {
  $MailBody+="<TR style=background-color:Orange;font-weight:bold;font-size:17px><TD colspan='1' align='Left'>$($Problematic_DC.Name)</TD><TD colspan='1' align='center'>$($Problematic_DC.Delay_Time)</TD></TR>"
  }
}
$MailBody+="</Table>"

#Check MB Version#

$version=get-centraladminmachine *MB* |group actualversion

$MailBody+="<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:800px'>"
$MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='center'>Build Version Of MB Servers</TD></TR>"
foreach($i in $version)
{
$MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='1' align='Left'>$($i.name)</TD><TD colspan='1' align='center'>$($i.count)</TD></TR>"
}
$MailBody+="</Table>"


$credentials = new-object Management.Automation.PSCredential “account”, (“psw” | ConvertTo-SecureString -AsPlainText -Force)
$recipient="recipient"



Send-mailmessage -Body $MailBody -BodyAsHtml -smtpServer smtp -to $recipient -from "sender" -subject "Monitor: DC Time Sync Delay Check" -credential $credentials -UseSsl 


