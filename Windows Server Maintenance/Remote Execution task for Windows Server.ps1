#
#此脚本可以在指定环境中运行。并远程分批次触发request,对指定的server打补丁， 或者运行特定的应用程序。
#该脚本可以实现邮件提醒，结果查询以及失败重试等功能。
#
#
##  Date      Version   Alias       Reason for change
##  --------  -------   --------    ---------------------------------------
##  11/7/2016   1.1       Fei      Update payload choosing logic 
##  12/6/2016   1.2       Fei      Update Patching sequence 
##  1/7/2017    1.3       Fei      Add mail notification 
##  1/30/2017   1.4       Fei      Update Error Throw actionpat
##  10/20/2017  1.5       Fei      Update POE Management session loading method
##  1/25/2018   1.6       Fei      Increase memory size to avoid powershell capacity limit issue.




if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
    $index = $Script:scriptname.LastIndOXEf("_")
    $scriptname = $scriptname.substring(0,$index)
}
else
{
    $Script:basePath = "."
}

# Increase memory size to avoid powershell capacity limit issue.
$MaximumFunctionCount = 8192
. D:\Operation\Tools\Library\LogHelper.ps1
. D:\Operation\Tools\Library\Patching_Helper_v2.2.ps1

#Load the POE management session.


Write-POELog "Check whether Management Session loaded...."

$Servers = Get-CentralAdminMachine ** 
if($Servers -eq $null -or $Servers.count -eq 0)
{
    Write-POELog "Load management session...."
    Load-ManagementShell 
}




# Load configuration.
 Write-POELog "Load XML Config "

$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load("\\protectioncn.gbl\sysvol\protectioncn.gbl\DropBox\OsPatching\Manifest.xml")

If ($Xml.HasChildNodes -eq $false) {
    Write-POELog "Can not load config file `"$XmlConfig`"!"
    Break
}

$date = get-date

#Start Stateless machine pataching

$payload = Get-CentralAdminDropBoxPayload -Filter "Name -like '*Patch*'"|sort CreateTime|select -Last 1

if($payload.Createtime -gt $date.AddDays(-1))
{
       if($payload.PayloadStatus -ne 'Success')
       {
            Monitor-StatelessPatching $payload
       }
       else
       {
          
           $Keywords_Color = @{}
           $Keywords_Color.add('Success','Green')
           Write-POELog "Payload $($payload.Name) complete"
           $payload = $payload|select Name,Identity,StartTime,EndTime,PayloadStatus,EscalationTeam,OwnerEmailAddress
           $body = Format-HtmlTable -Contents $payload -keywords_color $Keywords_Color
           if($Ismailsend -ne $true)
           {
                Send-Mail -Subject "POE-Patching_Payload_$($payload.Name)_Complete-$Logdate" -Body $body -to TestUser@TEST.COM 
           }
           $Ismailsend = $true
       }
}  
else      
{

    Write-POELog "Check Machine DesiredVersion & ActualVersion"
    $servers = get-centraladminmachine **
    $VIssueServers = $servers|?{$_.ActualVersion.RawIdentity -ne $_.DesiredVersion.RawIdentity}
    if($VIssueServers.count -ne 0)
    {
        Write-POELog " $($VIssueServers.count) servers actualversion & DesiredVersion are not match,Check Detail"
        $VIssueServers = $VIssueServers|select Name,ActualVersion,DesiredVersion
    }

    $Keywords_Color = @{}
    $Keywords_Color.add('ActualVersion','Green')
    $body = Format-HtmlTable -Contents $VIssueServers -keywords_color $Keywords_Color
    Send-Mail -Subject "POE-Patching_Issue_Below_Servers_version_Mismatched-$Logdate" -Body $body -to TestUser@TEST.COM

    Write-POELog "Update Machine's Actualversion to match its DesiredVersion" 
    $VIssueServers|?{$_.Name -notlike "*ooe" -and $_.Name -notlike "*ADS0*"}|%{Set-CentralAdminMachine $_.Name -ActualVersion $_.DesiredVersion}
    
    Write-POELog "Check AM Register Key setting"
    & "D:\Operation\Tools\Patching_Automation\Fix-MissingRegesiterKey.ps1"

    Patch-Stateless
    $payload = Get-CentralAdminDropBoxPayload -Filter "Name -like '*Patch*'"|sort CreateTime|select -Last 1
    Monitor-StatelessPatching  $payload
}

Write-POELog "Stateless Patch $($payload.name) complete, set version back "
if($VIssueServers -ne $Null)
{
    $VIssueServers|?{$_.Name -notlike "*ooe" -and $_.Name -notlike "*ADS0*"}|%{Set-CentralAdminMachine $_.Name -ActualVersion $_.ActualVersion}
}

# Start CWT machine pataching
Write-host "Start Patching  CWT"
Patch-CWT -DATE $DATE

# # Start QDS machine pataching
Write-POELog "CWT Patch complete, Start Patching QDS"
Patch-QDS -DATE $DATE
      
#Patch-CDB -date $date
Write-POELog "QDS Patch complete,Start Patching CDB"

Patch-CDB -DATE $DATE

Write-host "The last role CDB Patch complete, $(get-date -Format MMM) Patching is totally complete"

