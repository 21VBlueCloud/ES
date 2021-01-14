#####################################################################################
#

#   此脚本可以实现对域环境的多台服务器，通过远程的形式，来分析判断服务器的CPU健康情况。
#   此脚本可以对多台服务器使用，从而节省运维时间。
#>
                            
##
#####################################################################################

Param (
    [String] $xmlConfig="D:\Operation\Tools\System\DcMonitorConfig.xml",
    [String[]] $Module,
    [String[]] $Helper,
    [System.Management.Automation.PSCredential] $Credential
    
)

#Pre-loading
##==============================================================================================================

if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
}
else
{
    $Script:basePath = "."
}

. D:\Operation\Tools\Library\LogHelper.ps1
. D:\Operation\Tools\Library\CommonHelper.ps1
. D:\Operation\Tools\Library\RecordHelper.ps1
. D:\Operation\Tools\System\LocalSystemHelper.ps1


# Increase MaximumFunctionCount from 4096 to 8192 to avoid function loading failure
$MaximumFunctionCount  = 8192


#Create report folder to save all report files
$Folder = $Script:basePath + "\Reports"
if(!(test-path $folder))
{
    New-Item -Path $Folder -ItemType directory
}


#Load the POE management session.

Write-POELog "Check whether Management Session loaded...."

$Servers = Get-CentralAdminMachine ** 
if($Servers -eq $null -or $Servers.count -eq 0)
{
    Write-POELog "Load management session...."
    Load-ManagementShell 
}

#Import module to check connectivity

if ($Module -ne $Null) {
    Import-Module $Module
}
else
{
    Write-POELog "Load modules "
    & "D:\Operation\Tools\DailyCheck\Configuration\ReloadModule.ps1"
}

$StartTime = Get-Date
$Logdate = Get-Date -Format "yyyy-MM-dd"

#Load xml configuraion file
$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load($XmlConfig)

If ($Xml.HasChildNodes -eq $false) {
    Write-Host "Can not load config file `"$XmlConfig`"!"
    Break
}

<# Initialize HTML table
$HtmlBody = "<TABLE Class='CPUMonitoroing' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

##header
#Please replace the <HeaderName> with script Name or some words you are looking for
$TableHeader = "<HeaderName>"
$HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px>`
                    <TD colspan='6' align='center' style=color:$($CommonColor['White'])>$TableHeader</TD></TR>"
# Functions
#>

# Main code

$allServers = Get-ADComputer -Filter *  

#filter out the MM servers whose condition no  need to check

$MMServers  = $xml.DatacenterMonitoring.MMServers.Name

if(!(Test-Path $folder\GoodServers_$Logdate.csv))
{
    $GoodServers = Get-GoodCondtionServers -excludeServers $MMServers -logdate $logdate
}
else
{
    $GoodServers = Import-CSV $folder\GoodServers_$Logdate.csv
}


#Check CPU Rate Part

$checkpoint  = 1 
$threshold = $xml.DatacenterMonitoring.CPURate.LowerThreshold
$StartTime = $endtime = get-date
While($endtime -le ($StartTime.date.AddDays(1) - (New-TimeSpan -Minutes 3)))
{
    
    foreach($srv in $GoodServers)
    {
        Invoke-Command -ComputerName $srv -ScriptBlock ${function:Get-CPUUsage} -AsJob|Out-Null
    }
    $Jobs = Watch-Jobs -Activity "CPU Rate Checking checking..."
    $Jobs |Select TimeStamp,Server,CPUUsage| Export-Csv $folder\CPUUsage_$LogDate.csv -Append -NoTypeInformation
    Write-POELog "We'll gather the infomration 5 mins later"  
    #Sleep 900

    $checkpoint ++
    
    if($CheckPoint -gt 3)
    {
        $CPUHistory = Import-Csv $folder\CPUUsage_$logdate.csv  
        # Check the last 3 times result of the CPU rate.
        
        foreach($server in $GoodServers)
        {
            $alertDetails  = $CPUHistory|?{$_.Server -eq $server}|sort timestamp |select -Last 3
            
            $alertvalue  = 0
            foreach($value in $alertDetails)
            {
                if ([double]($value.CPUUsage.Split('')[0]) -gt [double]$threshold)
                {

                    Write-POELog "The CPU Rate is over $threshold for $server , please check the top 5 CPU usage process"

                    Write-POELog "Retrieve Top 5 CPU Process"
                    $T5Prs = Get-Top5ProcessUsage
                    $alertvalue +=1
                }
                   
            }
            #When all 3 checks meet the trigger alert value, trigger alert
            if($alertValue -eq 3)
            {
                
                $body = Format-HtmlTable -Contents $alertDetails -Title "CPU Rate for last 15 Mins"  
                $T5Prs = Get-Top5ProcessUsage
                $FormatPrResult = Format-HtmlTable -Contents $T5Prs -Title "Top 5 processes"
                $body += $FormatPrResult
                Send-mail -Subject "[Warning] SEV 2: CPU Usage is not healthy on $server" `
                            -Body $body  -to test@test.com
            
            }
        }
    }
    
}

##===============================================================================================================
#>
