#####################################################################################
#

#   此脚本可以实现对域环境的多台服务器，通过远程的形式，来分析判断服务器的内存健康情况。
#   此脚本可以对多台服务器使用，从而节省运维时间。
#
                            
##
#####################################################################################

<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER

.EXAMPLE
PS C:\>
#>

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
$script:Folder = $Script:basePath + "\Reports"
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




#filter out the MM servers whose condition no  need to check

$MMServers  = $xml.DatacenterMonitoring.MMServers.Name

# Check if connectivity check complete, and generate the good server files. If not, run connectivity check
if(!(Test-Path $folder\GoodServers_$Logdate.txt))
{
   
    $GoodServers = Get-GoodCondtionServers -excludeServers $MMServers  -logdate $logdate
}
else
{
    $GoodServers = Get-Content $folder\GoodServers_$Logdate.txt
}


#Check Meomory Rate Part

$checkpoint  = 1 
$threshold = $xml.DatacenterMonitoring.MemoryRate.LowerThreshold
#$threshold = 98
$StartTime = $endtime = get-date
While($endtime -le ($StartTime.date.AddDays(1) - (New-TimeSpan -Minutes 3)))
{
    
    foreach($srv in $GoodServers)
    {
        Invoke-Command -ComputerName $srv -ScriptBlock ${function:Get-MemoryUsage} -AsJob|Out-Null
    }
    $Jobs = Watch-Jobs -Activity "Meomory Rate Checking checking..."
    $Jobs |Select TimeStamp,Server,MemoryUsage| Export-Csv $folder\MemoryUsage_$LogDate.csv -Append -NoTypeInformation
    Write-POELog "We'll gather the infomration 5 mins later"  
    Sleep 900

    $checkpoint ++
    
    if($CheckPoint -gt 3)
    {
        $MeomoryHistory = Import-Csv $folder\MemoryUsage_$logdate.csv  
        # Check the last 3 times result of the Meomory rate.
        foreach($srv in $GoodServers)
        {
            $alertDetails  = $MeomoryHistory|?{$_.Server -eq $srv}|sort timestamp |select -Last 3
            
            $alertvalue  = 0
            foreach($value in $alertDetails)
            {
                if ([double]($value.MemoryUsage.Split('')[0]) -gt [double]$threshold)
                {

                    Write-POELog "The Meomory Rate is over $threshold for$srv , please check the top 5 Meomory usage process"

                    Write-POELog "Retrieve Top 5 Meomory Process"
                    $T5Prs =  Invoke-Command -ComputerName $srv -ScriptBlock ${function:Get-Top5ProcessUsage} -ArgumentList $true
                    $alertvalue +=1
                }
                   
            }
            #When all 3 checks meet the trigger alert value, trigger alert
            if($alertValue -eq 3)
            {
                
                $body = Format-HtmlTable -Contents $alertDetails -Title "Meomory Rate for last 15 Mins"  
                $T5Prs =  Invoke-Command -ComputerName $srv -ScriptBlock ${function:Get-Top5ProcessUsage } -ArgumentList $true
                $FormatPrResult = Format-HtmlTable -Contents $T5Prs -Title "Top 5 processes"
                $body += $FormatPrResult
                Send-mail -Subject "[Warning] SEV 2: Meomory Usage is not healthy on$srv" `
                            -Body $body  -to test@test.com
            
            }
        }
    }
    
}

##===============================================================================================================
#>
