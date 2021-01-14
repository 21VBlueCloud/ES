#
#此脚本可以在适合的云服务或者云系统中，实现重启Sql 服务的功能。
#此脚本可以通过在指定的命令窗口中， 实现服务器的分类，状态查询以及后续的SQL 服务的检查及重启功能。
#
#

Function Get-ProvisionedDBServers
{   
    try
    {
        $DBservers = @(Get-CentralAdminMachine *DB*)
    }
    catch
    {
        Write-EOPLog -Message $_.Exception.Message -Level Error       
    }

    $Servers = @($DBservers|?{$_.ProvisioningState -eq "Provisioned"})
    $retry = 0
    
    #If CA servers count is 0, we might hit issue, try to re-load management shell 5 times to check if this can be fixed
    while($Servers.count -eq 0 -and $retry -le 5)
    {
        Write-EOPLog "No Server is in Provisioned State, Try to load management shell again" -Level Warn

        Load-ManagementShell
        sleep 60
        
        try
        {
            $DBservers = @(Get-CentralAdminMachine *DB*)
        }
        catch
        {
            Write-EOPLog -Message $_.Exception.Message -Level Error       
        }
       
        $Servers = @($DBservers|?{$_.ProvisioningState -eq "Provisioned"})  
        $retry ++   
    }
        
    $issueServers = @($DBservers|?{$_.ProvisioningState -ne "Provisioned"})

    if($issueServers -ne $Null)
    {
        Write-EOPLog "$($issueServers.count) Servers are not in Provisioned state, skip checking Service on $($issueServers.Name)."
    }
    if($servers.count -ne $DBservers.count -and $issueServers.count -eq $null)
    {
        Write-EOPLog "Managementshell load might have issue" -Level WARN
        $Issues = $servers|Out-String
        Write-EOPLog "$Issues" -Level Error  
        continue

    }

    return $Servers
    
}

Function Check-ServiceStatus
{
    param(
            [Parameter(Mandatory=$true)][string]$ServiceName,
            [Parameter(Mandatory=$true)]$Server
           )

    $Service = Get-Service -ComputerName $Server.Name -ServiceName $ServiceName 

    if($Service -eq $null)
    {
        Write-EOPLog $ServiceName could not find on  $($Server.Name) -level Error
        return $null
    
    }
    elseif($Service.count -gt 1)
    {
        Write-EOPLog duplicate Service $ServiceName found on $($Server.Name) -level Warn
        return $null
    }
    else
    {
        return $Service
    }
}


Function HardMode-StartService
{
    
    param(
            [Parameter(Mandatory=$true)][string]$ServiceName,
            [Parameter(Mandatory=$true)]$Server,
            [HashTable]$ServerRertyHash = @{}
        )

   
    Invoke-Command -ComputerName $Server.name -ArgumentList $ServiceName -scriptblock `
    {
        param($ServiceName)
        $Service = gwmi Win32_Service -Filter "Name = '$ServiceName'"
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
        #Stop service's process
        try
        {
            If($Service.Processid -ne 0)
            {
 
               $Log = "$FormattedDate Warning: Try to Stop the processID for service on $($server.name)"
               Stop-Process -Id $Service.Processid -Force
            
            }
        }
        catch
        {
             
             $Log += "$FormattedDate Error: $_.Exception.Message"
        }
        Sleep 10;
        # Get Process again to make sure the process is gone
        try
        {
            $Process = Get-Process -Id $Service.Processid 
        }
        catch
        {
             $Log += "$FormattedDate Error: $_.Exception.Message"
        }

        if($Process -eq $null)
        {
            Write-EOPLog ""
            $Log += "$FormattedDate Warning: restart ffo service for service on $($server.name) "
            Set-Service -Name $ServiceName  -StartupType disabled 
            Sleep 10;
            Set-Service -Name $ServiceName  -StartupType automatic -Status Running
            Start-Service -Name $ServiceName
        }

        
        # return $log
    } 
   if($HardRertyHash -eq $null)
    {
        $HardRertyHash.add($server.Name,$retry)
    }
        else
    {
        $totalretry += $retry
        $HardRertyHash[$server.name] = $totalretry
    }
    return $HardRertyHash
}

<#
#.Synopsis
#Restart service maxmium 5 times. record the retry info for the server
#    
#>
Function GentleStart-Service 
{
    param(
            [Parameter(Mandatory=$true)][string]$ServiceName,
            [Parameter(Mandatory=$true)]$Server,
            [HashTable]$ServerRertyHash = @{}
         )   
              
    $RunningServer = 0
    $maxretry = 5
    
    $Retry = 0
    $result = $false 

    while($retry -lt $maxretry -and $result -eq $false)
    {
        $Service = Check-ServiceStatus -Server $Server -ServiceName $ServiceName    
        if($Service.status -ne "Running")  
        {         
            Write-EOPLog "Try to start Service first genterally retry $retry time for $server" -level Warn
            try
            {
                $Service|Start-Service
            }
            catch
            {
                Write-EOPLog -Message $_.Exception.Message -Level Error
            }
            $retry++
            sleep 60
        }
        else
        {
            $result = $true
            #Write-EOPLog "Service is Running on $($Server.name), return"      
        }     
    }
    #"regesitry the service retry count is on $($server.name)"
    if($ServerRertyHash -eq $null)
    {
        $ServerRertyHash.add($server.Name,$retry)
    }
    else
    {
        $totalretry += $retry
        $ServerRertyHash[$server.name] = $totalretry
    }
    return $ServerRertyHash
}

if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
    $index = $Script:scriptname.LastIndexOf("_")
    $scriptname = $scriptname.substring(0,$index)
}

else
{
    $Script:basePath = "."
}

. D:\21V-GalOps\Tools\Library\LogHelper.ps1
Load-ManagementShell

$ServerRertyHash = @{}
$StartTime = $endtime = get-date
$LogDate = Get-Date -Format yyyy-MM-dd
Write-EOPLog "Program started at $Starttime"

#$record  is how many times the script and how many gentelstart count & hardstart count executed in the same day.

$GentelStartCount = $HardStartCount = $loopcount = 0

$ServerRertyHash =$HardStartCount =@{}
$ServiceName ="SQLSERVERAGENT"
$ScanningResults = @()
While($endtime -le ($StartTime.date.AddDays(1) - (New-TimeSpan -Minutes 3)))
#if($true)
{
  $servers =@()
  $Servers  = @(Get-ProvisionedDBServers)
  #Load ManageMent session in WS to double confirm if servers count return 0.
  if($Servers.count -eq 0)
  {  
        
        Write-EOPLog "No Server is in Provisioned State, Try to connect to WS servers to check CA Provision State" -Level WARN

        $s=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://bjbffo30ws001.protectioncn.gbl/powershell" -Authentication Kerberos -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck);
        $CAs = Invoke-Command -SessionName $s -ScriptBlock {Get-CentralAdminMachine *db*}
        $ProvioningDBs = $DBs|?{$_.ProvisioningState -eq "Provisioned"}
        Remove-PSSession $s       
        
        if($ProvioningDBs.count -eq 0)
        {
          
            $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
            $body =  "$FormattedDate No DB Server is in Provisioned State by loading management shell & checking from WS server. Please check"    
            Write-EOPLog "$body"    
            #Send-Mail -Subject "EOP-RestartService Tool running status" -Body $body 
            Sleep 60          
        }
        sleep 300
        continue
  }
 
  Write-EOPLog "$($Servers.count) Servers are in Provisioned state, checking Service now"

  #if Service is running, mark it to make sure all running Servers count equal Servers count if no issue happended
  $runningServer = 0

  Foreach($Server in $Servers)
  {
    
    $Service = Check-ServiceStatus -Server $Server -ServiceName $ServiceName
    if($Service.status -ne "Running")
    {
        #Gentle restart Service 5 times
      
        $GSInfo = GentleStart-Service -Server $Server -ServiceName $ServiceName -ServerRertyHash $ServerRertyHash

        $Service = Check-ServiceStatus -Server $Server -ServiceName $ServiceName -ServerRertyHash $HardStartCount
       

        if($Service.status -ne "Running")
        {
            $Service = HardMode-StartService -Server $Server -ServiceName $ServiceName
                
            $Service = Check-ServiceStatus -Server $Server -ServiceName $ServiceName
            #Send mail to team once ervice cannot be startted by gentle restart & hard model
            if($Service.status -ne "Running")
            {

                Write-EOPLog "Service cannot be startted by gentle restart & hard model on $($Server.Name)" -Level Error
                try
                {
                    Send-Mail -Subject "EOP-FFOService_cannot_be_started_on_$($Server.Name)-$LogDate"  -Body "Service cannot be startted by gentle restart & hard model on $($Server.Name)" `
                                   
                }
                catch
                {
                    Write-EOPLog -Message $_.Exception.Message -Level Error
                }
            }
        }
        
        
    }
    else
    {
        $runningServer += 1
    } 

    #Add this step in case we don't start the service manually
    if($ServerRertyHash[($Server.Name)] -eq $Null -or $ServerRertyHash[($Server.Name)] -eq 0)
    {
        $GentleStart = 0
    }
    else
    {
        $GentleStart = $ServerRertyHash[($Server.Name)] 
    }
 
    if($HardStartCount[($Server.Name)] -eq $Null -or $HardStartCount[($Server.Name)] -eq 0)
    {
        $HardStart = 0
    }
    else
    {
        $HardStart = $ServerRertyHash[($Server.Name)] 
    }
    $ScanningResult = New-Object   PSObject
    $ScanningResult | Add-Member -MemberType NoteProperty -Name Server -Value $Server.Name
    $ScanningResult | Add-Member -MemberType NoteProperty -Name  GentleStart -Value $GentleStart
    $ScanningResult | Add-Member -MemberType NoteProperty -Name  HardStart -Value $HardStart 
    $ScanningResults += $ScanningResult
   }
   $EndTime = get-date

    #Summary of scanning  one time
    if($RunningServer -eq $Servers.count)
    {
        Write-EOPLog "$ServiceName Service are running on All Servers"
    }
    else
    {
        
    }

    Write-EOPLog "Checking Service will be running in later 1 mins"
    
    sleep 60
    
    $loopcount++
}

<#
foreach($ServerName in $ServerRertyHash.keys)
{
    Write-EOPLog "$ServiceName service retried $($ServerRertyHash[$ServerName]) on $ServerName"
    $RetryInfo +=  "$ServiceName service retried $($ServerRertyHash[$ServerName]) on $ServerName"|Select-Object @{Name='ServerStatus';Expression={$_}} | select ServerStatus 
}
#>
Write-EOPLog "Ready to send mail to team for FFO service restart tool running status"

$title  =  "$LogDate $scriptname executed $loopcount times"
$keywords_color = @{}
$keywords_color.add("$ServiceName",'cyan')
$body = Format-HtmlTable -Contents $ScanningResults -Title $title -keywords_color $keywords_color
Send-Mail -Subject "EOP-FFOService_Monitoring_Tool_running_status-$LogDate" -Body $body 

