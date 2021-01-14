Function Check-ServiceStatus
{
    param(
            [String]$ServiceName,
            $Server
           )
    
    $Service = Get-Service -ComputerName $Server -ServiceName $ServiceName 

    if($Service -eq $null)
    {
        #Write-Log $ServiceName could not find on  $($Server.Name) -level Error
        return $null
    
    }

    else
    {
        return $Service
    }
}
################Operation for one server###############
Function Restart-RemoteService
{
    
    param(
            [String]$ServiceName,
            $Server,
            $Cred
        )
   $s = New-PSSession -ComputerName $Server -Credential $Cred
   $log= Invoke-Command -Session $s -ArgumentList $ServiceName,$Server -scriptblock `
    {
        param($ServiceName,$Server)
        $Service = gwmi Win32_Service -Filter "Name = '$ServiceName'"
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
        #Stop service's process
        try
        {
            If($Service.Processid -ne 0)
            {
 
               $Log = "$FormattedDate Warning: Try to Stop the processID for service on $server"
               Stop-Process -Id $Service.Processid -Force
            
            }
        }
        catch
        {
             
             $Log += "$FormattedDate Error: $_.Exception.Message"
        }
        Sleep 5;
        # Get Process again to make sure the process is gone

        
            $Log += "Warning: restart $ServiceName of $Server"
            Set-Service -Name $ServiceName  -StartupType disabled 
            Sleep 1
            $Log+=" server status: $((Get-Service $ServiceName).Status)   "
            Set-Service -Name $ServiceName  -StartupType automatic -Status Running
            Start-Service -Name $ServiceName
            $Log+="service status: $((Get-Service $ServiceName).Status)"
        
        Write-host "Log:$log"
    } 
  Remove-PSSession $s
   # Write-Log -Message $log
   
}

Function Reboot-RemoteServer
{
  
   param(
          $Server,
          $Cred
        )
#$s = New-PSSession -ComputerName $Server -Credential $Cred
Restart-Computer -ComputerName $Server -Credential $Cred -Force
<#
 Invoke-Command -Session $s -scriptblock `
 {
   Restart-Computer  -Force
 }
 Remove-PSSession $s
  <#
   try
    {  
     Set-MachinePowerState -Machine $Server -State Restart
    }
    catch
    {
      #Write-log $_.Exception.Message -level Error
      Write-Host “$_.Exception.Message”
     }
  #>
#>
}
Function Change-FileName
{
  param(
            [String]$FilePath,
            $Server,
            $NewName,
            $Cred
        )
  $s = New-PSSession -ComputerName $Server -Credential $Cred
  Invoke-Command -Session $s -ArgumentList $FilePath,$NewName -scriptblock `
    {
     param($FilePath,
           $NewName)
     
     Rename-Item -Path $FilePath -NewName $NewName 
    }
   Remove-PSSession $s
}

#########################Bulk operation for servers####################
Function Bulk-RestartService
{
   param(
            $Servers,
            $ServiceName,
            $Cred
        )

 Foreach($Server in $Servers)
  {

    Restart-RemoteService -Server $Server -ServiceName $ServiceName -Cred $Cred

        }
    
  }
 
 
 
 Function Bulk-RebootServer
{

   param(
          $Servers,
          $Cred
        )
   Foreach($Server in $Servers)
  {
     Reboot-RemoteServer $Server -Cred $Cred
     sleep 5
     $r=Test-Connection -ComputerName $Server -Quiet
     While((Test-Connection -ComputerName $Server -Quiet) -eq $false)
     {
        Write-host "$server has not started"
        sleep 5
     }
   }
}

Function Bulk-ChangeFileName
{
 param(
            [String]$FilePath,
            $Servers,
            $NewName,
            $Cred
        )
  Foreach($Server in $Servers)
  {
   Change-FileName -FilePath $FilePath -Server $Server -NewName $NewName -Cred $Cred

  }

}


$value=Read-host "Please input the task number you want and click Enter: `n 1--Restart server`n 2--Restart Service `n 3--Rename File`n" 
Switch($value)
{
1{$subvalue=Read-host "Please input the task number you want and click Enter:`n 1--Restart one server`n 2--Bulk restart multiple servers`n"
    Switch($subvalue)
    {
    1{Write-host 'Use this command: Reboot-RemoteServer -server <ComputerName> -Cred $Cred' -ForegroundColor Cyan}
    2{
       $Path=Read-Host "Please input the path of server list file. For example: D:\Eric_Script\servers.txt"
       $Servers=Get-Content -Path $Path
       Write-Host 'Please run command: Bulk-RebootServer -servers $Servers -Cred $Cred' -ForegroundColor Cyan
     }
    }  
  }
2{$subvalue=Read-host "Please input the task number you want and click Enter: `n 1--Restart service on one server `n 2--Bulk Restart Service on multiple servers`n"
    Switch($subvalue)
    {
    1{Write-host 'Use this command: Restart-RemoteService -server <ComputerName> -ServiceName <ServiceName> -Cred $Cred' -ForegroundColor Cyan}
    2{ $Path=Read-Host "Please input the path of server list file. For example: D:\Eric_Script\servers.txt"
       $Servers=Get-Content -Path $Path
       Write-Host 'Please run command: Bulk-RestartService -servers $Servers -servicename <ServiceName> -Cred $Cred' -ForegroundColor Cyan
     }
    }  
  }
3{$subvalue=Read-host "Please input the task number you want and click Enter: `n 1--Rename file name in one server `n 2--Bulk rename file name on multiple servers`n"
    Switch($subvalue)
    {
    1{Write-host 'Use this command:  Change-FileName -FilePath <D:\test.txt> -Server <ComputerName> -NewName <test1.txt> -Cred $Cred' -ForegroundColor Cyan}
    2{ $Path=Read-Host "Please input the path of server list file. For example: D:\Eric_Script\servers.txt"
       $Servers=Get-Content -Path $Path
       Write-Host 'Please run command: Bulk-ChangeFileName -servers $Servers -FilePath $FilePath -NewName $NewName -Cred $Cred' -ForegroundColor Cyan
     }
    }  
  }
}
$Cred= Get-Credential


    

