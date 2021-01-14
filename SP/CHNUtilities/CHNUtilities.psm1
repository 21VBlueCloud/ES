#####################################################################################
#
#File: CHNUtilities.psm1
#
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  9/26/2017   1.0.0      Wind        Combination of CHNMachineFunctions, CHNOEAccountFunctions,
##                                     CHNPerformance, CHNTenantFunctions, CHNTopologyFunctions,
##                                     MyUtilities and SPSHealthChecker.
##                                    
##  11/3/2017   1.0.1      Wind        Add parameter 'Credential' for Test-WindowsUpdates.
##
##  11/5/2017   1.0.2      Wind        Add parameter 'PassThru' on Get-HWHealthStatus for return objects.
##
##  11/9/2017   1.1.0      JK          new funciton: Get-PAVCResult
##
##  12/4/2017   1.1.1      Wind        Inprove performance for Test-MachineConnectivity
##                                       1. Deprecated parameter 'WithPSSessionTest'
##                                       2. Invoke Win32_ComputerSystem to test.
##                                       3. Accept pipeline input for ComputerName.
##
##  12/6/2017   1.1.2      Wind        Bug fix for Invoke-ZhouBapi:
##                                       1. To avoid currnet Centeral job empty.
##
##  12/17/2017  1.1.3      Wind        Change function Get-HWHealthStatus:
##                                       1. Add IP address.
##
##  12/17/2017  1.1.4      Wind        Change function Reset-OCEPassword
##                                       1. Change password at logon as FALSE due to MOR RDP policy changed.
##
##  1/24/2017   1.1.5      Wind        Add functions for pooling operators:
##                                       1. Set-OpsAccount
##
##  3/13/2018   1.1.6      Maggie      Add function: Get-HotfixReport
##  
##  3/13/2018   1.1.7      Wind        Update Functions:
##                                       1. Invoke-ZhouBapi: version 2.
## 
##  3/27/2018   1.1.8      Maggie      Remove function: Get-HotfixReport.
##                                       If you want to use this function, ask Maggie to obtain.
##
##  2/6/2019    1.1.9      Wind        Update function: Get-FolderSize
##                                       Add a field 'LastWriteTime' in result object.
##
##  7/30/2019   1.2.0      Wind        Add function Test-F5Login
##
##  9/9/2019    1.2.1      Wind        Add parameter 'Credential' for all HP functions.
##
##  3/5/2020    1.2.2      JK          Update function: Get-SyncEndPoint
##                                       Adapting multiple content farm.
##
##  3/18/2020   1.2.3      Yuejun      Update function Send-NotificationMail in CHNHelper.psm1
##                                       Change smtp from CME to EXO.
#####################################################################################

# <CHNMachineFunctions
Function Test-MachineConnectivity {

<#
.SYNOPSIS
Test machines is live or not by WMI query.

.DESCRIPTION
This function is used for get machines live status. Generally, it will return TRUE or FALSE to present all 
computers are reached or not. If parameter 'ReturnBad' or 'ReturnGood' is present, it will return computer
name array instead of boolean.
This function invoke WMI query and PowerShell job to imporve performace. All 

.PARAMETER ComputerName
Alias is 'CN'. It accept string array and pipeline.

.PARAMETER Credential
Provide query credential.

.PARAMETER ReturnBad
It will return un-reached machines list if this parameter is present.

.PARAMETER ReturnGood
It will return reached machines list if this parameter is present.

.INPUTS
The string array of target computer name.

.OUTPUTS
It will return TRUE or FALSE by default. All Computers reached, it will return TRUE. Otherwise, return FLASE.
Return string array of computer name if parameter 'RetrunBad' or 'ReturnGood' is present.

.EXAMPLE
PS C:\> Test-MachineConnectivity "bjbchnvut144","bjbchnvut155"
Return TURE if these 2 servers are reached.

.EXAMPLE
PS C:\> Test-MachineConnectivity $gfe -ReturnBad
Return bad machines' name from $gfe.

.EXAMPLE
PS C:\> $Server | Test-MachineConnectivity -ReturnGood
Return reachable machines' name from $Server.

#>

    Param (
        [Parameter(Mandatory=$true,Position=0,
            ValueFromPipeline=$true,ParameterSetName="bool")]
        [Parameter(Mandatory=$true,Position=0,
            ValueFromPipeline=$true,ParameterSetName="Bad")]
        [Parameter(Mandatory=$true,Position=0,
            ValueFromPipeline=$true,ParameterSetName="Good")]
        [Alias("CN")]
        [String[]] $ComputerName,
        [System.Management.Automation.PSCredential] $Credential,
        [Parameter(Mandatory=$true,ParameterSetName="Bad")]
        [Switch] $ReturnBad,
        [Parameter(Mandatory=$true,ParameterSetName="Good")]
        [Switch] $ReturnGood
    )

    Begin{

        $Server = @()
        $LocalHost = $env:COMPUTERNAME
        $return = $true
        $Expression = "Get-WmiObject -Class Win32_ComputerSystem -ComputerName `$Server"
        if ($Credential) {
            $Expression += " -Credential `$Credential"
        }

    }

    Process{

        $Server += $ComputerName

    }

    End {

        # Exclude localhost from list
        if ($Server -match $LocalHost) {

            $Server = $Server -notmatch $LocalHost
            if (!$Server -or $Server -is [Boolean]) {
                # It means the ComputerName is localhost only. Stop to test.
                Write-Host "The Computer '$LocalHost' is localhost! Stop to test." -ForegroundColor Red
                break
            }

            Write-Host "Localhost '$LocalHost' has been excluded from 'ComputerName'!" -ForegroundColor Yellow
            Write-Host "Continue to test..." -ForegroundColor Yellow
        }

        # Start to test
        # If the servers count more than 50, we should query as job for speed up.
        if ($Server.Count -le 50) {
            $Result = @(Invoke-Expression $Expression)
        }
        else {
            $Expression += " -AsJob"
            $Result = @(Invoke-Expression $Expression | Watch-Jobs -Activity "Test connectivity..." -Status Running)
        }

        
        if ($ReturnGood.IsPresent) {
            return $Result.PSComputerName
        }

        if ($ReturnBad.IsPresent) {
            $PassedServer = $Result.PSComputerName
            # To find out the failed servers
            $BadServer = Compare-Object $Server $PassedServer | ? SideIndicator -EQ "<=" | select -ExpandProperty InputObject
            return $BadServer
        }

        if ($Result.Count -ne $Server.Count) {
            
            $return = $false

        }

        return $return
    }

}

Function Get-HPIML {
<#
.SYNOPSIS
Get HP IML logs.

.DESCRIPTION
This function can get IML logs from WMI or iLO. Provide iLO credential, it will retrive logs from iLO.

.PARAMETER ComputerName
The remote PMs which IML logs you want to get. Alias is "CN".

.PARAMETER Last
The nearest entries count.

.PARAMETER CreationTimeAfter
Filter logs which creation time great than this time.

.PARAMETER Severity
Log severity level. Display all of levels log entries if omitted.
    0 (Unknown)
    2 (Informatinal)
    3 (Repaired)
    4 (Caution)
    7 (Critial)

.PARAMETER iLoCredential
iLO credential. It will retrive logs from iLO if provided this parameter.

.EXAMPLE
Get-HPIML -CN bjbchnvut144 -Severity Critical -iLoCredential $cnops
Get IML logs which severity equal critial from bjbchnvut144 via iLO.
#>

    Param (
        [Parameter(Mandatory=$true)]
        [Alias("CN")]
        [String[]] $ComputerName,
        [Int] $Last,
        [DateTime] $CreationTimeAfter,
        [ValidateSet("Unknown","Informational","Repaired","Caution","Critical")]
        [string[]] $Severity,
        [PSCredential] $iLoCredential
    )

    Begin {

        # filter PM out
        $PMs = @($ComputerName | %{Get-CenteralPMachine -Name $_ -ErrorAction Ignore})
        if ($PMs.Count -lt $ComputerName.Count) {
            $InvalidPMs = @($ComputerName | ?{$_ -notin $PMs.Name})
            Write-Host "Invalid PM: $(Convert-ArrayAsString $InvalidPMs)" -ForegroundColor Yellow
            if (!$PMs) {
                Write-Host "No valid PM name!"
                break
            }
        }

        $SeveritySample = @{
            0 = "Unknown"
            2 = "Informational"
            3 = "Repaired"
            4 = "Caution"
            7 = "Critical"
        }

        if ($iLoCredential) {
            try { Import-Module HPiLOCmdlets -ErrorAction Stop }
            catch { Write-Host "Cannot find module 'HPiLOCmdlets'!" -ForegroundColor Red; break }
            
            $NoILOPMs = @($PMs | ? iloIpaddress -EQ "")
            if ($NoILOPMs) {
                Write-Host "PM: $(Convert-ArrayAsString $NoILOPMs.Name) iLO IP address absence." -ForegroundColor Yellow
                if ($NoILOPMs.Count -eq $PMs.Count) {
                    Write-Host "No valid iLO IP addresses!" -ForegroundColor Red
                    break
                }
            }
            
            # Nested function for ILO
            function ConvertIMLEntries {
                param ([Object[]] $InputObject,[Object[]] $IPTable)

                $result = @()
                $LocalTimeZone = [System.TimeZoneInfo]::Local
                foreach ($object in $InputObject) {
                    $iloIP = $object.IP
                    $ComputerName = $IPTable | ? iLoIPaddress -EQ $iloIP | select -ExpandProperty Name
                    $event = $object.event
                    foreach ($e in $event) {
                        if ($e.INITIAL_UPDATE -ne '[NOT SET] ') {
                            $CreationTime = [DateTime]::Parse($e.INITIAL_UPDATE)
                            $CreationTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($CreationTime,$LocalTimeZone)
                            $UpdateTime = [DateTime]::Parse($e.LAST_UPDATE)
                            $UpdateTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UpdateTime,$LocalTimeZone)
                        }
                        else {
                            $CreationTime = $null
                            $UpdateTime = $null
                        }

                        $result += New-Object -TypeName PSObject -Property @{
                            ComputerName = $ComputerName
                            Description = $e.Description
                            CreationTime = $CreationTime
                            UpdateTime = $UpdateTime
                            Severity=$e.Severity
                        }
                    }
                }

                return $result
            }

        }
        else {
            # Nested function for WMI
            function ConvertWmiLogEntries {
                param ([object[]] $WmiEntry)
                
                $result = @()
                foreach ($entry in $WmiEntry) {
                    if ($entry.CreationTimeStamp -ne "00000000000000.000000+000") {
                        $CreationTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($entry.CreationTimeStamp)
                        $UpdateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($entry.UpdateTimeStamp)
                    }
                    else {
                        $CreationTime = $null
                        $UpdateTime = $null
                    }
                    $obj = New-Object -TypeName PSObject -Property @{
                        ComputerName = $entry.PSComputerName
                        Description = $entry.Description
                        CreationTime = $CreationTime
                        UpdateTime = $UpdateTime
                        Severity=$SeveritySample[($entry.Severity -as [Int])]
                    }
                    $result += $obj
                }
                return $result
            }

        }
    }

    Process {
        If (!$iLoCredential) {
            # fetch logs from WMI
            if ($ComputerName.Count -lt 10) {
                $WmiLogEntries = Get-WmiObject -ComputerName $ComputerName -Namespace root\hpq -Class HP_CommonLogEntry
            }
            else {
                Get-WmiObject -ComputerName $ComputerName -Namespace root\hpq -Class HP_CommonLogEntry -AsJob | Out-Null
                $WmiLogEntries = Watch-Jobs -Activity "Fetching IML entries" -Status Querying
            }

            $LogEntries = ConvertWmiLogEntries -WmiEntry $WmiLogEntries

            <#
            # Convert DMTF datetime to local datetime.
            foreach ($log in $WmiLogEntries) {
                if ($log.CreationTimeStamp -ne "00000000000000.000000+000") {
                    Add-Member -InputObject $log -NotePropertyName CreationTime -NotePropertyValue ([System.Management.ManagementDateTimeConverter]::ToDateTime($log.CreationTimeStamp))
                    Add-Member -InputObject $log -NotePropertyName UpdateTime -NotePropertyValue ([System.Management.ManagementDateTimeConverter]::ToDateTime($log.UpdateTimeStamp))
                }
                else {
                    Add-Member -InputObject $log -NotePropertyName CreationTime -NotePropertyValue $null
                    Add-Member -InputObject $log -NotePropertyName UpdateTime -NotePropertyValue $null
                }
            }
            #>

        }
        else {
            # fetch logs from iLO
            $iLOLog = Get-HPiLOIML -Server $PMs.iLOIpAddress -Credential $iLoCredential -DisableCertificateAuthentication -WarningAction SilentlyContinue
            if ($iLOLog) {
                # I input iLO IP and ComputerName for reference to speed up reverse lookup Computer Name.
                $LogEntries = ConvertIMLEntries -InputObject $iLOLog -IPTable ($PMs | select -Property Name,iLoIPaddress)
            }
        }


        # filter out severity logs if needed.
        if ($Severity) {
            $TempLog = @()
            foreach ($sev in $Severity) {
                $TempLog += $LogEntries | ? Severity -EQ $sev
            }
            $LogEntries = $TempLog
        }

        $LogEntries = $LogEntries | Sort-Object ComputerName,CreationTime
        
        If ($CreationTimeAfter) {
            $LogEntries = $LogEntries | ? CreationTime -GT $CreationTimeAfter
        }

        If ($Last) {
            $LogEntries = $LogEntries | Select-Object -Last $Last
        }

        Return $LogEntries
    }
}

Function Get-HyperVMachine {
<#
.SYNOPSIS
Get VMs from PMachine.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine name. Alias are "PM","PMachine","CN".

.EXAMPLE
PS C:\> Get-HyperVMachine BJBCHNVUT082
Get all of VMs from PM BJBCHNVUT082.
#>

    Param (
        [Parameter(Mandatory=$true,Position=0)] [Alias("PM","PMachine","CN")]
        [String] $ComputerName
    )
    try {
        $HyperVFeature = Get-WmiObject -Class Win32_ServerFeature -ComputerName $ComputerName | ? Id -EQ 20
    }
    catch {
        Write-Host "Can not get Features on $ComputerName!" -ForegroundColor Red
        Break
    }
    If ($HyperVFeature) {
        $os = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName | Select-Object -ExpandProperty Caption
        If ($os -match 2012) {
            $namespace = "Root\Virtualization\V2"
        }
        Else {
            $namespace = "Root\Virtualization"
        }
        $VMs = Get-WmiObject -Namespace $namespace -Class Msvm_ComputerSystem -ComputerName $ComputerName
        # Filter Virtual Machine out
        $VMs = $VMs | ? Caption -EQ "Virtual Machine"
    }
    Else {
        Write-Host "Can not find Hyper-v feature on $ComputerName!" -ForegroundColor Yellow
        Break
    }

    Return $VMs

}

Function Get-MemoryUsage {
<#
.SYNOPSIS
Get memory usage.

.DESCRIPTION

.PARAMETER ComputerName
Machines which you want to check physical memory.

.EXAMPLE
PS C:\> Get-MemoryUsage BJBCHNVUT144
Get memory usage for BJBCHNVUT144.
#>

    Param (
        [Parameter(Position=0)] [Alias("CN")]
        [String[]] $ComputerName = $env:COMPUTERNAME
    )

    Process {
        If ($ComputerName.Count -lt 10) {
            $WmiPhysicalMemory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $ComputerName
            $WmiPerfMemory = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Memory -ComputerName $ComputerName
        }
        Else {
            # Clean jobs which are not kicked off by this function.
            Get-Job | Remove-Job -Force

            $PMJob = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $ComputerName -AsJob
            $PeJob = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Memory -ComputerName $ComputerName -AsJob

            $WmiPhysicalMemory = $PMJob | Watch-Jobs -Activity "Fetching physical memory size..."
            $WmiPerfMemory = $PeJob | Watch-Jobs -Activity "Fetching free memory..."
        }
        $MemoryUsage = @()
        Foreach ($c in $ComputerName) {
            $Pmemory = $WmiPhysicalMemory | ? PSComputerName -EQ $c
            $TotalMemory = 0
            foreach ($m in $Pmemory) { $TotalMemory += $m.Capacity }
            $FreeMemory = $WmiPerfMemory | ? PSComputerName -EQ $c | Select-Object -ExpandProperty AvailableBytes
            $Usage = ($TotalMemory - $FreeMemory)/$TotalMemory
            $properties = @{
                ComputerName = $c
                Size = $TotalMemory
                FreeSpace = $FreeMemory
                Usage = $Usage
            }

            $MemoryUsage += New-Object -TypeName PSObject -Property $properties
        }
        Return $MemoryUsage
      }
}

Function Get-DiskUsage {
<#
.SYNOPSIS
Get disk usage.

.DESCRIPTION
Just get logical hard drives. If computers more than 10, it will query in job mode.
PM and VM are acceptable both.

.PARAMETER ComputerName
Which computers you want to get.

.PARAMETER DriveLetter
Specify drive letter which you want to get.

.EXAMPLE
PS C:\> Get-DiskUsage GFE20662-001
Get all drive usage for GFE20662-001.

.EXAMPLE
PS C:\> Get-DiskUsage GFE20662-001 C:
Get drive C usage for GFE20662-001.
#>
    Param (
        [Parameter(Position=0)] [Alias("CN")]
        [String[]] $ComputerName,
        [Parameter(Position=1)]
        [ValidateScript({$_ -match "^[a-z]:?$"})]
        [String] $DriveLetter
    )

    $Disks = @()

    if (!$ComputerName) { $ComputerName = @("localhost") }

    # Inspect drive letter
    if ($DriveLetter) { 
        if ($DriveLetter.Length -eq 1) { $DriveLetter = $DriveLetter + ":" }
        $filter = "drivetype=3 and DeviceID='$DriveLetter'"
    }
    else { $filter = "drivetype=3" }

    # Query in job mode if computers more than 10.
    If ($ComputerName.Count -lt 10) {
        $WmiDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter $filter -ComputerName $ComputerName
    }
    Else {
        # Clean jobs which are not kicked off by this function.
        Get-Job | Remove-Job -Force

        foreach ($cn in $ComputerName) {
            Get-WmiObject -Class Win32_LogicalDisk -Filter $filter -ComputerName $cn -AsJob | Out-Null
        }
        $WmiDisks = Watch-Jobs -Activity "Fetching disk usage..."
    }
    
    foreach ($Wd in $WmiDisks) {
        $Usage = ($Wd.Size - $Wd.FreeSpace) / $Wd.Size
        $preporties = @{
            ComputerName = $Wd.PSComputerName
            DriveLetter = $Wd.DeviceID
            FreeSpace = $Wd.FreeSpace
            Size = $Wd.Size
            Usage = $Usage
        }
        $Disks += New-Object -TypeName PSObject -Property $preporties
    }

    Return $Disks
}

Function Get-HPPhysicalDisk {
<#
.SYNOPSIS
Get HP physical disks from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get physical disks.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get physical disks from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>
    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad
            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        # Start job mode if computers more than 10.
        If ($ComputerName.Count -ge 10) {
            #Clean old jobs which are not kicked off by this function.
            Get-Job | Remove-Job -Force
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Class HPSA_DiskDrive -Namespace root\hpq -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Class HPSA_DiskDrive -Namespace root\hpq -ComputerName $cn -AsJob | Out-Null
                }
            }

            $PDisks = Watch-Jobs -Activity "Fetching HP physical disks ..."
        }
        Else {
            if ($Credential) {
                $PDisks = Get-WmiObject -Credential $Credential -Class HPSA_DiskDrive -Namespace root\hpq -ComputerName $ComputerName
            }
            else {
                $PDisks = Get-WmiObject -Class HPSA_DiskDrive -Namespace root\hpq -ComputerName $ComputerName
            }
        }

        Return $PDisks
    }
}

Function Get-HPCPU {
<#
.SYNOPSIS
Get HP CPU from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get CPU.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get CPU from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad
            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 3) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Processor -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_Processor -ComputerName $cn -AsJob | Out-Null
                }
            }
            $CPUs = Watch-Jobs -Activity "Fetching CPU ..."
        }
        Else {

            if ($Credential) {
                $CPUs = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Processor -ComputerName $ComputerName
            }
            else {
                $CPUs = Get-WmiObject -Namespace root\hpq -Class HP_Processor -ComputerName $ComputerName
            }
        }

        Return $CPUs
    }
}

Function Get-HPMemory {
<#
.SYNOPSIS
Get HP Memory from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get Memory.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get Memory from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Memory -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_Memory -ComputerName $cn -AsJob | Out-Null
                }
            }

            $Memory = Watch-Jobs -Activity "Fetching memory ..."
        }
        Else {
            
            if ($Credential) {
                $Memory = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Memory -ComputerName $ComputerName
            }
            else {
                $Memory = Get-WmiObject -Namespace root\hpq -Class HP_Memory -ComputerName $ComputerName
            }
        }

        Return $Memory
    }
}

Function Get-HPMemoryModule {
<#
.SYNOPSIS
Get HP Memory Module from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get Memory Module.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get Memory Module from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_MemoryModule -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_MemoryModule -ComputerName $cn -AsJob | Out-Null
                }
            }

            $MemoryModule = Watch-Jobs -Activity "Fetching memory module ..."
        }
        Else {

            if ($Credential) {
                $MemoryModule = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_MemoryModule -ComputerName $ComputerName
            }
            else {
                $MemoryModule = Get-WmiObject -Namespace root\hpq -Class HP_MemoryModule -ComputerName $ComputerName
            }
            
        }

        Return $MemoryModule
    }
}

Function Get-HPFan {
<#
.SYNOPSIS
Get HP Fan from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get Fan.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get Fan from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Fan -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_Fan -ComputerName $cn -AsJob | Out-Null
                }
            }
            $Fans = Watch-Jobs -Activity "Fetching fan ..."
        }
        Else {
            if ($Credential) {
                $Fans = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_Fan -ComputerName $ComputerName
            }
            else {
                $Fans = Get-WmiObject -Namespace root\hpq -Class HP_Fan -ComputerName $ComputerName
            }
        }

        Return $Fans
    }
}

Function Get-HPPowerSupply {
<#
.SYNOPSIS
Get HP power supply from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get power supply.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get power supply from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                    foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_PowerSupply -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_PowerSupply -ComputerName $cn -AsJob | Out-Null
                }
            }

            $PowerSupply = Watch-Jobs -Activity "Fetching power supply ..."
        }
        Else {
            if ($Credential) {
                $PowerSupply = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_PowerSupply -ComputerName $ComputerName
            }
            else {
                $PowerSupply = Get-WmiObject -Namespace root\hpq -Class HP_PowerSupply -ComputerName $ComputerName
            }
        }

        Return $PowerSupply
    }
}

Function Get-HPPowerRedundancySet {
<#
.SYNOPSIS
Get HP power redundancy set from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get power redundancy set.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get power redundancy set from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_PowerRedundancySet -ComputerName $cn -AsJob | Out-Null
            }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_PowerRedundancySet -ComputerName $cn -AsJob | Out-Null
                }
            }
            
            $PowerRedundancySet = Watch-Jobs -Activity "Fetching power redundancy set ..."
        }
        Else {
            if ($Credential) {
                $PowerRedundancySet = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_PowerRedundancySet -ComputerName $ComputerName
            }
            else {
                $PowerRedundancySet = Get-WmiObject -Namespace root\hpq -Class HP_PowerRedundancySet -ComputerName $ComputerName
            }
        }

        Return $PowerRedundancySet
    }
}

Function Get-HPSmartArray {
<#
.SYNOPSIS
Get HP smart array from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get smart array.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get smart array from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HPSA_ArraySystem -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HPSA_ArraySystem -ComputerName $cn -AsJob | Out-Null
                }
            }
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HPSA_ArraySystem -ComputerName $cn -AsJob | Out-Null
                }
            $ArraySystem = Watch-Jobs -Activity "Fetching HP smart array ..."
        }
        Else {
            if ($Credential) {
                $ArraySystem = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HPSA_ArraySystem -ComputerName $ComputerName
            }
            else {
                $ArraySystem = Get-WmiObject -Namespace root\hpq -Class HPSA_ArraySystem -ComputerName $ComputerName
            }
        }

        Return $ArraySystem
    }
}

Function Get-HPArrayController {
<#
.SYNOPSIS
Get HP array controller from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get array controller.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get array controller from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HPSA_ArrayController -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {

                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HPSA_ArrayController -ComputerName $cn -AsJob | Out-Null
                }
            }
            $ArrayController = Watch-Jobs -Activity "Fetching HP array controller ..."
        }
        Else {
            if ($Credential) {
                $ArrayController = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HPSA_ArrayController -ComputerName $ComputerName
            }
            else {
                $ArrayController = Get-WmiObject -Namespace root\hpq -Class HPSA_ArrayController -ComputerName $ComputerName
            }
        }

        Return $ArrayController
    }
}

Function Get-HPEthernetPort {
<#
.SYNOPSIS
Get HP ethernet adapter from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get ethernet adapter.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get ethernet adapter from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_EthernetPort -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_EthernetPort -ComputerName $cn -AsJob | Out-Null
                }
            }
            $EthernetPort = Watch-Jobs -Activity "Fetching HP Ethernet Port ..."
        }
        Else {
            if ($Credential) {
                $EthernetPort = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_EthernetPort -ComputerName $ComputerName
            }
            else {
                $EthernetPort = Get-WmiObject -Namespace root\hpq -Class HP_EthernetPort -ComputerName $ComputerName
            }
        }

        Return $EthernetPort
    }
}

Function Get-HPComputerSystem {
<#
.SYNOPSIS
Get HP computer system info from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get computer system info.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get computer system info from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 3) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                    foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_WinComputerSystem -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_WinComputerSystem -ComputerName $cn -AsJob | Out-Null
                }
            }
            $WinComputerSystem = Watch-Jobs -Activity "Fetching HP Computer system info ..."
        }
        Else {
            if ($Credential) {
                $WinComputerSystem = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_WinComputerSystem -ComputerName $ComputerName
            }
            else {
                $WinComputerSystem = Get-WmiObject -Namespace root\hpq -Class HP_WinComputerSystem -ComputerName $ComputerName
            }
        }

        Return $WinComputerSystem
    }
}

Function Get-HPComputerSystemChassis {
<#
.SYNOPSIS
Get HP hardware info from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get hardware info.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get hardware info from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                    foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_ComputerSystemChassis -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_ComputerSystemChassis -ComputerName $cn -AsJob | Out-Null
                }
            }
            $HWinfo = Watch-Jobs -Activity "Fetching HP Computer system info ..."
        }
        Else {
            if ($Credential) {
                $HWinfo = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_ComputerSystemChassis -ComputerName $ComputerName
            }
            else {
                $HWinfo = Get-WmiObject -Namespace root\hpq -Class HP_ComputerSystemChassis -ComputerName $ComputerName
            }
        }

        Return $HWinfo
    }
}

Function Get-HPManagementProcessor {
<#
.SYNOPSIS
Get HP iLO info from WMI.

.DESCRIPTION

.PARAMETER ComputerName
Physical machine names. Alias are "PM","PMachine","CN".

.PARAMETER WithPingTest
Execute ping test before get iLO info.

.EXAMPLE
PS C:\> Get-HPPhysicalDisk "bjbchnvut144","shachnvut144","GFE20662-001" -WithPingTest
Get iLO info from "bjbchnvut144" and "shachnvut144" after ping test them.
"GFE20662-001" will be filtered out because it is not PM.
#>

    Param (
        [Parameter(Mandatory=$true)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [PsCredential] $Credential,
        [Switch] $WithPingTest
    )

    Begin {
        
        # filter VM out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        If ($WithPingTest.IsPresent) {

            # Validate PMs connectivity.
            $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad

            If ($FailedComputers) {
                Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                Write-Host $FailedComputers -ForegroundColor Red
                If ($FailedComputers.Count -eq $ComputerName.Count) {
                    Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                    Break
                }
            }
            $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
        }
    }

    Process {
        If ($ComputerName.Count -ge 10) {
            # Remove jobs which are not kicked of from here.
            Get-Job | Remove-Job -Force
            
            if ($Credential) {
                    foreach ($cn in $ComputerName) {
                    Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_ManagementProcessor -ComputerName $cn -AsJob | Out-Null
                }
            }
            else {
                foreach ($cn in $ComputerName) {
                    Get-WmiObject -Namespace root\hpq -Class HP_ManagementProcessor -ComputerName $cn -AsJob | Out-Null
                }
            }
            $iLO = Watch-Jobs -Activity "Fetching HP Computer system info ..."
        }
        Else {
            if ($Credential) {
                $iLO = Get-WmiObject -Credential $Credential -Namespace root\hpq -Class HP_ManagementProcessor -ComputerName $ComputerName
            }
            else {
                $iLO = Get-WmiObject -Namespace root\hpq -Class HP_ManagementProcessor -ComputerName $ComputerName
            }
        }

        Return $iLO
    }
}

Function Get-HWHealthStatus {
<#
.SYNOPSIS
Qeury hardware health status from PM via HP WMI.

.DESCRIPTION
It just display bad status if parameter 'IncludeGoodStatus' not present.

.PARAMETER ComputerName
Physical machines which you want to check health status.
Do not care server type, it will filter PMs automatically.
Alias are "PM","PMachine" and "CN".

.PARAMETER IncludeGoodStatus
Display Good status for every check point.

.PARAMETER WithPingTest
Will ping test before query.

.PARAMETER PassThru
For return objects. If not with this parameter, it just print result in console.

.EXAMPLE
PS C:\> Get-HWHealthStatus BJBCHNVUT144 -IncludeGoodStatus
Check BJBCHNVUT144 hardware with good status.

.EXAMPLE
PS C:\> Get-HWHealthStatus -ComputerName $cn
Check PMs which names are stored in $cn. Just display bad status.
#>
    Param (
        [Parameter(Mandatory=$true,Position=0)] [Alias("PM","PMachine","CN")]
        [String[]] $ComputerName,
        [Switch] $IncludeGoodStatus,
        [Switch] $WithPingTest,
        [Switch] $PassThru,
        [PsCredential] $Credential
    )

    Begin {

        ## Add credential
        if ($Credential) {
            
        }
        
        # filter VMs out
        $NonPMs = $ComputerName | ?{$_ -notmatch "BJB"} | ?{$_ -notmatch "SHA"}
        If ($NonPMs) {
            Write-Host "Below machines are not PM:" -ForegroundColor Yellow
            Write-Host $NonPMs -ForegroundColor Yellow
            Write-Host
            $ComputerName = Compare-Object $ComputerName $NonPMs | Select-Object -ExpandProperty InputObject
        }

        
            # Validate PMs connectivity.
            if($WithPingTest.IsPresent) {
                $FailedComputers = Test-MachineConnectivity -ComputerName $ComputerName -ReturnBad
                If ($FailedComputers) {
                    Write-Host "Below computers cannot pass ping test:" -ForegroundColor Yellow
                    Write-Host $FailedComputers -ForegroundColor Red
                    If ($FailedComputers.Count -eq $ComputerName.Count) {
                        Write-Host "`nAll of Computers have not passed ping test!" -ForegroundColor Red
                        Break
                    }
                }
                $ComputerName = $PingResult | ? Ping -EQ $true | Select-Object -ExpandProperty ComputerName
            }

    }

    Process {
        
        Write-Progress -Activity "Collecting PM info ..." -Status "Running"

        if (!$Credential) {
            $WmiHPWCS = Get-HPComputerSystem -ComputerName $ComputerName
        }
        else {
            $WmiHPWCS = Get-HPComputerSystem -ComputerName $ComputerName -Credential $Credential
        }

        # check health state
        #    0  (Unknown)
        #    5  (OK)
        #    10 (Degraded)
        #    20 (Major Failure)
        If (!$IncludeGoodStatus.IsPresent) { $WmiCandiddateObjs = $WmiHPWCS | ? HealthState -NE 5 }
        Else { $WmiCandiddateObjs = $WmiHPWCS }

        If ($WmiCandiddateObjs) {
            
            $MachineName = $WmiCandiddateObjs | Select-Object -ExpandProperty PSComputerName

            # Check all of HW health.
            if (!$Credential) {
                <#
                CPU HealthState 
                     0 (Unknown), when OperationalStatus[0]=0 (Unknown)
                     5 (OK), when OperationalStatus[0]=2 (OK)
                     15 (Minor Failure), when OperationalStatus[0]=10 (Stopped)
                     20 (Major Failure), when OperationalStatus[0]=5 (Predictive Failure)
                     25 (Critical Failure), when OperationalStatus[0]=6 (Error)
                #>
                $CPUs = @(Get-HPCPU -ComputerName $MachineName)

                <#
                Memory HealthState 
                     5 (OK), when OperationalStatus[0]=2 (OK)
                     0 (Unknown), when OperationalStatus[0]=0 (Unknown)
                     10 (Degraded/Warning), when OperationalStatus[0]=3 (Degraded)
                #>
                $Memory = @(Get-HPMemory -ComputerName $MachineName)

                <#
                Memory module HealthState 
                    Enumerator indicating the memory module health state:
                     0 (Unknown), when OperationalStatus[0] = 0
                    (Unknown)
                     5 (OK), when OperationalStatus[0] = 2 (OK)
                     10 (Degraded/Warning), when
                    OperationalStatus[0] = 3 (Degraded)
                #>
                $MemoryModule = @(Get-HPMemoryModule -ComputerName $MachineName)

                <#
                Fan HealthState 
                     5 (OK)—If fan is operating properly
                     20 (Major Failure)—If fan has failed
                #>
                $Fans = @(Get-HPFan -ComputerName $MachineName)

                <#
                PowerSupply HealthState 
                     5 (OK)—If the power supply is operating properly
                     20 (Major Failure)—if the power supply has failed
                #>
                $Power = @(Get-HPPowerSupply -ComputerName $MachineName)

                <#
                Redundancy Status 
                     0 (Unknown)—If the redundancy status is unknown
                     2 (Fully Redundant)—If the all power supplies in the
                    set are operating properly and enough to achieve
                    redundancy
                     3 (Degraded Redundancy)—If there are at least
                    enough power supplies for the redundancy set to
                    provide power, but power supplies have failed
                     4 (Redundancy Lost)—If there are not enough
                    power supplies required to achieve redundancy,
                    but there are enough for the redundancy set to
                    provide power
                     5 (Overall Failure)—If there are not enough power
                    supplies operating properly
                #>
                $PwrRedundantSet = @(Get-HPPowerRedundancySet -ComputerName $MachineName)

                <#
                Smart array OperationalStatus[0] 
                    Overall status of the Array System and attached
                    devices. This is calculated per the algorithm desciribed
                    in the HP Smart Array Profile.
                     0 (Unknown)
                     2 (OK)
                     3 (Degraded)
                     6 (Error)
                #>
                $SmartArray = @(Get-HPSmartArray -ComputerName $MachineName)

                <#
                OperationalStatus[0] 
                    Status for Array Controller
                     0 (Unknown)
                     2 (OK)
                     6 (Error)
                #>
                $ArrayController = @(Get-HPArrayController -ComputerName $MachineName)

                <#
                Operational Status for the disk drive
                     0 (Unknown)
                     2 (OK)
                     5 (Predictive Failure)
                     6 (Error)
                #>
                $PDisks = @(Get-HPPhysicalDisk -ComputerName $MachineName)

                <#
                Ethernet HealthState 
                    5 (OK) if port has link, 
                    20 (Major Failure) otherwise
                #>
                $network = @(Get-HPEthernetPort -ComputerName $MachineName)

                $iLOInfo = @(Get-HPManagementProcessor -ComputerName $MachineName)
            
                $HWinfo = Get-HPComputerSystemChassis -ComputerName $MachineName

            }
            else {
                $CPUs = @(Get-HPCPU -ComputerName $MachineName -Credential $Credential)

                $Memory = @(Get-HPMemory -ComputerName $MachineName -Credential $Credential)

                $MemoryModule = @(Get-HPMemoryModule -ComputerName $MachineName -Credential $Credential)

                $Fans = @(Get-HPFan -ComputerName $MachineName -Credential $Credential)

                $Power = @(Get-HPPowerSupply -ComputerName $MachineName -Credential $Credential)

                $PwrRedundantSet = @(Get-HPPowerRedundancySet -ComputerName $MachineName -Credential $Credential)

                $SmartArray = @(Get-HPSmartArray -ComputerName $MachineName -Credential $Credential)

                $ArrayController = @(Get-HPArrayController -ComputerName $MachineName -Credential $Credential)

                $PDisks = @(Get-HPPhysicalDisk -ComputerName $MachineName -Credential $Credential)

                $network = @(Get-HPEthernetPort -ComputerName $MachineName -Credential $Credential)

                $iLOInfo = @(Get-HPManagementProcessor -ComputerName $MachineName -Credential $Credential)
            
                $HWinfo = Get-HPComputerSystemChassis -ComputerName $MachineName -Credential $Credential
            }

            If (!$IncludeGoodStatus.IsPresent) {
                $CPUs = $CPUs | ? HealthState -NE 5
                $Memory = $Memory | ? HealthState -NE 5
                $MemoryModule = $MemoryModule | ? HealthState -NE 5
                $Fans = $Fans | ? HealthState -NE 5
                $Power = $Power | ? HealthState -NE 5
                $PwrRedundantSet = $PwrRedundantSet | ? RedundancyStatus -NE 2
                $SmartArray = $SmartArray | ?{$_.OperationalStatus[0] -NE 2} 
                $ArrayController = $ArrayController | ?{$_.OperationalStatus[0] -NE 2}
                $PDisks = $PDisks | ?{$_.OperationalStatus[0] -NE 2}
                $network = $network | ? HealthState -NE 5
                $iLOInfo = $iLOInfo | ? HealthState -NE 5
            }

            

            $PMs = @()
            foreach ($WmiObject in $WmiCandiddateObjs) {
                $Properties = @{
                    ComputerName = $WmiObject.PSComputerName
                    OperationalStatus = $WmiObject.OperationalStatus
                    StatusDescriptions = $WmiObject.StatusDescriptions
                    HealthState = $WmiObject.HealthState
                    iLOURL = $iLOInfo | ? PSComputerName -EQ $WmiObject.PSComputerName | Select-Object -ExpandProperty URL
                    SerialNumber = $HWinfo | ? PSComputerName -EQ $WmiObject.PSComputerName | Select-Object -ExpandProperty SerialNumber
                    ProductID = $HWinfo | ? PSComputerName -EQ $WmiObject.PSComputerName | Select-Object -ExpandProperty ProductID
                    Model = $HWinfo | ? PSComputerName -EQ $WmiObject.PSComputerName | Select-Object -ExpandProperty Model
                    # update 1.1.3: add 'IPAddress'
                    IPAddress = $PingResult | ? ComputerName -EQ $WmiObject.PSComputerName | Select-Object -ExpandProperty IPAddress
                }
                $PMs += New-Object -TypeName PSObject -Property $Properties
            }

            Write-Progress -Activity  "Collecting PM info ..." -Status "Done"

            # update 1.1.3: add 'IPAddress'
            if ($PassThru) {
                # Made PS object for every issue.
                $Result = @()

                foreach ($cpu in $CPUs) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $cpu.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $cpu.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "CPU"
                        Name = $cpu.Caption
                        HealthState = $HPSystemHealthTable[$cpu.HealthState]
                    }
                }
                
                foreach ($m in $Memory) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $m.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $m.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "Memory"
                        Name = $m.Caption
                        HealthState = $HPSystemHealthTable[$m.HealthState]
                    }
                }
                
                foreach ($mm in $MemoryModule) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $mm.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $mm.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "MemoryModule"
                        Name = $mm.Caption
                        HealthState = $HPSystemHealthTable[$mm.HealthState]
                    }
                }
                
                foreach ($f in $Fans) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $f.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $f.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "Fan"
                        Name = $f.Caption
                        HealthState = $HPSystemHealthTable[$f.HealthState]
                    }
                }
                
                foreach ($p in $Power) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $p.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $p.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "Power"
                        Name = $p.Caption
                        HealthState = $HPSystemHealthTable[$p.HealthState]
                    }
                }
                
                foreach ($pr in $PwrRedundantSet) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $pr.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $pr.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "PowerRedundant"
                        Name = $pr.Caption
                        HealthState = $HPRedundancyStatusTable[$pr.RedundancyStatus]
                    }
                }
                
                foreach ($n in $Network) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $n.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $n.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "Network"
                        Name = $n.Caption
                        HealthState = $HPSystemHealthTable[$n.HealthState]
                    }
                }
                
                foreach ($sa in $SmartArray) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $sa.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $sa.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "SmartArray"
                        Name = $sa.ElementName
                        HealthState = $sa.StatusDescriptions
                    }
                }
                
                foreach ($ac in $ArrayController) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $ac.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $ac.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "ArrayController"
                        Name = $ac.ElementName
                        HealthState = $HPOperationalStatusTable[$ac.OperationalStatus]
                    }
                }

                foreach ($pd in $PDisks) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $pd.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $pd.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "PyshicalDisk"
                        Name = $pd.ElementName
                        HealthState = $HPSAOperationalStatusTable[$pd.OperationalStatus]
                    }
                }

                foreach ($ilo in $iLOInfo) {
                    $Result += New-Object -TypeName PSObject -Property @{
                        ComputerName = $ilo.PSComputerName
                        SerialNumber = $HWinfo | ? PSComputerName -EQ $ilo.PSComputerName | Select-Object -ExpandProperty SerialNumber
                        Parts = "PyshicalDisk"
                        Name = $ilo.ElementName
                        HealthState = $HPSystemHealthTable[$ilo.HealthState]
                    }

                }
                

                return $Result

            }
            else {
                $PMs | fl ComputerName,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}},IPAddress,iLOURL,SerialNumber,ProductID, Model
                If ($CPUs) {$CPUs | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($Memory) {$Memory | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($MemoryModule) {$MemoryModule | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($Fans) {$Fans | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($Power) {$Power | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($PwrRedundantSet) {$PwrRedundantSet | ft -a PSComputerName,Caption,@{Name="RedundancyStatus";Expression={$HPRedundancyStatusTable[$_.RedundancyStatus]}}}
                If ($Network) {$Network | ft -a PSComputerName,Caption,@{Name="OperationalStatus";Expression={$HPOperationalStatusTable[$_.OperationalStatus]}},@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}}}
                If ($SmartArray) {$SmartArray | ft -a PSComputerName,ElementName,@{Name="OperationalStatus";Expression={$_.StatusDescriptions}}}
                If ($ArrayController) {$ArrayController | ft -a PSComputerName,ElementName,@{Name="OperationalStatus";Expression={$HPOperationalStatusTable[$_.OperationalStatus]}},BatteryStatus}
                If ($PDisks) {$PDisks | ft -a PSComputerName,ElementName,@{Name="StatusDescriptions";Expression={$HPSAOperationalStatusTable[$_.OperationalStatus]}}}
                If ($iLOInfo) {$iLOInfo | ft -a PSComputerName,Caption,StatusDescriptions,@{Name="HealthState";Expression={$HPSystemHealthTable[$_.HealthState]}},@{Name="NICCondition";Expression={$HPiLONICConditionTable[$_.NICCondition[0]]}}}
            }
        }
    }
}

Function Open-HPHomePage {
<#
.SYNOPSIS
Open HP homepage from MOR.

.DESCRIPTION
We cannot open HP homepage from PM because security policy deny open IE on it.
Use this command open homepage from MOR.

.PARAMETER ComputerName
Physical machine name.

.EXAMPLE
PS C:\> Open-HPHomePage -CN SHACHNVUT100
Open SHACHNVUT100 homepage from MOR.

.EXAMPLE
PS C:\> "bjbchnvut144","shachnvut100" | Open-HPHomePage
Open HP homepage for "bjbchnvut144","shachnvut100".
#>
    
    Param (
        [Alias("CN","ComputerName")]
        [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true)]
        [String] $Name
    )

    Process {
        $address = "https://" + $Name + ":2381"
        Start-Process $address
    }
    
}

Function Open-HPiLOPage {
<#
.SYNOPSIS
Open iLo from IE.

.DESCRIPTION
This function will open iLo from IE to save your time.

.PARAMETER ComputerName
Must be PM name.

.Inputs
Accept CenteralPMachine.

.EXAMPLE
PS C:\> Get-CenteralPMachine 458 | Open-HPiLOPage
Open iLo page for SHACHNVUT100.

.EXAMPLE
PS C:\> Open-HPiLOPage BJBCHNVUT144
Open iLo page for BJBCHNVUT144.
#>

    Param (
        [Alias("CN","ComputerName")]
        [Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
        [String] $Name
    )

    Process {
        $CenteralPM = Get-CenteralPMachine -Name $Name
        if ($CenteralPM) {
            $address = "https://" + $CenteralPM.ILoIpAddress
            Start-Process $address
        }
    }
}
# CHNMachineFunctions>

# <CHNOEAccountFunctions

Function Start-OCEProvisioning {
    <#
    .SYNOPSIS
    Provision CHN account for Microsoft OCE who don't have it.
    .DESCRIPTION
    This function will read CSV file in d:\temp\CSVGalltin folder and create AD account for it. It will be move to d:\temp\CSV_OLD folder if it done.
    CHN domain admin credential is necessary.
    .PARAMETER Credential
    Provide your CHN domain admin account for create OCE account. 
    #>
    Param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )

    $ObjCSVs = Read-CSV
    If ($ObjCSVs -ne $null) {
        $ObjCSVs | %{New-OCEAccount -ProvUserInfo $_ -Credential $Credential}
    }
    Else {
        Write-Host "No CSV file can be found!"
    }
}

Function New-OCEAccount {
    <#
    .SYNOPSIS
    Create CHN account for Microsoft OCE who don't have it.
    .DESCRIPTION
    This will create AD account for it. 
    CHN domain admin credential is necessary.
    .PARAMETER Credential
    Provide your CHN domain admin account for create OCE account. 
    #>
    Param (
        $ProvUserInfo,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )
    try {
        $ADUser = Get-ADUser -Identity $ProvUserInfo.SAMAccountName -Credential $Credential
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your admin account is avariable." -ForegroundColor Red
        Break
    }
    catch {
        Write-Host "Success: Provisioning user does not exist! Continue to create it." -ForegroundColor Green
    }

    If ($ADUser -eq $null) {
        $Name = $ProvUserInfo.givenName + " " + $ProvUserInfo.sn
        
        $NewADUser = New-ADUser -Name $Name -SamAccountName $ProvUserInfo.SAMAccountName -UserPrincipalName ($ProvUserInfo.SAMAccountName + "@CHN.SPONETWORK.COM") `
                    -GivenName $ProvUserInfo.GivenName -Surname $ProvUserInfo.sn -DisplayName $ProvUserInfo.DisplayName -Office $ProvUserInfo.physicalDeliveryOfficeName `
                    -Company $ProvUserInfo.company -Title $ProvUserInfo.title -EmailAddress $ProvUserInfo.mail `
                    -Path "OU=Employees,OU=People,DC=CHN,DC=SPONETWORK,DC=COM" -PassThru -ErrorAction Stop -Credential $Credential
        

        If ($NewADUser -ne $null) {
            Write-Host "Success: $($NewADUser.DistinguishedName) has been created successfully!" -ForegroundColor Green
            Add-OCEGroups -ADUser $NewADUser -Credential $Credential
            Enable-OCEAccount -SamId $NewADUser.SAMAccountName -Credential $Credential -WithOutNotification
        }
        Else {
            Write-Host "Error: Can not create provisioning user!" -ForegroundColor Red
        }
    }
    Else {
        Write-Host "Error: Existed user `"$($ADUser.DistinguishedName)`"." -ForegroundColor Red
        Break
    }

    $ADUser = $NewADUser | Get-ADUser -Properties EmailAddress
    $To = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList @($ADUser.EmailAddress,$ADUser.Name)
    $Subject = "Your new CHN accoount has been created"
    $Body = "Your CHN account has been created! Your Password will send in separate mail!`r`n`r`n" + 
            "USERNAME `r`n" + 
            "============================`r`n" +
            "$($ADUser.SamAccountName) `r`n" + 
            "============================`r`n`r`n" +
            "In order to logon Gallatin you need two accounts: `r`n" + 
            "•	CME account `r`n" + 
            "•	CHN account `r`n" + 
            "NOTE: Account information will be send via two different e-mails and may show up on different days.`r`n" + 
            "Once you have both accounts follow these directions to reset your passwords and to logon to Gallatin: `r`n" + 
            "http://sharepoint/sites/spo-service-readiness/SitePages/Requesting%20Debug%20Access%20For%20Gallatin.aspx"

    Send-NotificationMail -To $To -Subject $Subject -Body $Body
    
}

Function Add-OCEGroups {
    
    Param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADUser] $ADUser,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )

    $Groups = @()
    $Groups += "cn=OCE Users,ou=Roles,ou=Security,ou=Groups,dc=chn,dc=sponetwork,dc=com"
    $Groups += "cn=OCE_USA,ou=Roles,ou=Security,ou=Groups,dc=chn,dc=sponetwork,dc=com"

    
    try {
    Add-ADPrincipalGroupMembership -Identity $ADUser -MemberOf $Groups -Credential $Credential -ErrorAction Stop
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your admin account is avariable." -ForegroundColor Red
        Break
    }
    catch {
        Write-Host "WARNING: Could not add member to ADGroup `"OCE Users`" and `"OCE_USA`"!" -ForegroundColor Yellow
        
    }

}

Function Reset-OCEPassword {
    <#
    .SYNOPSIS
    Reset password for Microsoft OCE.
    .DESCRIPTION
    This function will reset AD account password and unblock it.
    CHN domain admin credential is necessary.
    .PARAMETER Credential
    Provide your CHN domain admin account for create OCE account. 
    #>
    Param (
        [String] $SamId,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )
    

    try {
        $ADUser = Get-ADUser -Identity $SamId -Properties EmailAddress
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "Error: Reset OCE password failed!" -ForegroundColor Red
        Write-Host "Error: Can not found user `"$SamId`"." -ForegroundColor Red
        Break
    }
    $PlainText = New-RandomPassword
    Write-Host "User password: $PlainText"
    try {
        $Password = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        
        Set-ADAccountPassword -Identity $ADUser -Reset -NewPassword $Password -Credential $Credential
        
        # Update 1.1.4: Change password at logon as FALSE due to MOR RDP policy changed.
        Set-ADUser -Identity $ADUser -ChangePasswordAtLogon $false -Credential $Credential

        Unlock-OCEAccount -SamId $SamId -Credential $Credential -WithOutNotification
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your account is avariable." -ForegroundColor Red
        Break
    }
    
    Write-Host "Success: Password has been reset." -ForegroundColor Green

    # Create mail message.
    $To = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList $ADUser.EmailAddress,$ADUser.Name
    $Subject = "Your CHN initial password"
    $Body = "Your Password has been reset. `n`n" +
            "Please use below password for first logon and change it. `n`n" + 
            "Password`n===================`n$PlainText`n===================`n`n" + 
            "NOTE: Your new password must be 10 characters and complex`r`n`r`nThe password must contain 3 of the following 4 characters types:`r`n`r`n" + 
            "Uppercase characters (A, B, C, D, E, ...)`r`n" + 
            "Lowercase characters (a, b, c, d, e, ...)`r`n" + 
            "Numerals (0, 1, 2, 3, ...)`r`n" + 
            "Special characters (@,!,_,#,..)`r`n`r`n" + 
            "REFERENCES:`r`n" + 
            "CME / CHN  Access Management Requests`n" +
            "http://sharepoint/sites/spo-service-readiness/SitePages/Requesting%20Debug%20Access%20For%20Gallatin.aspx"

    Send-NotificationMail -To $To -Subject $Subject -Body $Body
    Write-Host "Success: Password mail has been sent to user and you!" -ForegroundColor Green
}

Function Enable-OCEAccount {
    <#
    .SYNOPSIS
    Enable AD account for Microsoft OCE.
    .DESCRIPTION
    This function will reset AD account password and enable it.
    CHN domain admin credential is necessary.
    .PARAMETER Credential
    Provide your CHN domain admin account for create OCE account. 
    #>
    Param (
        [String] $SamId,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential,
        [Switch] $WithOutNotification
    )


    try {
        $ADUser = Get-ADUser -Identity $SamId -Properties EmailAddress -Credential $Credential
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your admin account is avariable." -ForegroundColor Red
        Break
    }
    catch {
        Write-Host "Error: Can not find OCE account '$SamId'." -ForegroundColor Red
        Break
    }

    Reset-OCEPassword -SamId $SamId -Credential $Credential

    
    try {
        Enable-ADAccount -Identity $ADUser -Credential $Credential -ErrorAction Stop
        
    }
    catch {
        Write-Host "Warning: Can not enable OCE account!" -ForegroundColor Yellow
        Break
    }

    try {
        Move-ADObject -Identity $ADUser -TargetPath "OU=Employees,OU=People,DC=CHN,DC=SPONETWORK,DC=COM" -Credential $Credential -ErrorAction Stop
    }
    catch {
        Write-Host "Warning: Can not move OCE account to OU 'Employees'! Make sure your account has the permission." -ForegroundColor Yellow
        Break
    }
    
    Write-Host "Success: OCE account has been enabled!" -ForegroundColor Green

    If ($WithOutNotification -eq $false) {
        $To = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList $ADUser.EmailAddress,$ADUser.Name
        $Subject = "Your CHN account has been enabled!"
        $Body = "Hi $($ADUser.Name),`t`n`t`n" + 
                "Your CHN account $($ADUser.SamAccountName) has been enabled!`t`n`t`n" +
                "Your initial password will be sent in separated mail.`t`n`t`n"
        Send-NotificationMail -To $To -Subject $Subject -Body $Body
    }
}

Function Disable-OCEAccount {
    <#
    .SYNOPSIS
    Disable AD account for Microsoft OCE.
    .DESCRIPTION
    This function will disable AD account.
    CHN domain admin credential is necessary.
    .PARAMETER Credential
    Provide your CHN domain admin account for create OCE account. 
    #>
    Param (
        [String] $SamId,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential,
        [Switch] $WithOutNotification
    )


    try {
        $ADUser = Get-ADUser -Identity $SamId -Properties EmailAddress -Credential $Credential
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your admin account is avariable." -ForegroundColor Red
        Break
    }
    catch {
        Write-Host "Error: Can not find OCE account '$SamId'." -ForegroundColor Red
        Break
    }

    Try {
        Disable-ADAccount -Identity $ADUser -Credential $Credential -ErrorAction Stop
    }
    catch {
        Write-Host "Error: Can not disable OCE account '$SamId'." -ForegroundColor Red
    }

    Try {
        Move-ADObject -Identity $ADUser -TargetPath "OU=GRAVEYARD,DC=CHN,DC=SPONETWORK,DC=COM" -Credential $Credential -ErrorAction Stop
    }
    catch {
        Write-Host "Warning: Your account didn't have permission to move the disabled OCE account to OU 'GRAVEYARD'." -ForegroundColor Yellow
    }
    Write-Host "Success: OCE account has been Disabled!" -ForegroundColor Green

    If ($WithOutNotification -eq $false) {
        $To = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList $ADUser.EmailAddress,$ADUser.Name
        $Subject = "Your CHN account has been disabled!"
        $Body = "Hi $($ADUser.Name),`t`n`t`n" + 
                "Your CHN account $($ADUser.SamAccountName) has been disabled!`t`n`t`n" +
                "Once you want to enable your account, please refer below link:`t`n" +
                "http://sharepoint/sites/spo-service-readiness/SitePages/Requesting%20Debug%20Access%20For%20Gallatin.aspx"
        Send-NotificationMail -To $To -Subject $Subject -Body $Body
    }
}

Function Unlock-OCEAccount {

<#
.SYNOPSIS
Unlock CHN account.

.DESCRIPTION
Please use domain admin accunt for crednetial.

.PARAMETER SamId
SAM account name.

.PARAMETER Credential
CHN domain admin account.

.PARAMETER WithOutNotification
No notification mail send if it switch on.

.EXAMPLE
PS C:\> Unlock-OCEAccount -SamId oe-songwende -Credential $cred -WithOutNotification
Unlock user 'oe-songwende' without notification mail.
#>

    Param (
        [String] $SamId,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential,
        [Switch] $WithOutNotification
    )

    
    try {
        $ADUser = Get-ADUser -Identity $SamId -Properties EmailAddress
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "Error: Unlock OCE account failed!" -ForegroundColor Red
        Write-Host "Error: Can not found user `"$SamId`"." -ForegroundColor Red
        Break
    }
    try {
        Unlock-ADAccount -Identity $ADUser -Credential $Credential
    }
    catch [System.Security.Authentication.AuthenticationException] {
        Write-Host "Error: Your credential not valid! Please verify your admin account is avariable." -ForegroundColor Red
        Break
    }
    catch {
        Write-Host "Error: Can not unlock OCE account '$SamId'." -ForegroundColor Red
        break
    }

    Write-Host "Success: OCE account '$SamId' has been unblocked!" -ForegroundColor Green

    If ($WithOutNotification -eq $false) {
        $To = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList $ADUser.EmailAddress,$ADUser.Name
        $Subject = "Your CHN account has been unlocked!"
        $Body = "Hi $($ADUser.Name),`t`n`t`n" + 
                "Your CHN account $($ADUser.SamAccountName) has been unlocked!`t`n`t`n"
        Send-NotificationMail -To $To -Subject $Subject -Body $Body
        Write-Host "Success: Notification mail sent to user." -ForegroundColor Green
    }
}

Function Set-OpsAccount {
<#
.SYNOPSIS
This is combine command of Reset-OCEPassword, Enable and Disable OCEAccount.

.PARAMETER Account
The SAM account name of Active Directory.

.PARAMETER Reset
Reset switch is used for invoke Reset-OCEPassword.

.PARAMETER Enable
Enable switch is used for invoke Enable-OCEAccount.

.PARAMETER Disable
Disable switch is used for invoke Disable-OCEAccount.

.PARAMETER Credential
The credential which is used for AD execution.

.EXAMPLE
Set-OpsAccount oe-songwende -Reset -Credential $cred
To reset Wind's account password.

.EXAMPLE
Set-OpsAccount oe-songwende -Disable -Credential $cred
To disable Wind's account.

.EXAMPLE
Set-OpsAccount oe-songwende -Enable -Credential $cred
To enable Wind's account.
#>

    
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [String] $Account,
        [Parameter(Mandatory=$true,ParameterSetName="Reset")]
        [Switch] $Reset,
        [Parameter(Mandatory=$true,ParameterSetName="Enable")]
        [Switch] $Enable,
        [Parameter(Mandatory=$true,ParameterSetName="Disable")]
        [Switch] $Disable,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )

    Switch ($Pscmdlet.ParameterSetName) {

        "Reset" {
            Reset-OCEPassword -SamId $Account -Credential $Credential
            break
        }

        "Enable" {
            Enable-OCEAccount -SamId $Account -Credential $Credential
            break
        }

        "Disable" {
            Disable-OCEAccount -SamId $Account -Credential $Credential
            break
        }

    }


}
# CHNOEAccountFunctions>

# <CHNPerformance
Function Get-OSPerformance {
<#
.SYNOPSIS
Get OS performance once in batch.

.DESCRIPTION
Just query performance one time for mutiple machines.

.PARAMETER Types
"CPU", "Memory", "Network" and "Disk" are optional. Default is all.

.PARAMETER ComputerName
Remote computer name. 

.PARAMETER Role
The query computers by role. VM is allowed only.

.EXAMPLE
PS C:\> Get-OSPerformance -Role USR
Get OS performance for all USRs once.
#>


    Param (
        [ValidateSet("CPU","Memory","Network","Disk")]
        [String[]] $Types,
        [Alias("CN")][Parameter(Mandatory=$true,ParameterSetName="ByComputer")]
        [String[]] $ComputerName,
        [Parameter(Mandatory=$true,ParameterSetName="ByRole")]
        [String] $Role
    )

    if($PSCmdlet.ParameterSetName -eq "ByRole") {
        $ComputerName = Get-CenteralVM -Role $Role -ErrorAction SilentlyContinue | select -ExpandProperty Name
    }

    
    if (!$Types) {
        $Types = "CPU","Memory","Network","Disk"
    }

    if ($ComputerName) {
        
        $session = New-PSSession -ComputerName $ComputerName -ErrorAction SilentlyContinue
        $SessionHost = $session.ComputerName
        $BadHost = $ComputerName | ?{$_ -notin $SessionHost}
        if ($BadHost) {
            if ($BadHost.Count -eq $ComputerName.Count) {
                Write-Host "No available computer!" -ForegroundColor Red
                break
            }
            Write-Host "Computer $(Convert-ArrayAsString $BadHost) not available!" -ForegroundColor Yellow
        }
    }
    else{
        Write-Host "Cannot find any computer!" -ForegroundColor Red
        break
    }

    # query performance
    $InvokeCommand = "Invoke-Command -ScriptBlock `$function:GetOSPerformance -ArgumentList (,`$Types) -Session `$Session -AsJob"
    Invoke-Expression -Command $InvokeCommand | Out-Null
    $result = Watch-Jobs -Activity "Query performance"
    if ($session) { Remove-PSSession $session }

    # Sort result
    $CPU = @()
    $Memory = @()
    $Network = @()
    $Disk = @()
    foreach ($r in $result) {
        switch ($true) {
            {$Types -contains "CPU"} {
                $CPU += $r['CPU']
            }

            {$Types -contains "Memory"} {
                $Memory += $r['Memory']
            }

            {$Types -contains "Network"} {
                $Network += $r['Network']
            }

            {$Types -contains "Disk"} {
                $Disk += $r['Disk']
            }
        }
    }

    $CPU | ft -AutoSize ComputerName, Type, @{Expression="Utilization";FormatString="P2"} | Out-Host
    $Memory | ft -AutoSize ComputerName, Type, @{Name="Total";Expression={$_.Total / 1GB};FormatString="# GB"}, `
                @{Name="Available";Expression={$_.Available / 1GB};FormatString="# GB"}, `
                @{Expression="Utilization";FormatString="P2"} | Out-Host
    $Network | ft -AutoSize ComputerName, Type, @{Name="BandWidth";Expression={$_.BandWidth / 1000};FormatString="# Kbps"}, `
                @{Name="ReceivedBps";Expression={$_.ReceivedBps / 1000};FormatString="0 KBps"}, `
                @{Name="SentBps";Expression={$_.SentBps / 1000};FormatString="0 KBps"}, `
                @{Name="TotalBps";Expression={$_.ReceivedBps / 1000};FormatString="0 KBps"} | Out-Host
    $Disk | ft -AutoSize ComputerName, Type, @{Expression="Activity";FormatString="P2"} | Out-Host
    # return $CPU,$Memory,$Network,$Disk
    
}

Function Watch-OSPerformance {
<#
.SYNOPSIS
Display OS performance in console real time. Just watch one machine.

.DESCRIPTION
It will caculate max, min and average utilization of CPU, memory and disk activity, network throughput.

.PARAMETER Types
"CPU", "Memory", "Network" and "Disk" are optional. Default is all.

.PARAMETER ComputerName
Remote computer name. 

.PARAMETER Interval
The time interval between query. Default is 1 second.

.PARAMETER Duration
The query duration. Default is 60 seconds.

.EXAMPLE
PS C:\> Watch-OSPerformance -CN bjbchnvut144
Watch os performance for bjbchnvut144 in 1 minute.
#>

    Param (
        [ValidateSet("CPU","Memory","Network","Disk")]
        [String[]] $Types,
        [Int] $Interval = 1,
        [ValidateRange(10,3600)]
        [Int] $Duration = 60,
        [Alias("CN")]
        [String] $ComputerName
    )

    if (!$Types) {
        $Types = "CPU","Memory","Network","Disk"
    }

    if ($ComputerName) {
        try {
            $session = New-PSSession -ComputerName $ComputerName
        }
        catch {
            Write-Host $_
            break
        }
    
    }

    
    if ($session) {
        $BasicInfo = Invoke-Command -ScriptBlock $Function:GetOSInfo -Session $session
        $InvokeCommand = "Invoke-Command -ScriptBlock `$function:GetOSPerformance -ArgumentList (,`$Types) -Session `$Session"
    }
    else {
        $BasicInfo = GetOSInfo
        $InvokeCommand = "GetOSPerformance -Types `$Types"
    }

    Clear-Host

    $now = Get-Date
    $EndTime = $now.AddSeconds($Duration)
    $remaining = $EndTime - $now

     do {
        Start-Sleep -Seconds $Interval
        $remaining = $EndTime - (Get-Date)
        $result = Invoke-Expression -Command $InvokeCommand
        $hash = AddResultInHashTable -InputResult $result -hash $hash
        $outputs = CalculateOutputs -hash $hash -Outputs $outputs
        DisplayOutputs -Outputs $outputs -ComputerInfo $BasicInfo
        if ($remaining -gt 0) { Write-Host ("remaining: {0:hh\:mm\:ss}" -f $remaining) }
        
    } while ($remaining.TotalSeconds -gt 0)
    
    if ($session) { Remove-PSSession $session }
    
}

Function Test-OSPerformance {
<#
.SYNOPSIS
Sample test OS performance in a while and record it.

.DESCRIPTION
This function support in batch mode by role or machine names.
Save result in files.

.PARAMETER ComputerName
Host name.

.PARAMETER Role
To query VM hosts from role.

.PARAMETER Types
Test performance items include "CPU","Memory","Network" and "Disk".

.PARAMETER Interval
The time interval between query. Default is 1 second.

.PARAMETER Duration
The query duration. Default is 60 seconds.

.PARAMETER Path
The result stored path.

.EXAMPLE
PS C:\> Test-OSPerformance -Role SEC
Test OS performance for all SEC VMs.

.EXAMPLE
PS C:\> Test-OSPerformance -CN bjbchnvut144 -Types CPU
Just test CPU for bjbchnvut144 in 1 minute.
#>

    Param (
        [Alias("CN")]
        [Parameter(Mandatory=$true,ParameterSetName="ByHost")]
        [String[]] $ComputerName,
        [Parameter(Mandatory=$true,ParameterSetName="ByRole")]
        [String] $Role,
        [ValidateSet("CPU","Memory","Network","Disk")]
        [String[]] $Types,
        [Int] $Interval = 1, # seconds
        [ValidateRange(1,3600)]
        [Int] $Duration = 60, # seconds
        [String] $Path = "\\chn\tools\Performance"
    )

    
    $separator = "=" * $PSCmdlet.Host.UI.RawUI.WindowSize.Width

    if (!(Test-Path $Path)) {
        Write-Host "Path $Path is not valid!" -ForegroundColor Red
        break
    }

    if(!$Types) {
        $Types = @("CPU","Memory","Network","Disk")
    }

    # create save path
    $Date = Get-Date
    $Folder = Join-Path $Path ("TestOSPerformance", $Date.ToString("yyyyMMddhhmmss") -join "_")
    
    try {
        New-Item -ItemType Directory -Path $Folder -ErrorAction Stop | Out-Null
    }
    catch { 
        Write-Host $_ -ForegroundColor Red
        break
    }

    # Convert role to computername
    if ($PSCmdlet.ParameterSetName -eq "ByRole") {
        $ComputerName = Get-CenteralVM -Role $Role | select -ExpandProperty Name
    }

    # Create PS session and identify which computers cannot connect.
    $PSSession = New-PSSession -ComputerName $ComputerName -ErrorAction SilentlyContinue
    if(!$PSSession) {
        Write-Host "Cannot create any PS session!" -ForegroundColor Red
        break
    }
    if($PSSession.Count -lt $ComputerName.Count) {
        $SessionHost = $PSSession.ComputerName
        $BadHost = $ComputerName | ?{ $_ -notin $SessionHost}
        Write-Host "Cannot create PS session for $(Convert-ArrayAsString $BadHost)" -ForegroundColor Yellow
    }

    # Initialize root folder
    try {
        $RecordFilePath = InitializeRootFolder -Path $Folder -Type $Types -ComputerName $ComputerName
    }
    catch {
        Write-Host "Create root folder '$Folder' failed!" -ForegroundColor Red
    }
    $StatisticPath = $RecordFilePath['StatisticPath']
    $RecordPath = $RecordFilePath['RecordPath']

    Clear-Host
    Write-Host ("`n" * 12)
    Write-Host $separator
    Write-Host "Querying performance for below hosts:"
    Write-Host (Convert-ArrayAsString $ComputerName) -ForegroundColor Green
    Write-Progress -Activity "Test OS performance" -Status Initializing -Id 1


    # Get basic OS info
    $job = Invoke-Command -ScriptBlock $Function:GetOSInfo -Session $PSSession -JobName OSInfo
    $OsInfo = $job | Watch-Jobs -Activity "Getting OS info..."

    WriteOSInfo -Path $Folder -OSInfo $OsInfo

    # query performance
    $now = Get-Date
    $EndTime = $now.AddSeconds($Duration)
    $remaining = $EndTime - $now
    $InvokeCommand = "Invoke-Command -ScriptBlock `$function:GetOSPerformance -ArgumentList (,`$Types) -Session `$PSSession -AsJob"

    do {
        $escape = $Duration - $remaining.TotalSeconds
        $Percent = $escape / $Duration * 100
        if ($Percent -gt 100) {$Percent = 100}
        Write-Progress -Activity "Test OS performance" -Status Testing -Id 1 -SecondsRemaining $remaining.TotalSeconds -PercentComplete $Percent
        Invoke-Expression -Command $InvokeCommand | Out-Null
        $result = Watch-Jobs -Activity "Query performance"
        Write-Progress -Activity "Write records" -Status Writing
        $Statistic = CalculateStatistics -InputObject $result -Statistic $Statistic
        WriteRecords -Path $RecordPath -InputObject $result
        WriteStatisticRecords -Path $StatisticPath -InputObject $Statistic
        Start-Sleep -Seconds $Interval
        $remaining = $EndTime - (Get-Date)

    } while ($remaining.TotalSeconds -gt 0)

    Remove-PSSession $PSSession
    
    Write-Host "Show statistics:"
    ShowStatistics -Statistics $Statistic -OsInfo $OsInfo
    Write-Host $separator
    Write-Host "Test result saved in '$Folder'`n" -ForegroundColor Green

}

Function Get-TestPerformanceResult {
<#
.SYNOPSIS
Get test OS performance result from saved path.

.DESCRIPTION
This function will retrive statistic result from saved path and reurn or show.

.PARAMETER Path
The test OS performance path which prefix is 'TestOSPerformance'.

.PARAMETER Show
Just display statistics in console not return object.

.Outputs
Return statistic object.
Display statistic result in console if 'Show' switch is present.

.EXAMPLE
PS C:\> $s = Get-TestPerformanceResult '\\chn\tools\Performance\TestOSPerformance_20170813103420'
Save performance result for 'TestOSPerformance_20170813103420' in variable s.

.EXAMPLE
PS C:\> Get-TestPerformanceResult '\\chn\tools\Performance\TestOSPerformance_20170813103420' -Show
Display performance result in console.
#>

    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String] $Path,
        [Alias("CN")]
        [String[]] $ComputerName,
        [Switch] $Show
    )

    if ($ComputerName) {
        try {$ChildPath = Get-ChildItem -Path $Path -ErrorAction Stop | select -ExpandProperty FullName}
        catch {Write-Host $_; break}
        $ValidPath = $()
        foreach ($computer in $ComputerName) {
            $ValidPath += $ChildPath -match $computer
        }

        if (!$ValidPath) {
            Write-Host "No valid result for your query computers $(Convert-ArrayAsString $ComputerName)!" -ForegroundColor Red
            break
        }
    }
    else {
        $ValidPath = $Path
    }

    # try to find statistics.csv in every computer folder.
    try {
        $StatisticFile = Get-ChildItem -Path $ValidPath -Include "Statistics.csv" -Recurse -ErrorAction Stop
    }
    catch {
        Write-Host $_
        break
    }

    # import statistics
    $statistic = $StatisticFile.FullName | Import-Csv

    if ($Show.IsPresent) {
        try {
            $InfoFile = Get-ChildItem -Path $ValidPath -Include "BasicInfo.csv" -Recurse -ErrorAction Stop
        }
        catch {
            Write-Host $_
            break
        }

        $OsInfo = $InfoFile.FullName | Import-Csv

        ShowStatistics -Statistics $statistic -OsInfo $OsInfo

    }
    else {
        return $statistic
    }

}
# CHNPerformance>

# <CHNTenantFunctions
Function Disable-Tenants {
<#
    .SYNOPSIS
    Disable tenant in batch.

    .DESCRIPTION
    Diable tenant in batch mode. 
    Sometimes, the government will disable unregistered web site before important festival. 
    They always provide several RULs. Disable them will cost a lot of time.
    Use this function can save time for us.

    .PARAMETER SiteURLs
    Can accept mutilpe URLs as array.
    The public URL or internal URL are accepted both. Prefix 'http://' or 'https://' are not necessary. 
    If you entered public URL, it will add 'http://'. Otherwise, it add 'https://'.

    .PARAMETER TenantIDs
    Can accept mutilpe tenant IDs as array.
    Tenant ID must be vaild in Centeral.

    .EXAMPLE
    Disable-Tenants -SiteURLs spotest.sharepoint.cn
    Disable a tenant through site collection.

    .EXAMPLE
    Disable-Tenants -SiteURLs "spotest.sharepoint.cn","test21vianet360test.sharepoint.cn"
    Disable two tenants through site collections.

    .EXAMPLE
    Disable-Tenants -TenantIDs "123456","434532"
    Disable two tenants through tenant IDs.
#>

    Param (
        [Parameter(Mandatory=$true,
                    ParameterSetName="ByTenants")]
        [String[]] $TenantIDs,
        [Parameter(Mandatory=$true,
                    ParameterSetName="BySites")]
        [String[]] $SiteURLs
    )

    $SiteCollections = @()
    $InvalidSiteURLs = @()

    $Tenants = @()
    $InvalidTenantIDs = @()

    If ($PSCmdlet.ParameterSetName -eq "BySites") {
        
        foreach ($site in $SiteURLs) {
            $Site = Format-TenantUri -address $site
            $SiteCollection = $null
            $SiteCollection = Get-CenteralSiteCollection -Uri $Site
            If ($SiteCollection -ne $null) {
                $SiteCollections += $SiteCollection
            }
            else {
                $InvalidSiteURLs += $site
            }
        }

        If ($InvalidSiteURLs.Count -gt 0) {
            Write-Host "`nThe site collection `'$($InvalidSiteURLs -join "', '")`' can't be found in Centeral." -ForegroundColor Yellow
            Write-Host "Please check it manually!`n" -ForegroundColor Yellow
        }

        If ($SiteCollections.count -gt 0) {
            $TenantIDs = ($SiteCollections).Tenantid
        }
        else {
            Write-Host "No vaild site collectin be found!"
            Exit
        }
    }

    #Get tenants by tenantIDs.
    foreach ($id in $TenantIDs) {
        $Tenant = $null
        $Tenant = Get-CenteralTenant -Identity $id
        If ($Tenant -ne $null) {
            $Tenants += $Tenant
        }
        else {
            $InvalidTenantIDs += $id
        }
    }

    If ($InvalidTenantIDs.Count -gt 0) {
        Write-Host "The tenants $($InvalidTenantIDs -join ", ") can't be found in Centeral." -ForegroundColor Yellow
        Write-Host "Please check it manually!`n" -ForegroundColor Yellow
    }

    If ($Tenants.Count -gt 0) {
        
        $FTObjects = @()
        foreach ($t in $Tenants) {
            $formatObj = New-Object -TypeName PSObject -Property @{
                TenantId = $t.Tenantid;
                Urn = $t.Urn;
                PortalUrl = $t.PortalUrl;
                State = $t.State;
                PublicUrl = ""
            }
            If ($PSCmdlet.ParameterSetName -eq "BySites") {
                $formatObj.PublicUrl = ($SiteCollections | ?{$_.Tenantid -eq $formatObj.Tenantid}).PublicUrl
            }
            $FTObjects += $formatObj
        }


        Write-Host "Do you want to disable below tenants?"
        $FTObjects | ft -a Tenantid, Urn, PortalUrl, PublicUrl, State
        $Proceed = Read-Host -Prompt "Yes or No (Default: No)?"
        If ($Proceed -ne "Yes") { Write-Host "The operation has been aborted!";Break }


        $pxurl = "http://GM.CHN.SPONETWORK.COM/tenantmgr.asmx"
        $prx = new-webserviceproxy -uri $pxurl -usedefaultcredential

         $Tenants | %{$prx.DisableTenant($_.Urn)}

         Watch-TenantStatus -tenants $Tenants -Disabled
    }
    else {
        Write-Host "No vaild tenants be found!"
        Break
    }
    
}

Function Enable-Tenants {
    <#
    .SYNOPSIS
    Enable tenant in batch.

    .DESCRIPTION
    Diable tenant in batch mode. 
    Sometimes, the government will enable unregistered web site before important festival. 
    They always provide several RULs. Enable them will cost a lot of time.
    Use this function can save time for us.

    .PARAMETER SiteURLs
    Can accept mutilpe URLs as array.
    The public URL or internal URL are accepted both. Prefix 'http://' or 'https://' are not necessary. 
    If you entered public URL, it will add 'http://'. Otherwise, it add 'https://'.

    .PARAMETER TenantIDs
    Can accept mutilpe tenant IDs as array.
    Tenant ID must be vaild in Centeral.

    .EXAMPLE
    Enable-Tenants -SiteURLs spotest.sharepoint.cn
    Enable a tenant through site collection.

    .EXAMPLE
    Enable-Tenants -SiteURLs "spotest.sharepoint.cn","test21vianet360test.sharepoint.cn"
    Enable two tenants through site collections.

    .EXAMPLE
    Enable-Tenants -TenantIDs "123456","434532"
    Enable two tenants through tenant IDs.
#>

    Param (
        [Parameter(Mandatory=$true,
                    ParameterSetName="ByTenants")]
        [String[]] $TenantIDs,
        [Parameter(Mandatory=$true,
                    ParameterSetName="BySites")]
        [String[]] $SiteURLs
    )

    $SiteCollections = @()
    $InvalidSiteURLs = @()

    $Tenants = @()
    $InvalidTenantIDs = @()

    If ($PSCmdlet.ParameterSetName -eq "BySites") {
        
        foreach ($site in $SiteURLs) {
            $Site = Format-TenantUri -address $site
            $SiteCollection = $null
            $SiteCollection = Get-CenteralSiteCollection -Uri $Site
            If ($SiteCollection -ne $null) {
                $SiteCollections += $SiteCollection
            }
            else {
                $InvalidSiteURLs += $site
            }
        }

        If ($InvalidSiteURLs.Count -gt 0) {
            Write-Host "`nThe site collection `'$($InvalidSiteURLs -join "', '")`' can't be found in Centeral." -ForegroundColor Yellow
            Write-Host "Please check it manually!`n" -ForegroundColor Yellow
        }

        If ($SiteCollections.count -gt 0) {
            $TenantIDs = ($SiteCollections).Tenantid
        }
        else {
            Write-Host "No vaild site collectin be found!"
            Exit
        }
    }

    #Get tenants by tenantIDs.
    foreach ($id in $TenantIDs) {
        $Tenant = $null
        $Tenant = Get-CenteralTenant -Identity $id
        If ($Tenant -ne $null) {
            $Tenants += $Tenant
        }
        else {
            $InvalidTenantIDs += $id
        }
    }

    If ($InvalidTenantIDs.Count -gt 0) {
        Write-Host "The tenants $($InvalidTenantIDs -join ", ") can't be found in Centeral." -ForegroundColor Yellow
        Write-Host "Please check it manually!`n" -ForegroundColor Yellow
    }

    If ($Tenants.Count -gt 0) {
        
        $FTObjects = @()
        foreach ($t in $Tenants) {
            $formatObj = New-Object -TypeName PSObject -Property @{
                TenantId = $t.Tenantid;
                Urn = $t.Urn;
                PortalUrl = $t.PortalUrl;
                State = $t.State;
                PublicUrl = ""
            }
            If ($PSCmdlet.ParameterSetName -eq "BySites") {
                $formatObj.PublicUrl = ($SiteCollections | ?{$_.Tenantid -eq $formatObj.Tenantid}).PublicUrl
            }
            $FTObjects += $formatObj
        }


        Write-Host "Do you want to enable below tenants?"
        $FTObjects | ft -a Tenantid, Urn, PortalUrl, PublicUrl, State
        $Proceed = Read-Host -Prompt "Yes or No (Default: No)?"
        If ($Proceed -ne "Yes") { Write-Host "The operation has been aborted!";Break }


        $pxurl = "http://GM.CHN.SPONETWORK.COM/tenantmgr.asmx"
        $prx = new-webserviceproxy -uri $pxurl -usedefaultcredential

         $Tenants | %{$prx.EnableTenant($_.Urn)}

         Watch-TenantStatus -tenants $Tenants -Enabled
    }
    else {
        Write-Host "No vaild tenants be found!"
        Break
    }
}

Function Find-TenantByDomain {
<#
    .SYNOPSIS
    Look up tenant by domain.

    .DESCRIPTION
    It's looking for tenant by customer public domain or o365 domain.

    .PARAMETER Domain
    Customer's public domain or Tenant name.

    .EXAMPLE
    Find-TenantByDomain songwende.site
    Look for tenant which doamin is songwende.site.

    .EXAMPLE
    Find-TenantByDomain -Domain spotest.partner.onmschina.cn
    Look for tenant which o365 account is spotest.partner.onmschina.cn.

    .EXAMPLE
    Find-TenantByDomain -Domain spotest.sharepoint.cn
    Look for tenant which SPO portal is spotest.sharepoint.cn.
#>

    Param (
        [Parameter(Mandatory=$true)]
        [String] $Domain
    )

    $SQLModule = Get-Module SQLPS
    If (!$SQLModule) {Import-Module SQLPS -WarningAction SilentlyContinue}
    If ($? -eq $false) {
        Write-Debug "Cannot load SQLPS module! exit."
        Exit 1
    }

    #Locate principal CenteralTenant DB
    $CenteralTenantDBs = Get-CenteralDB -Name CenteralTenant | ? MirrorId -ne 0
    $SQLHosts = $CenteralTenantDBs.SqlHostName

    Foreach ($SQLhost in $SQLHosts) {
        $DBPorperty = Invoke-Sqlcmd -HostName $SQLhost -serverinstance $SQLHost -Query "Select DATABASEPROPERTYEX('CenteralTenant','Status') Status"
        If ($DBPorperty.Status -eq "ONLINE") {
            Write-Debug "Find principal CenteralTenant database host [$SQLhost]." 
            $SQLserver = $SQLhost
            Break
        }
    }
    If (!$SQLserver) {
        Write-Debug "Cannot find principal database!"
        Exit 1
    }

    $Query = "select * from [dbo].[fqdn] where Name = '$Domain'"
    $Result = Invoke-Sqlcmd -HostName $SQLserver -ServerInstance $SQLserver -DataBase CenteralTenant -Query $Query
    If ($Result) {
        Write-Debug "Find Record for $Domain in CenteralTenant DB."
        $CenteralTenant = Get-CenteralTenant -Identity $Result.TenantId
    }

    Return $CenteralTenant

}
# CHNTenantFunctions>

# <CHNTopologyFunctions
Function Get-CHNTopology {

<#
.SYNOPSIS
This function is used for get topology including VM, PM, farm, network, zone, loadbalancer, vlan and VM IP from current environment or saved topology.

.DESCRIPTION 
Get toplogy from current environment or sotre and return a hash table as result.
Topoloyg types include "VMs","PMs","Farms","Networks","Zones","NLBs","Vlans" and "VMIPs". "All" is default. 
Topology information will be stored at \\chn\tools\topology folder if parameter 'Export' presents for saving current topology.
'Last' and 'Version' are used to get history topology from store. Use Get-CHNTopologyVersions to list all of versions.

.PARAMETER Export
Export toplogy to "\\chn\tools\topology".

.PARAMETER Types
Including "VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs" and "All". "All" is default.

.PARAMETER CurrentTopology
Get topology from current environment.

.PARAMETER Version
Get topology from store by version.

.PARAMETER Last
Get topology from store by last number.

.OUTPUTS
Return a hash table.

.EXAMPLE
C:\PS> Get-CHNTopology -CurrentTopology
Obtain topology from current environment without save.

.EXAMPLE
C:\PS> Get-CHNTopology -CurrentTopology -Export
Obtain topology from current environment and save.

.EXAMPLE
C:\PS>  Get-CHNTopology -Types "VMs","PMs" -CurrentTopology -Export
Just obtain VMs and PMs topology  from current environment and save.

.EXAMPLE
C:\PS>  Get-CHNTopology -Version 20161114125913
Obtain topology by version 20161114125913 from store.

.EXAMPLE
C:\PS>  Get-CHNTopology -Last 1
Obtain the letast topology from store.

#>
    
    Param ( 
        [ValidateSet("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs","All")]
        [String[]] $Types="All",
        [Parameter(ParameterSetName="Current")]
        [Switch] $CurrentTopology,
        [Parameter(ParameterSetName="Current")]
        [Switch] $Export,
        [Parameter(Mandatory=$true,ParameterSetName="Legacy_Version",Position=0)]
        [String] $Version,
        [Parameter(Mandatory=$true,ParameterSetName="Legacy_Number",Position=0)]
        [ValidateScript({$_ -ge 1})]
        [Int] $Last
    )
    
    

    $Vlans = $NetWorks = $Farms = $VMIPs = $VMs = $PMs = $Zones = $NLBs = @()

    If ($Types -eq "All") {
        $Types = @("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs")
    }

    $Topology = @{}

    # Get current topology
    If ($PSCmdlet.ParameterSetName -eq "Current") {

        Write-Debug "Current topology is inspecting ..."

        Switch ($Types) {

            "VMs" {
                $Vlans = Get-CenteralVlan
                $NetWorks = Get-CenteralNetwork
                $Farms = Get-CenteralFarm
                $VMIPs = $Vlans | %{Get-CenteralManagedIP -VlanId $_.vlanid}
                $VMs = $NetWorks | %{Get-CenteralVM -Network $_} | ? PMachineid -gt 0 | sort NetworkId,Role,FarmId

                foreach ($vm in $VMs) {
                    $IP = ($VMIPs | ? vmid -EQ $vm.VMachineId).Address
                    $Datacenter = ($NetWorks | ? NetworkId -EQ $vm.NetworkId).DataCenter
                    $FarmRole = ($Farms | ? FarmId -EQ $vm.FarmId).Role
                    $FarmState = ($Farms | ? FarmId -EQ $vm.FarmId).State
                    $NoteProperties = @{
                        IPAddress = $IP
                        DataCenter = $Datacenter
                        FarmRole = $FarmRole
                        FarmState = $FarmState
                    }
                    $vm | Add-Member -NotePropertyMembers $NoteProperties

                }
                $Topology.Add("VMs",$VMs)
            }

            "PMs" {
                $PMs = Get-CenteralPmachine
                $Topology.Add("PMs",$PMs)
            }

            "Farms" {
                If (!$Farms) { $Farms = Get-CenteralFarm }
                $Topology.Add("Farms",$Farms)
            }

            "Networks" {
                If (!$NetWorks) { $NetWorks = Get-CenteralNetwork }
                $Topology.Add("Networks",$NetWorks)
            }

            "Zones" {
                $Zones = Get-CenteralZone -IncludeDefaultZone -IncludePhysicalZones
                $Topology.Add("Zones",$Zones)
            }

            "NLBs" {
                $NLBs = Get-CenteralLoadBalancer
                $Topology.Add("NLBs",$NLBs)
            }

            "Vlans" {
                If (!$Vlans) { $Vlans = Get-CenteralVlan }
                $Topology.Add("Vlans",$Vlans)
            }

            "VMIPs" {
                If (!$VMIPs) { $VMIPs = Get-CenteralVlans | %{Get-CenteralManagedIP -VlanId $_.vlanid} }
                $Topology.Add("VMIPs",$VMIPs)
            }
        }

        # Export to XML
        If ($Export) {
            #$ExportPath = "D:\WindSONG\Documents\Topology"
            $ExportPath = "\\CHN.SPONETWORK.COM\Tools\Topology"
            If (!(Test-Path $ExportPath)) {
                Write-Debug "Can not find $ExportPath" -ForegroundColor Red
                Return $Topology
            }
            $Date = Get-Date
            $StrDate = Get-Date -Date $Date -Format "yyyyMMdd"
            $StrDateTime = Get-Date -Date $Date -Format "yyyyMMddhhmmss"
            $StoreFolder = Join-Path $ExportPath $StrDate
            If (!(Test-Path $StoreFolder)) {
                try { New-Item -Path $StoreFolder -ItemType Directory | Out-Null }
                catch {
                    Write-Debug "Can not create folder $StoreFolder!"
                    $_
                    Return $Topology
                }
            }

            foreach ($Item in @($Topology.Keys)) {
                $xml = $Item + "_" + $StrDateTime + ".xml"
                $Path = Join-Path $StoreFolder $xml
                $Topology[$Item] | Export-Clixml -Path $Path
                If (Test-Path $Path) {
                    Write-Host "Export: '$Path'" -ForegroundColor Green
                }
            }
        }

        Write-Debug "Current topology done."

    }

    # Get legacy topology by number or version from store.
    If ($PSCmdlet.ParameterSetName -match "Legacy") {

        $Versions = Get-CHNTopologyVersions

        If ($PSCmdlet.ParameterSetName -eq "Legacy_Version") {
        
            Write-Debug "Getting topology from store by version..."
            If ($Version -in $Versions.Version) {
                $XmlFiles = Get-CHNLastTopologyXmls -Version $Version
                
            }
            Else {
                Write-Host "Cannot find version [$Version] in store." -ForegroundColor Red
                Break
            }
        }

        If ($PSCmdlet.ParameterSetName -eq "Legacy_Number") {
            Write-Debug "Getting topology from store by version index..."
            If ($Last -in $Versions.Last) {
                $XmlFiles = Get-CHNLastTopologyXmls -Last $Last
            }
            Else {
                Write-Host "Last number [$last] out of range [1 - $($Versions[-1].Last)]." -ForegroundColor Red
                Break
            }
        }

        $Topology = Import-CHNTopology -Types $Types -XmlFiles $XmlFiles

    }
    
    Return $Topology

}

Function Get-CHNTopologyVersions {

<#
.SYNOPSIS
This function is used for list all of versions of the CHN topology which was stored before.

.DESCRIPTION
 If you didn't specify parameter 'Last', it will dispaly such as below:
  Last Version                                              Types
 ---- -------                                              -----
    1 20160908081945                                       {PMs, Zones}
    2 20160908081913                                       {PMs, Zones}
    3 20160908080923                                       {PMs, Zones}
    4 20160831080054                                       {Farms, Networks, NLBs, PMs...}
    5 20160831074330                                       {Farms, Networks, NLBs, PMs...}
    6 20160830071004                                       {Farms, Networks, NLBs, PMs...}
    7 20160830065557                                       {Farms, Networks, NLBs, PMs...}
    8 20160830065319                                       {Farms, Networks, NLBs, PMs...}
    9 20160826015321                                       {Farms, Networks, NLBs, PMs...}

If you want to get topology info from another path, please use parameter 'StorePath'.
Default path is "\\CHN.SPONETWORK.COM\Tools\Topology".

#>
    
    Param (
       [Int] $Last,
       [String]$StorePath="\\CHN.SPONETWORK.COM\Tools\Topology" 
    )

    $XmlFiles = Get-ChildItem -Path $StorePath -file -Recurse
    $FilesName = $XmlFiles.BaseName
    $Versions = $FilesName -split '_' | ? {$_ -Match "\d+"} | Select-Object -Unique | Sort-Object -Descending
    $ObjVer = @()
    $LastNumber = 0
    foreach ( $v in $Versions) {
        $Obj = New-Object -TypeName PsObject
        $LastNumber ++
        $Obj | Add-Member -NotePropertyName Last -NotePropertyValue $LastNumber
        $Obj | Add-Member -NotePropertyName Version -NotePropertyValue $v
        $thisVersionFileNames = $FilesName | ?{ $_ -match $v } 
        $types = $thisVersionFileNames -split '_' | ? {$_ -notmatch "\d+"} 
        $Obj | Add-Member -NotePropertyName Types -NotePropertyValue $types 
        $ObjVer += $Obj
    }
    If ($Last -gt 0) {
        $ObjVer = $ObjVer | ? Last -EQ $Last
    }

    Return $ObjVer
}

Function Get-CHNLastTopologyXmls {
    
    Param (
        [Parameter(ParameterSetName="Number")]
        [ValidateScript({$_ -gt 0})]
        [Int] $Last=1,
        [Parameter(ParameterSetName="Version")]
        [String] $Version,
        [String]$StorePath="\\CHN.SPONETWORK.COM\Tools\Topology"
    )

    Write-Debug "Last topology xmls is finding ..."

    
    $LastXmlFiles = @()
    If ($PSCmdlet.ParameterSetName -eq "Number") {
        $LastVersion = Get-CHNTopologyVersions -StorePath $StorePath -Last $Last | Select-Object -ExpandProperty Version
    }
    Else {
        $LastVersion = $Version
    }

    If ($LastVersion) {
        $LastXmlFiles = Get-ChildItem -Path $StorePath -file -Recurse | ? Name -Match $LastVersion
        If (!$LastXmlFiles) { Write-Host "Cannot find xml file for version [$LastVersion]!"; Break }
    }
    
    Return $LastXmlFiles
}

Function Import-CHNTopology {
    
    Param (
        [ValidateSet("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs","All")]
        [String[]] $Types="All",
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo[]] $XmlFiles
    )

    Write-Debug "Starting import reference topology ..."

    If ($Types -eq "All") {
        $Types = @("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs")
    }

    $Topology =@{}

    Switch ($Types) {

        "VMs" {
            $VMs = $XmlFiles | ? BaseName -Match "^VMs" | Import-Clixml
            If ($VMs) {$Topology.Add($_,$VMs)}
        }

        "PMs" {
            $PMs = $XmlFiles | ? BaseName -Match "^PMs" | Import-Clixml
            If ($PMs) {$Topology.Add($_,$PMs)}
        }

        "Farms" {
            $Farms = $XmlFiles | ? BaseName -Match "^Farms" | Import-Clixml
            If ($Farms) {$Topology.Add($_,$Farms)}
        }

        "Networks" {
            $Networks = $XmlFiles | ? BaseName -Match "^Networks" | Import-Clixml
            If ($Networks) {$Topology.Add($_,$Networks)}
        }

        "Zones" {
            $ZONEs = $XmlFiles | ? BaseName -Match "^ZONEs" | Import-Clixml
            If ($ZONEs) {$Topology.Add($_,$ZONEs)}
        }

        "NLBs" {
            $NLBs = $XmlFiles | ? BaseName -Match "^NLBs" | Import-Clixml
            If ($NLBs) {$Topology.Add($_,$NLBs)}
        }

        "Vlans" {
            $VLANs = $XmlFiles | ? BaseName -Match "^VLANs" | Import-Clixml
            If ($VLANs) {$Topology.Add($_,$VLANs)}
        }

        "VMIPs" {
            $VMIPs = $XmlFiles | ? BaseName -Match "^VMIPs" | Import-Clixml
            If ($VMIPs) {$Topology.Add($_,$VMIPs)}
        }
    }

    Write-Debug "Import done."

    Return $Topology
}

Function Compare-CHNTopology {

<#
.SYNOPSIS
This function is used for compare topology with former.

.DESCRIPTION
Topology information stored at \\chn\tools\topology folder. 
comparison types include "VMs","PMs","Farms","Networks","Zones","NLBs","Vlans" and "VMIPs" as default. 
Just including one or more types is acceptable. Send comparison result by mail is optional.
If you omit parameter 'Last', topology versions will be listed for your choice.

.PARAMETER Last
This parameter is used for specify which version will compare with current topology. Version list will 
display for selection if you omit it. It must great than 0 and less than versions count.

.PARAMETER Types
including "VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs" and "All". All is default.

.PARAMETER SnapShot
Save current topoloy to store folder.

.PARAMETER MailTo
This parameter is optional. Omit this will not send mail.

.EXAMPLE
C:\PS> Compare-CHNTopology.ps1 -MailTo "o365_spo@21vianet.com"
Last Version        Types
---- -------        -----
   1 20160908081945 {PMs, Zones}
   2 20160908081913 {PMs, Zones}
   3 20160908080923 {PMs, Zones}
   4 20160831080054 {Farms, Networks, NLBs, PMs...}
   5 20160831074330 {Farms, Networks, NLBs, PMs...}
   6 20160830071004 {Farms, Networks, NLBs, PMs...}
   7 20160830065557 {Farms, Networks, NLBs, PMs...}
   8 20160830065319 {Farms, Networks, NLBs, PMs...}
   9 20160826015321 {Farms, Networks, NLBs, PMs...}

Please select version ('Last' number):

Select 'Last' number to compare with current topology and send result to spo team.

.EXAMPLE
C:\PS> Compare-CHNTopology.ps1 -Types VMs -Last 4

Just compare VMs with last version 4 without mail.

.EXAMPLE
C:\PS> Compare-CHNTopology.ps1 -Types "VMs","Farms" -Last 5 -MailTo "song.wende@oe.21vianet.com","zhang.peng@oe.21vianet.com"

Just compare VMs and Farms with last version 5 and send result to Wind and Frank.

#>

    Param (
        [ValidateSet("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs","All")]
        [String[]] $Types = "All",
        [Parameter(ParameterSetName="CurrentComparison")]
        [ValidateScript({$_ -gt 0})]
        [Int] $Last,
        [Parameter(ParameterSetName="CurrentComparison")]
        [Alias("Export")]
        [Switch] $SnapShot,
        [Parameter(Mandatory=$true,ParameterSetName="HashComparison",Position=0)]
        [System.Collections.Hashtable] $DifferenceTopology,
        [Parameter(Mandatory=$true,ParameterSetName="HashComparison",Position=1)]
        [System.Collections.Hashtable] $ReferenceTopology,
        [Parameter(Mandatory=$true,ParameterSetName="IndexComparison",Position=0)]
        [ValidateScript({$_ -gt 0})]
        [Int] $DifferenceIndex,
        [Parameter(Mandatory=$true,ParameterSetName="IndexComparison",Position=1)]
        [ValidateScript({$_ -gt $DifferenceIndex})]
        [Int] $ReferenceIndex,
        [Parameter(Mandatory=$true,ParameterSetName="VersionComparison",Position=0)]
        [String] $DifferenceVersion,
        [Parameter(Mandatory=$true,ParameterSetName="VersionComparison",Position=1)]
        [String] $ReferenceVersion,
        [String[]] $MailTo
    )

    $Results = @{}
    $MailBody = @()

    If ($Types -eq "All") {
            $Types = @("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs")
    }

    $date = Get-Date
    $StrDate = Get-Date -Format "MM/dd/yyyy"

    # Start to load topology.

    If ($PSCmdlet.ParameterSetName -eq "CurrentComparison") {

        Write-Host "Starting current topology comparison..."
        
        # List versions for choice.
        If ($Last -eq 0) {
            $Allversion = Get-CHNTopologyVersions | ft -AutoSize
            Out-Host -InputObject $Allversion
            $Last = Read-Host "Please select version ('Last' number)"
            If ($Last -gt $Allversion.Count) {
                Write-Host "Last number [$Last] out of range." -ForegroundColor Red
                Break
            }
        }

        Write-Host "Importing reference topoogy..." 
        $RefVersion = Get-CHNTopologyVersions -Last $Last
        <#
        If ($Types.Count -gt $RefVersion.Types.Count) {
            Write-Host "Reference topology types is less than query types!" -ForegroundColor Red
            Break
        }
        #>
        $RefTypes = $RefVersion.Types
        $RefTopology = Get-CHNTopology -Types $Types -Last $Last

        $IsTypeMatched = InspectTypes -Types $Types -RefTypes $RefTypes -DifTypes $Types
        If (!$IsTypeMatched) { Break }
        
        Write-Host "Getting current topology..."
        If ($SnapShot.IsPresent) { $DifTopology = Get-CHNTopology -Types $Types -Export }
        Else { $DifTopology = Get-CHNTopology -Types $Types }

        
    }

    If ($PSCmdlet.ParameterSetName -eq "HashComparison") {

        Write-Host "Starting topology comparison..."
        $RefTypes = $ReferenceTopology.Keys
        $DifTypes = $DifferenceTopology.Keys

        $IsTypeMatched = InspectTypes -Types $Types -RefTypes $RefTypes -DifTypes $DifTypes
        If (!$IsTypeMatched) { Break }

        # Just compare query types. Remove unnecessary types.
        If ($Types.Count -lt $RefTypes.Count) {
            $RefTopology = @{}
            $DifTopology = @{}
            $Types | %{
                $RefTopology.Add($_,$ReferenceTopology["$_"])
                $DifTopology.Add($_,$DifferenceTopology["$_"])
            }
        }

    }

    If ($PSCmdlet.ParameterSetName -eq "IndexComparison") {

        Write-Host "Starting topology comparison by index..."

        $RefVersion = Get-CHNTopologyVersions -Last $ReferenceIndex
        If (!$RefVersion) {
            Write-Host "Reference index out of range!" -ForegroundColor Red
            Break
        }
        $RefTypes = $RefVersion.Types

        $DifVersion = Get-CHNTopologyVersions -Last $DifferenceIndex
        $DifTypes = $DifVersion.Types

        $IsTypeMatched = InspectTypes -Types $Types -RefTypes $RefTypes -DifTypes $DifTypes
        If (!$IsTypeMatched) { Break }

        Write-Host "Importing reference topology..."
        $RefTopology = Get-CHNTopology -Last $ReferenceIndex
        Write-Host "Importing Difference topology..."
        $DifTopology = Get-CHNTopology -Last $DifferenceIndex
    }

    If ($PSCmdlet.ParameterSetName -eq "VersionComparison") {

        Write-Host "Starting topology comparison by version..."

        $Versions = Get-CHNTopologyVersions
        $RefVersion = $Versions | ? Version -EQ $ReferenceVersion
        If (!$RefVersion) {
            Write-Host "Cannot find reference version [$ReferenceVersion]!" -ForegroundColor Red
            Break
        }
        $RefTypes = $RefVersion.Types

        $DifVersion = $Versions | ? Version -EQ $DifferenceVersion
        If (!$DifVersion) {
            Write-Host "Cannot find difference version [$DifferenceVersion]!" -ForegroundColor Red
            Break
        }
        $DifTypes = $DifVersion.Types

        $IsTypeMatched = InspectTypes -Types $Types -RefTypes $RefTypes -DifTypes $DifTypes
        If (!$IsTypeMatched) { Break }

        Write-Host "Importing reference topology..."
        $RefTopology = Get-CHNTopology -Version $ReferenceVersion
        Write-Host "Importing Difference topology..."
        $DifTopology = Get-CHNTopology -Version $DifferenceVersion

    }

    # End load.

    Write-Host "Starting to compare topology for " -NoNewline
    Write-Host (Convert-ArrayAsString -array $Types) -ForegroundColor Green

    Switch ($Types) {
        "VMs" {
            Write-Host "Comparing $_..."
            $VMItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -CompareField VMachineId
            If ($VMItemResult) {
                $Properties = @("RecordAction","NetworkId","FarmId","Role","Name","State","IPAddress","VMachineId","PMachineId","*Version","VMImageId","DataCenter","FarmRole","FarmState")
                # update 1.2.4
                $VMItemMail = $VMItemResult | Select-Object -Property $Properties | Sort-Object NetworkId,FarmId,Role,RecordAction
                $MailBody += Format-CHNTopologyHtmlTable -Contents $VMItemMail -Title "VMs Changes" -ColorItems
                $Results.Add($_,$VMItemResult)
            }
            
        }

        "PMs" {
            Write-Host "Comparing $_..."
            $PMItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -GroupField Type -CompareField PMachineId
            If ($PMItemResult) { 
                $Properties = @("RecordAction","PMachineId","Name","Type","State","ZoneId","SerialNum","Model","ILoIpAddress","OperatingSystem","PhysicalVlanId","PhysicalZoneId","MacAddress","IpAddress","ImagingRoleName")
                # update 1.2.4
                $PMItemMail = $PMItemResult |Select-Object -Property $Properties | Sort-Object ZoneId,Type,RecordAction
                $MailBody += Format-CHNTopologyHtmlTable -Contents $PMItemMail -Title "PMs Changes" -ColorItems 
                $Results.Add($_,$PMItemResult)
            }
            
        }

        "Networks" {
            Write-Host "Comparing $_..."
            $NetworkItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -GroupField DataCenter -CompareField NetworkId
            If ($NetworkItemResult) {
                $Properties = @("RecordAction","NetworkId","ZoneId","DataCenter","State","Name","Description","SharePath","FailoverState","RecoveryNetworkId")
                $NetworkItemMail = $NetworkItemResult | Select-Object -Property $Properties | Sort-Object RecordAction,NetWorkId
                $MailBody += Format-CHNTopologyHtmlTable -Contents $NetworkItemMail -Title "Network Changes" -ColorItems
                $Results.Add($_,$NetworkItemResult)
            }
            
        }

        "Farms" {
            Write-Host "Comparing $_..."
            $FarmItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -GroupField NetworkId -CompareField FarmId
            If ($FarmItemResult) {
                $Properties = @("RecordAction","FarmId","NetworkId","Role","State","Version","ProductVersion","LastUpdate","ServiceFarmId","FarmImageId","Label","DsType")
                $FarmItemMail = $FarmItemResult | Select-Object -Property $Properties | Sort-Object RecordAction,FarmId
                $MailBody += Format-CHNTopologyHtmlTable -Contents $FarmItemMail -Title "Farm Changes" -ColorItems
                $Results.Add($_,$FarmItemResult)
            }
            
        }

        "VLANs" {
            Write-Host "Comparing $_..."
            $VLANItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -GroupField ZoneId -CompareField VlanId
            If ($VLANItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $VLANItemResult -Title "VLAN Changes" -ColorItems
                $Results.Add($_,$VLANItemResult)
            }
            
        }

        "NLBs" {
            Write-Host "Comparing $_..."
            $NLBItemResult = Compare-CHNItems $RefTopology[$_] $DifTopology[$_] -GroupField Partition -CompareField LoadBalancerId
            If ($NLBItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $NLBItemResult -Title "NLB Changes" -ColorItems
                $Results.Add($_,$NLBItemResult)
            }
            
        }

        "VMIPs" {
            # Compare properties only. Coming soon...
        }
    }

    $TypeStr = Convert-ArrayAsString $Types


    # Logical not operator cannot test hash table. So use IsNullOrEmpty method to test the keys of hash table for judging.
    If ([String]::IsNullOrEmpty($Results.Keys)) {
        Write-Warning -Message "$TypeStr no any changes"
        $MailBody = "<h1 Style=`"color:Green`">No any changes!</h1>"
    }
    Else {
        Write-Host "Comaprison has been completed!" -ForegroundColor Green
    }

    If ($MailTo) {
        Write-Host "Send mail to $MailTo" -ForegroundColor Green
        
        If (!$DifVersion) { $DifVersionStr = "current topology" }
        Else { $DifVersionStr = $DifVersion.Version }

        $RefVersionStr = $RefVersion.Version

        $MailSubject = "Compare $DifVersionStr with $RefVersionStr for type: $TypeStr"

        $From = "TopologyReporter@21vianet.com"
        Send-Email -To $MailTo -mailsubject $MailSubject -mailbody $MailBody -From $From -BodyAsHtml
    }
    Else {
        Write-Host "Didn't send result as mail! If you want, please re-run this script with parameter 'MailTo'." -ForegroundColor Yellow
    }    

    Return $Results

}
# CHNTopologyFunctions>

# <MyUtilities
function Get-MachineInfo {
<#
.SYNOPSIS
Get OS information from Centeral and Windows.

.DESCRIPTION
This function is used for query OS information by WMI. It was used for asset audit but now it's very useful for looking up.

.PARAMETER ComputerName
Which machine do you want to query info. Alia is "CN". Accept array.

.PARAMETER ZoneId
Zone ID which physical machines are belonging to.

.PARAMETER NetworkId
Network ID which virtual machines are belonging to.

.PARAMETER Credential
The credential which you want to use to query.

.PARAMETER CsvPath
It will export result to a CSV file if this parameter is specified.

.EXAMPLE
PS C:\> Get-MachineInfo -CN "BJBCHNVUT082","USR07697-001"
Get OS information for BJBCHNVUT082 and USR07697-001.

.EXAMPLE
PS C:\> Get-MachineInfo -Zone 7 -CsvPath "$home\desktop"
Get all of machines info of zone 7 and store result at desktop.

#>
    param (
        [Parameter(ParameterSetName="Computer",Position=1,Mandatory=$true)]
        [Alias("CN")]
        [String[]] $ComputerName,
        [Parameter(ParameterSetName="Zone")]
        [Int[]] $ZoneId,
        [Parameter(ParameterSetName="Network")]
        [Int[]] $NetworkId,
        [PSCredential] $Credential,
        [String] $CsvPath
    )

    #===========================================================================
    # Nested functions
    function GetOSInfo {
        Param (
            [Object] $InputObject
        )

        if (!$InputObject.IsVM) {
            $HPmp = Get-WmiObject -Namespace root/hpq -Class hp_managementprocessor
            $HPcsc = Get-WmiObject -Namespace root/hpq -Class HP_ComputerSystemChassis
            $InputObject.iLOIpAddress = $HPmp.IPv4Address
            $InputObject.Model = $HPcsc.Model
            $InputObject.ProductID = $HPcsc.ProductID
            $InputObject.SerialNumber = $HPcsc.SerialNumber
        }
        $Win32_OS = Get-WmiObject -Class Win32_OperatingSystem
        $Win32_NAC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration
        $Win32_Processor = Get-WmiObject win32_processor
        $Win32_PhysicalMemory = Get-WmiObject Win32_PhysicalMemory
        $Win32_Logicaldisk = Get-WmiObject Win32_logicalDisk

        $InputObject.OperatingSystem = $Win32_OS.Caption
        $InputObject.OSVersion = $Win32_OS.Version
        $InputObject.ServicePack = $Win32_OS.CSDVersion
        $InputObject.Processor = $Win32_Processor.Name
        $InputObject.ProcessCore = $Win32_Processor.NumberOfCores
        $TotalMemory = 0
        $Win32_PhysicalMemory | %{$TotalMemory += $_.capacity}
        $InputObject.TotalMemory = $TotalMemory
        $TotalDisk = 0
        $Win32_Logicaldisk | %{$TotalDisk += $_.Size}
        $InputObject.TotalDiskSize = $TotalDisk
        $IP = $Win32_NAC.IpAddress
        $IPv4 = @($IP | ?{ $_ -match "\d\." })
        $IPv6 = @($IP | ?{ $_ -match "\d::" })
        $InputObject.NIC_IP = $IPv4
        $InputObject.NIC_IPv6 = $IPv6
        $InputObject.NIC_GateWay = @($Win32_NAC.DefaultIPGateway | ?{$_ -ne $null})
        $InputObject.NIC_MAC = @($Win32_NAC.MACAddress | ?{$_ -ne $null})
        

        return $InputObject
    }
    #===========================================================================

    $Properties = @{
            ComputerName = ""
            IsVM = $false
            Type = ""
            Role = ""
            State = ""
            NetworkId = ""
            ZoneId = ""
            OperatingSystem = ""
            OSVersion = ""
            ServicePack = ""
            Processor = ""
            ProcessCore = ""
            TotalMemory = ""
            TotalDiskSize = ""
            NIC_IP = ""
            NIC_IPv6 = ""
            NIC_GateWay = ""
            NIC_MAC = ""
            iLOIpAddress = ""
            Model = ""
            ProductID = ""
            SerialNumber = ""
    }

    $VMs = @()
    $PMs = @()

    # To filter out VMs and PMs.
    if ($pscmdlet.ParameterSetName -eq "Computer") {
        $PMs += $ComputerName | ?{ $_ -match "CHN" }
        $VMs += $ComputerName | ?{ $_ -notmatch "CHN" }

        foreach ($VM in $VMs) {
            $objVM = New-Object -TypeName PSObject -Property $Properties
            $CenteralVM = Get-CenteralVM -Name $VM
            $objVM.ComputerName = $VM
            $objVM.IsVM = $true
            $objVM.Type = $CenteralVM.Type
            $objVM.State = $CenteralVM.State
            $objVM.NetworkId = $CenteralVM.NetworkId
            if ($Credential) {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objVM -ComputerName $VM -AsJob -Credential $Credential | Out-Null
            }
            else {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objVM -ComputerName $VM -AsJob | Out-Null
            }
            
        }

        foreach ($PM in $PMs) {
            $objPM = New-Object -TypeName PSObject -Property $Properties
            $CenteralPM = Get-CenteralPMachine -Name $PM
            $objPM.ComputerName = $PM
            $objPM.Type = $CenteralPM.Type
            $objPM.State = $CenteralPM.State
            $objPM.Role = $CenteralPM.Role
            $objPM.NetworkId = $CenteralPM.NetworkId
            $objPM.ZoneId = $CenteralPM.ZoneId
            if ($Credential) {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objPM -ComputerName $PM -AsJob -Credential $Credential | Out-Null
            }
            else {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objPM -ComputerName $PM -AsJob | Out-Null
            }
            
        }

    }

    if ($pscmdlet.ParameterSetName -eq "Zone") {
        $CenteralPMs = $ZoneId | %{Get-CenteralPMachine -Zone $_ -ErrorAction Ignore}
        foreach ($CenteralPM in $CenteralPMs) {
            $objPM = New-Object -TypeName PSObject -Property $Properties
            $objPM.ComputerName = $CenteralPM.Name
            $objPM.Type = $CenteralPM.Type
            $objPM.State = $CenteralPM.State
            $objPM.Role = $CenteralPM.Role
            $objPM.NetworkId = $CenteralPM.NetworkId
            $objPM.ZoneId = $CenteralPM.ZoneId
            if ($Credential) {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objPM -ComputerName $CenteralPM.Name -AsJob -Credential $Credential | Out-Null
            }
            else {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objPM -ComputerName $CenteralPM.Name -AsJob | Out-Null
            }
            
        }
    }

    if ($pscmdlet.ParameterSetName -eq "Network") {
        $CenteralVMs += $NetworkId | %{Get-CenteralVM -Network $_ -ErrorAction Ignore}
        foreach ($CenteralVM in $CenteralVMs) {
            $objVM = New-Object -TypeName PSObject -Property $Properties
            $objVM.ComputerName = $CenteralVM.Name
            $objVM.IsVM = $true
            $objVM.Type = $CenteralVM.Type
            $objVM.State = $CenteralVM.State
            $objVM.NetworkId = $CenteralVM.NetworkId
            if ($Credential) {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objVM -ComputerName $CenteralVM.Name -AsJob -Credential $Credential | Out-Null
            }
            else {
                Invoke-Command -ScriptBlock $Function:GetOSInfo -ArgumentList $objVM -ComputerName $CenteralVM.Name -AsJob | Out-Null
            }
            
        }
    }

    $jobs = Get-Job
    if ($jobs) {
        $Result = Watch-Jobs -Activity "Retriving OS infomation" -Status Running
    }
    else {
        Write-Host "Cannot find retriving OS info jobs!" -ForegroundColor Red
        break
    }

    If ($CsvPath) {
        $Result | select ComputerName,IsVM,Type,Role,State,NetworkId,ZoneId,OperatingSystem,OSVersion,ServicePack,`
                        Processor,ProcessCore,TotalMemory,TotalDiskSize,NIC_IP,NIC_IPv6,NIC_GateWay,`
                        NIC_MAC,iLOIpAddress,Model,ProductID ,SerialNumber | Export-Csv $CsvPath
    }

    return $Result
}

Function Get-FreePMachine {
<#
.SYNOPSIS
It used for getting PMs which is not hosting VMs.

.DESCRIPTION

.PARAMETER Zone
Zone ID. If omit, will get all of zones.

.PARAMETER Type
PM type. Default is 'Compute'. Validate set "Compute","Infrastructure","Transact".

.EXAMPLE
PS c:\> Get-FreePMachine
It will find out all of compute PMs which not host VMs.

.EXAMPLE
PS c:\> Get-FreePMachine -Zone 6 -Type Transact
It will find out all of Transact PMs in zone 6 which not host VMs.
#>

    Param (
        [Int] $Zone = 0,
        [ValidateSet("Compute","Infrastructure","Transact")]
        [String] $Type = "Compute"
    )

    $FreePMs = @()
    # Get all of PMahcines of Zone.
    If ($Zone -eq 0) {
        $PMachines = Get-CenteralPmachine -Type $Type
        $Networks = Get-CenteralZone | %{ Get-CenteralNetwork -Zone $_ }
    }
    else {
        $PMachines = Get-CenteralPmachine -Zone $Zone -Type $Type
        $Networks = Get-CenteralNetwork -Zone $Zone
    }

    $PMachines = $PMachines | ? PMachineId -gt 0 | ? State -NE Decommissioned

    #Get all of VMs which hosted in PMachines.
    # For improve performance, we get VMs in batch.
    $VMs = $Networks | %{ Get-CenteralVM -Network $_ } | ? PMachineId -GT 0
    $PMIDinVMs = $VMs | Select-Object -ExpandProperty PMachineId -Unique

    If ($PMIDinVMs)
    {
        ForEach ($PM in $PMachines) {
            If ($PMIDinVMs -notcontains $PM.PMachineId) { $FreePMs += $PM }
        }
    }

    Return $FreePMs
}

Function Get-RebootRequired {
<#
.SYNOPSIS
This script is used for check server pending reboot status.

.DESCRIPTION
If you want to know all of servers pending reboot status, running it without parameter 'ComputerName'. 
It will retrieve all of servers which state in "Running" and "Reserved" from Centeral.

.PARAMETER ComputerName
Machines name which you want to check reboot status.

.PARAMETER Show
Just show statistic info if it present. Otherwise, return result as objects.

.PARAMETER ExportCSV
Export result as "PendingRebootMachines_yyyyMMdd.csv" in current folder if it switch on.

.OUTPUTS
Display result list as formatted table and group table.
Save result as CSV file if 'ExportCSV' present.

.EXAMPLE
PS c:\> Get-RebootRequired -ComputerName "SHA02YL1VUT0187","SHA02YL1VUT0193"
Get pending reboot status for machines "SHA02YL1VUT0187","SHA02YL1VUT0193".

.EXAMPLE
PS c:\> Get-RebootRequired -ComputerName "SHA02YL1VUT0187","SHA02YL1VUT0193" -ExportCSV
Get pending reboot status for machines "SHA02YL1VUT0187","SHA02YL1VUT0193" and export as CSV file at current folder.

#>

    Param ( 
        [Alias("CN")]
        [String[]] $ComputerName,
        [Switch] $Show,
        [Switch] $ExportCSV,
        [PSCredential] $Credential
    )

    $scriptblock = {
        $resultObject = New-Object -TypeName PSObject -Property @{ RebootRequired = $null }
        $a = get-item "hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction silentlycontinue
        If ($a -ne $null) {
            $resultObject.RebootRequired = $true
        }
        else { $resultObject.RebootRequired = $false }

        Return $resultObject
    }

    # Retrieve all of available servers, if not specify $ComputerName.
    If ($ComputerName -eq $null) {
        Write-Host "Please wait, it's retrieving all of server name from Centeral..."
        $VMs = Get-CenteralNetwork | %{ Get-CenteralVM -Network $_ } | ? PMachineid -gt 0 | ? Name -NotMatch "SPD"
        $VMs += Get-SPDFQDN
        $PMs= Get-CenteralPmachine | ? PMachineid -gt 0
        $ServerList += $VMs | ? State -in "Running","Reserved"
        $ServerList += $PMs | ? State -in "Running","Reserved"
        $ComputerName = $ServerList.Name
        if (!$Credential) {
            $Credential = Get-Credential -Message "Enter CHN domain admin credential!"
        }
    }
    if (!$Credential) {
        $job = Invoke-Command -ScriptBlock $scriptblock -ComputerName $ComputerName -JobName RebootQuery -Credential $Credential
    }
    else {
        $job = Invoke-Command -ScriptBlock $scriptblock -ComputerName $ComputerName -JobName RebootQuery
    }

    $results = $job | Watch-Jobs

    if ($show.IsPresent) {
        $results | ? RebootRequired -eq $true | Format-Table -AutoSize RebootRequired, PSComputerName
        $results | Group-Object RebootRequired
    }
    else {
        return $results
    }

    # Export results to CSV file in current folder.
    If ($ExportCSV.IsPresent -and $results) {
        $date = Get-Date -Format "yyyyMMdd"
        $FileName = "PendingRebootMachines_" + $date + ".csv"
        $results | Select-Object RebootRequired, PSComputerName | Export-csv -Path $FileName
    }
}

Function Get-RouteTable {

<#
.SYNOPSIS
This script for getting route table.

.PARAMETER ComputerName
Alias are "CN" and "HostName".

.PARAMETER WithPersistedRoute
Getting route table in addtion to Persisted route table if this parameter present.

.EXAMPLE
PS C:\> Get-RouteTable
Get all of machines route table.

.EXAMPLE
PS C:\> Get-RouteTable -CN "BJBCHNVUT144" -WithPersistedRoute
Get route table with persisted route for BJBCHNVUT144.

#>

    Param (
        [Alias("CN","HostName")]
        [String[]] $ComputerName,
	    [Switch] $WithPersistedRoute
    )

    $ServerList = @()
    $results = @()
    If (!$ComputerName) {
        $ServerList += Get-CenteralPMachine | ? PmachineId -GT 0 | ? State -NE Decommissioned | ? State -NE Dead | Select-Object -ExpandProperty Name
        $serverList += Get-CenteralNetwork | %{Get-CenteralVM -Network $_} | ? PmachineId -GT 0 `
                        | ? State -NE Deleting | ? State -NE Dead | ? State -NE Deleted | Select-Object -ExpandProperty Name
    }
    Else {
        $ServerList = $ComputerName
    }

    $RT_job = Get-WmiObject Win32_IP4RouteTable -ComputerName $ServerList -AsJob
    $results += $RT_job | Watch-Jobs -Activity "Getting IPv4 route table ..."
    

    If ($WithPersistedRoute.IsPresent) {
        $PRT_job = Get-WmiObject Win32_IP4PersistedRouteTable -ComputerName $ServerList -AsJob
        $results += $PRT_job | Watch-Jobs -Activity "Getting IPv4 Persisted route table ..."
    }

    Return $results

}

Function Get-MaintenanceMode {
<#
.SYNOPSIS
This script is used to query maintenance mode in SCOM for specify server or server set.
.DESCRIPTION
If you not specify ComputerName, it will query all of instances's maintenance mode.

.PARAMETER ComputerName
The machine which you want to set in maintenance mode. Array is acceptable. Alias: CN

.EXAMPLE
c:\ps> Get-MaintenanceMode.ps1 -CN "bjbchnvut123","shachnvut123"
Query maintenance mode for those machines.

.EXAMPLE
c:\ps> Get-MaintenanceMode.ps1
Query maintenance mode for all of machines.
#>

    param (
    
        [Alias("CN")]
        [String[]] $ComputerName
    )

    # script block
    $script = {
        Param (
            [String[]] $computers
        )

        try{Import-Module OperationsManager -ea SilentlyContinue}
        catch{Write-Warning("Could not load the OperationsManager module on $env:computername"); break}
        if (!(Get-Module Operations*)) {Write-Warning("Could not load the OperationsManager module on $env:computername"); break}
    
        $Results = @()

        If ($computers) {
            $Instances = Get-SCOMClassInstance -Name $computers | ? FullName -Match "Computer:"
            If (!$Instances) { Return }
            $MaintenanceWindows = Get-SCOMMaintenanceMode -Instance $Instances
        }
        Else {
            $MaintenanceWindows = Get-SCOMMaintenanceMode
        }
    
    
        If ($MaintenanceWindows) {
            If (!$Instances) {
                    $Instances = Get-SCOMClassInstance -Id $MaintenanceWindows.MonitoringObjectId | ? FullName -Match "Computer:"
            }
        
            foreach ($w in $MaintenanceWindows) {
            
                Add-Member -InputObject $w -NotePropertyName MonitoringObject -NotePropertyValue ($Instances | ? Id -EQ $w.MonitoringObjectId).Name
                # convert timezone from UTC to PST
                $PSTTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")
                $CSTTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("China Standard Time")
                Add-Member -InputObject $w -NotePropertyName PSTStartTime -NotePropertyValue ([System.TimeZoneInfo]::ConvertTimeFromUtc($w.StartTime,$PSTTimeZone))
                Add-Member -InputObject $w -NotePropertyName PSTScheduledEndTime -NotePropertyValue ([System.TimeZoneInfo]::ConvertTimeFromUtc($w.ScheduledEndTime,$PSTTimeZone))
                Add-Member -InputObject $w -NotePropertyName CSTStartTime -NotePropertyValue ([System.TimeZoneInfo]::ConvertTimeFromUtc($w.StartTime,$CSTTimeZone))
                Add-Member -InputObject $w -NotePropertyName CSTScheduledEndTime -NotePropertyValue ([System.TimeZoneInfo]::ConvertTimeFromUtc($w.ScheduledEndTime,$CSTTimeZone))
            
            }
        
            $Results = $MaintenanceWindows | ? MonitoringObject -NE $null
        }

        Return $Results
    }

    # Find out the SCOM server which is monitoring the computer.
    $BJBSCOMServer = Get-CenteralVM -Role TMT -Network 9 | Select-Object -First 1 -ExpandProperty Name
    $SHASCOMServer = Get-CenteralVM -Role TMT -Network 11 | Select-Object -First 1 -ExpandProperty Name

    # Convert to FQDN
    $DomainName = "chn.sponetwork.com"
    $FQDNs = @()
    foreach ($c in $ComputerName) {
        $FQDNs += $C,$DomainName -join '.'
    }

    # Pick PMs out.
    $BJBComputers = $FQDNs -match "BJB"
    $SHAComputers = $FQDNs -match "SHA"

    # Pick VMs out.
    $VMs = $FQDNs -notmatch "BJB"
    $VMs = $VMs -notmatch "SHA"
    If ($VMs) {
        foreach ($vm in $VMs) {
            If ((Get-CenteralNetwork -Identity (Get-CenteralVM ($vm -split "\.")[0]).networkid |select -ExpandProperty datacenter) -match "BJB") {
                $BJBComputers += $vm
            }
            Else {
                $SHAComputers += $vm
            }
        }
    }

    # Maintenance mode query
    $timeout = 60
    $results = @()
    $flag = $true
    If ($ComputerName) {
        If ($BJBComputers) {
            Invoke-Command -ComputerName $BJBSCOMServer -ScriptBlock $script -ArgumentList (,$BJBComputers) -JobName BJBQuery | Out-Null
        }

        If ($SHAComputers) {
            Invoke-Command -ComputerName $SHASCOMServer -ScriptBlock $script -ArgumentList (,$SHAComputers) -JobName SHAQuery | Out-Null
        }
    }
    Else {
         Invoke-Command -ComputerName $BJBSCOMServer -ScriptBlock $script -JobName BJBjob | Out-Null
         Invoke-Command -ComputerName $SHASCOMServer -ScriptBlock $script -JobName SHAjob | Out-Null
    }

    $results = Watch-Jobs -Activity "Getting maintenance mode info"

    return $results
}

Function Set-MaintenanceMode {
<#
.SYNOPSIS
This script is used to set or stop maintenance mode in SCOM for specify machine or machine set.
.DESCRIPTION
You can set time window (in minutes) or specify end time for duration of maintenance mode.
Your alias will be set in Comment if you omit it. Default Reasion is UnplannedHardwareMaintenance.
If the machine has been in maintenance mode, this script will reset time window.
It also provide stop maintenance mode with parameter '-StopMaintenanceMode'.

.PARAMETER ComputerName
The machine which you want to set in maintenance mode. Array is acceptable.

.PARAMETER TimeInterval
The time window (in minutes) must be great than 5.

.PARAMETER EndTime
The DateTime format. It must be greate current time than 5 minutes.

.PARAMETER Comment
The comment for maintenance mode. If you omit it, script will write your alias in it.

.PARAMETER Reason
The maintenance options as below (default is 'UnplannedHardwareMaintenance'):
    -- "PlannedOther"
    -- "UnplannedOther"
    -- "PlannedHardwareMaintenance"
    -- "UnplannedHardwareMaintenance"
    -- "PlannedHardwareInstallation"
    -- "UnplannedHardwareInstallation"
    -- "PlannedOperatingSystemReconfiguration"
    -- "UnplannedOperatingSystemReconfiguration"
    -- "PlannedApplicationMaintenance"
    -- "ApplicationInstallation"
    -- "ApplicationUnresponsive"
    -- "ApplicationUnstable"
    -- "SecurityIssue"
    -- "LossOfNetworkConnectivity"

.PARAMETER StopMaintenanceMode
Stop maintenance mode. Don't worry ComputerName include non maintenance machines.

.EXAMPLE
c:\ps> $computers = "bjbchnvut123","shachnvut123"; Set-MaintenanceMode.ps1 -ComputerName $computers -TimeInterval 30
Set maintenance mode 30 minutes from now for "bjbchnvut123","shachnvut123". 
If one of those machines has been in maintenance mode already, it will reset time window 30 minutes from now.

.EXAMPLE
c:\ps> Set-MaintenanceMode.ps1 -ComputerName "bjbchnvut123","shachnvut123" -EndTime "11/16/2016 5:00:00 AM" -Reason "PlannedHardwareInstallation"
Set machines in maintenance mode due time at 11/16/2016 5:00:00 AM and set reason is PlannedHardwareInstallation.

.EXAMPLE
c:\ps> Set-MaintenanceMode.ps1 -CN "bjbchnvut123","shachnvut123" -StopMaintenanceMode
Stop maintenance mode for those machines.
#>

    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias("CN")]
        [String[]] $ComputerName,
        [Parameter(Mandatory=$true,ParameterSetName="Interval",Position=1)]
        [ValidateScript({$_ -ge 5})]
        [Int] $TimeInterval, # Minutes
        [Parameter(Mandatory=$true,ParameterSetName="EndTime")]
        [ValidateScript({$_ -ge (Get-Date).AddMinutes(5)})]
        [DateTime] $EndTime,
        [Parameter(ParameterSetName="Interval")]
        [Parameter(ParameterSetName="EndTime")]
        [String] $Comment,
        [ValidateSet("PlannedOther","UnplannedOther","PlannedHardwareMaintenance",
                    "UnplannedHardwareMaintenance","PlannedHardwareInstallation","UnplannedHardwareInstallation",
                    "PlannedOperatingSystemReconfiguration","UnplannedOperatingSystemReconfiguration",
                    "PlannedApplicationMaintenance","ApplicationInstallation",
                    "ApplicationUnresponsive","ApplicationUnstable",
                    "SecurityIssue","LossOfNetworkConnectivity")]
        [Parameter(ParameterSetName="Interval")]
        [Parameter(ParameterSetName="EndTime")]
        [String] $Reason = "UnplannedHardwareMaintenance",
        [Parameter(Mandatory=$true,ParameterSetName="Stop")]
        [Switch] $StopMaintenanceMode
    )

    # This is a workaround. End time is not necessary in stop mode. But the remote script can not accept null value of DateTime.
    If ($PSCmdlet.ParameterSetName -eq "Stop") {
        $EndTime = (Get-Date).AddMinutes(10)
    }

    If ($PSCmdlet.ParameterSetName -eq "Interval") {
        $EndTime = (Get-Date).AddMinutes($TimeInterval)
    }

    # This is for avoid time kind not set as "Local" if parameter EndTime input as a string.
    If ($PSCmdlet.ParameterSetName -eq "EndTime") {
        $EndTime = [System.TimeZoneInfo]::ConvertTime($EndTime,[System.TimeZoneInfo]::Local)
    }

    If (!$Comment) {
        $Comment = ("Maintenance mode set by {0}" -f $env:USERNAME)
    }

    $script = {
         Param (
            [String[]] $computers,
            [DateTime] $EndTime,
            [String] $Comment,
            [String] $Reason,
            [Bool] $Stop
         )

        try{Import-Module OperationsManager -ea SilentlyContinue}
        catch{Write-Warning("Could not load the OperationsManager module on $env:computername"); break}
        if (!(Get-Module Operations*)) {Write-Warning("Could not load the OperationsManager module on $env:computername"); break}
    
        $Instances = Get-SCOMClassInstance $computers
    
        # Filter out which machine was in maintenance mode already.

        $MaintainedIns = $Instances | ? InMaintenanceMode -EQ $true
        $NotMaintainedIns = $Instances | ? InMaintenanceMode -EQ $false
    
        If ($Stop) {
            foreach ($ins in $MaintainedIns) {
                $Time = [System.TimeZoneInfo]::ConvertTimeToUtc((Get-Date).AddSeconds(3))
                $ins.StopMaintenanceMode($Time)
            }
            Start-Sleep 3

            $ReturnInstances = Get-SCOMClassInstance $computers | ? FullName -Match "Computer:"
            Return $ReturnInstances
        }
        Else {
            # filter out the Computer object only. Because just need to set this object in maintenance mode.
            $MaintainedIns = $MaintainedIns | ? FullName -Match "Computer:"
            $NotMaintainedIns = $NotMaintainedIns | ? FullName -Match "Computer:"
        }

        If ($MaintainedIns) {

            # Use set SCOM maintenace mode command
            $MaintenanceModeEntries = Get-SCOMMaintenanceMode -Instance $MaintainedIns
            Set-SCOMMaintenanceMode -MaintenanceModeEntry $MaintenanceModeEntries -EndTime $EndTime -Comment $Comment -Reason $Reason
        
        }

        If ($NotMaintainedIns) {
            # Use Start SCOM maintenance mode command
            Start-SCOMMaintenanceMode -Instance $NotMaintainedIns -EndTime $EndTime -Comment $Comment -Reason $Reason
        }
        Start-Sleep 1

        $ReturnInstances = Get-SCOMClassInstance $computers | ? FullName -Match "Computer:"
        Return $ReturnInstances
    
    }

    # Find out the SCOM server which is monitoring the computer.
    $BJBSCOMServer = Get-CenteralVM -Role TMT -Network 9 | Select-Object -First 1 -ExpandProperty Name
    $SHASCOMServer = Get-CenteralVM -Role TMT -Network 11 | Select-Object -First 1 -ExpandProperty Name

    # Convert to FQDN
    $DomainName = "chn.sponetwork.com"
    $FQDNs = @()
    foreach ($c in $ComputerName) {
        $FQDNs += $C,$DomainName -join '.'
    }

    # Pick PMs out.
    $BJBComputers = $FQDNs -match "BJB"
    $SHAComputers = $FQDNs -match "SHA"

    # Pick VMs out.
    $VMs = $FQDNs -notmatch "BJB"
    $VMs = $VMs -notmatch "SHA"
    If ($VMs) {
        foreach ($vm in $VMs) {
            If ((Get-CenteralNetwork -Identity (Get-CenteralVM ($vm -split "\.")[0]).networkid |select -ExpandProperty datacenter) -match "BJB") {
                $BJBComputers += $vm
            }
            Else {
                $SHAComputers += $vm
            }
        }
    }

    $timeout = 60
    $results = @()

    If ($BJBComputers) {
        Invoke-Command -ComputerName $BJBSCOMServer -ScriptBlock $script -ArgumentList $BJBComputers,$EndTime,$Comment,$Reason,$StopMaintenanceMode.ToBool() -JobName BJBExec | Out-Null
    
    }

    If ($SHAComputers) {
        Invoke-Command -ComputerName $SHASCOMServer -ScriptBlock $script -ArgumentList $SHAComputers,$EndTime,$Comment,$Reason,$StopMaintenanceMode.ToBool() -JobName SHAExec | Out-Null
    
    }


    $results = Watch-Jobs -Activity "Setting maintenance mode info" 

    $results | ft -a name,Id,InMaintenanceMode

}

Function Test-WindowsUpdates {
<#
.SYNOPSIS
This function is used to check machines security updates.

.DESCRIPTION
To test windows updates are ready for installation for computers.

.PARAMETER ComputerName
The target machines name. Default is local.

.PARAMETER Credential
The special credential for query windows updates.

.EXAMPLE
PS C:\> Test-WindowsUpdates -CN BJBCHNVUT144
Check windows updates for BJBCHNVUT144.
#>

    Param (
        [Alias("CN")]
        [Parameter(Position=0)]
        [String[]] $ComputerName,
        [PSCredential] $Credential
    )

    $script = {
        $UpdateSession = New-Object -com "Microsoft.Update.Session"
        $Criteria="IsInstalled=0 and Type='Software'"
        $Updates = $UpdateSession.CreateUpdateSearcher().Search($Criteria).Updates
        Return $Updates
    }
    
    If (!$ComputerName) {
        # invoke check for local machine.
        if (!$Credential) {
            Start-Job -ScriptBlock $script | Out-Null
        }
        else {
            Start-Job -ScriptBlock $script -Credential $Credential | Out-Null
        }
    }
    Else {
        if (!$Credential) {
            Invoke-Command -ScriptBlock $script -ComputerName $ComputerName -AsJob | Out-Null
        }
        else {
            Invoke-Command -ScriptBlock $script -ComputerName $ComputerName -AsJob -Credential $Credential | Out-Null
        }
    }

    $Result = Watch-Jobs -Activity "Testing Windows updates..."

    Return $Result

}

Function Get-FolderSize {
<#
.SYNOPSIS
To see folder size.

.DESCRIPTION
If you want to query root drive (c:\ or d:\ etc.), please use Get-DiskUsage for more efficient.
Parameter 'SubFolder' will query one level sub folder size. Query root drive's sub folder is not very sensible.
Note: one level folder size sum may not equal folder size if folder include files.

.PARAMETER Path
Query folder path.

.PARAMETER ComputerName
Remote computers which you want to query.

.PARAMETER SubFolder
On level sub foder of path.

.INPUTS
None.

.OUTPUTS
Return PS custom object. Properites include 'ComputerName', 'FullPath' and 'Size'.

.EXAMPLE
PS C:\> Get-FolderSize c:\output -ComputerName BJBCHNVUT128
Query folder size of 'C:\output' for BJBCHNVUT128.

.EXAMPLE
PS C:\> Get-FolderSize -Path "C:\CircuitBreakerV3" -ComputerName (Get-CenteralFrontEnd).Name -SubFolder
Query one level sub folder size of "C:\CircuitBreakerV3" for all GFEs.

#>

    Param(
        [Parameter(Mandatory=$true,Position=0)]
        [String] $Path,
        [Alias("CN")]
        [String[]] $ComputerName,
        [Switch] $SubFolder
    )

    $RemoteComputerName = @()

    # Abstract local host in case of ComputerName include it.
    # No specify computer name same as localhost.
    $IsIncludeLocalHost = $false

    If ($ComputerName) {
        $LocalHost = @(".","localhost",$env:COMPUTERNAME)
        
        Foreach ($c in $ComputerName) {
            # trim domain name.
            If ($c -match $env:USERDNSDOMAIN) { 
                $c = ($c -split "\.")[0] 
            }

            If ($c -in $LocalHost) { $IsIncludeLocalHost = $true }
            Else { $RemoteComputerName += $c }
        }
    }
    Else {
        $IsIncludeLocalHost = $true
    }

    # Remote code.
    $ScriptBlock = {
        Param ([String] $Path, [Bool] $SubFolder)

        If (!(Test-Path $Path)) {
            Write-Host "Path '$Path' on $($env:COMPUTERNAME) is not exist!" -ForegroundColor Red
            Return
        }

        $Result = @()

        If ($SubFolder) { $QueryPath = Get-ChildItem -Path $Path -Directory -Force }
        Else { 
            $QueryPath = Get-Item -Path $Path

            # Identify the path is not root of dirve.
            If ($QueryPath.Root.FullName -eq $QueryPath.FullName) {
                Write-Host "Please use Get-DiskUsage to query drive size." -ForegroundColor Yellow
                Return
            }
         }

        Foreach ($QP in $QueryPath) {
            $Files = Get-ChildItem -Path $QP.FullName -File -Recurse -Force
            $Size = 0
            $Files | ForEach-Object { $Size += $_.Length}
            $Properties = @{
                ComputerName = $env:COMPUTERNAME
                FullPath = $QP.FullName
                LastWriteTime = $QP.LastWriteTime
                Size = $Size
            }
            $Obj = New-Object -TypeName PSObject -Property $Properties
            $Result += $Obj
        }

        Return $Result
    }

    If ($IsIncludeLocalHost) {
        Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Path,($SubFolder.IsPresent) | Out-Null
    }

    If ($RemoteComputerName) {
        Invoke-Command -ScriptBlock $ScriptBlock -ComputerName $RemoteComputerName -ArgumentList $Path,($SubFolder.IsPresent) -AsJob | Out-Null
    }

    $Result = Watch-Jobs -Activity "Fetching Folder size..."

    Return $Result

}

Function Get-SyncEndPoint {
<#
.SYNOPSIS
This script for getting sync endpoint for SPO.

.DESCRIPTION
Sometimes, we need to know which BOT is hosting tenant sync timer job.
#>

    $proxy = New-WebServiceProxy -URI ("http://" + $env:GMSvr+ "/tenantmgr.asmx") -UseDefaultCredential
    $ContentFarm = Get-Centeralfarm -Role Content -State Active
    $endPoint = $contentFarm | ForEach-Object {$proxy.GetTenantAPIEndPoint($_.farmid)}
    $TestProxy = $endpoint | ForEach-Object { New-WebServiceProxy -URI $_ -UseDefaultCredential}
    return $TestProxy.GetRunningOnboardingTimerJob()
    
}

Function Search-DomainUser {
<#
.SYNOPSIS
Search CHN domain user.

.DESCRIPTION
To get CHN domain user by SAMAccountName or fuzzy search in SAMAccountName, name or display name.

.PARAMETER Name
The name which appeared in SAMAccountName, name or display name.

.OUTPUTS
ADUser objects.

#>
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String] $Name
    )

    # Get SAM account name first.
    $filter = "samaccountname -eq `"{0}`"" -f $Name
    $ADuser = Get-ADUser -Filter $filter -Properties *

    # Search name in name or display name if didn't find in SAM account name.
    If (!$ADuser) {
        $filter = "samaccountname -like `"*{0}*`" -or name -like `"*{0}*`" -or displayname -like `"*{0}*`"" -f $Name
        $ADuser = Get-ADUser -Filter $filter -Properties *
    }

    Return $ADuser
}

Function New-RandomPassword {

<#
.SYNOPSIS
Generate a new random password which comply MS password policy.

.DESCRIPTION
Password start with prefix '_Sp0' and follow with a new guid.

#>


    $PlainText = "_Sp0" + $([GUID]::NewGuid()).ToString("N")
    Return $PlainText
}

Function Get-MachineUpTime {
<#
.SYNOPSIS
Get machine boot up time.

.DESCRIPTION
Batch query is valid in this funtion.

.PARAMETER ComputerName
Machine names. localhost is default.

.PARAMETER Credential
CHN account credential.

.EXAMPLE
PS C:\> Get-MachineUpTime
Get local machine up time.

.EXAMPLE
PS C:\> Get-MachineUpTime -CN BJBCHNVUT144
Get remote machine up time.
#>

    Param (
        [Alias("CN", "HostName")]
        [String[]] $ComputerName,
        [System.Management.Automation.PSCredential] $Credential
    )

    # Nested functions
    Function AbstractLastUpTime {
        Param ($WmiOS)

        # Convert argument as array force
        $WmiOS = $WmiOS -as [System.Array]
        
        $Result = @()

        foreach ($os in $WmiOS) {
            $BootUpTime = $os.ConvertToDateTime($os.LastBootUpTime)
            $CurrentTime = $os.ConvertToDateTime($os.LocalDateTime)
            $p = @{
                ComputerName = $os.PSComputerName
                BootUpTime = $BootUpTime
                CurrentTime = $CurrentTime
                UpTime = $CurrentTime - $BootUpTime
            }

            $Result += New-Object -TypeName PSObject -Property $p
        }

        return $Result
    }

    # determine the target machine is remote or local.
    if (!$ComputerName) { $ComputerName = @("localhost") }

    $command = "Get-WmiObject Win32_OperatingSystem -ComputerName `$ComputerName"

    if ($Credential) {
        $command = $command, "-Credential `$Credential" -join " "
    }

    # Starting job mode if computer count gt 5 for performance improvement.
    if ($ComputerName.Count -lt 5) {
        $WmiOS = Invoke-Expression $command
    }
    else {
        Invoke-Expression ($command, "-AsJob" -join " ") | Out-Null
        $WmiOS = Watch-Jobs -Activity "Up time querying"
    }
    if ($WmiOS) {
        $Result = AbstractLastUpTime $WmiOS
    }

    return $Result
}

function Test-MachineWarmUp {
<#
.SYNOPSIS
Test machine is whether starting up or started and return reboot time after up.

.Description
This function will monitoring computer in 15 minutes by default.
In the duration, the monitoring status will be changed if the computer restarting.
The status is one of "Running, Restarting, Warming, WarmUp and Down".
It will return a PS object which including fileds "ComputerName", "Status" and "DateTime" to record
the computer restarting.
If you find the last record's status is "Down", you should check the computer manually to determine
there is a real problem or not.
If during the monitoring, no any restarting action, it just return "Running" status record.
If the restarting was performed before run this function, the first record must be "Restarting".

.Parameter ComputerName
The target computer name.

.Parameter Duration
The monitoring duration which range is from 15 to 60 minutes. Default is 15 minutes.

.Parameter HideProgressBar
Show progress bar if omit this switch.

.Outputs
The PS object which fields are "ComputerName", "Status" and "DateTime".
The status including "Running, Restarting, Warming, WarmUp and Down".

.Example
Test-MachineWarmUp -CN bjbchnvut084 -Verbose
VERBOSE: ComputerName: bjbchnvut084
VERBOSE: Starting to monitor...DateTime:[8/16/2017 1:57:06 AM]
VERBOSE: Computer is running
VERBOSE: Computer is restarting from 08/16/2017 01:57:43
VERBOSE: Computer is warming from 08/16/2017 02:01:29
VERBOSE: Warm up!
VERBOSE: Monitoring end up...DateTime:[8/16/2017 2:02:30 AM]

Status                                         ComputerName                                   DateTime
------                                         ------------                                   --------
Running                                        bjbchnvut084                                   8/16/2017 1:57:06 AM
Restarting                                     bjbchnvut084                                   8/16/2017 1:57:43 AM
Warming                                        bjbchnvut084                                   8/16/2017 2:01:29 AM
WarmUp                                         bjbchnvut084                                   8/16/2017 2:02:30 AM

Monitor bjbchnvut084 restarting.

#>

    Param (
        [Parameter(Mandatory=$true)][Alias("CN")]
        [String] $ComputerName,
        [ValidateRange(15,60)]
        [Int] $Duration = 15, #minutes
        [Switch] $HideProgressBar
    )

    
    $TestInterval = 1  # seconds
    $WaitDuration = 10  # minutes
    $WarmUpDuration = 1 # minutes

    $StartTime = Get-Date
    $EndTime = $StartTime.AddMinutes($Duration)

    $TimeRecord = @()
    
    Write-Verbose "ComputerName: $ComputerName"
    Write-Verbose ("Starting to monitor...DateTime:[{0}]" -f $StartTime)

    # Begin monitoring
    do {
        $ping = Test-Connection -ComputerName $ComputerName -Quiet -Count 1
        Start-Sleep $TestInterval
        $currentTime = Get-Date
        $Remaining = $EndTime - $currentTime
        $PercentComplete = ($Duration - $Remaining.TotalMinutes) / $Duration * 100
        if ($PercentComplete -gt 100) { $PercentComplete = 100 }
        
        if (!$HideProgressBar.IsPresent) {
            Write-Progress -Activity "Monitoring Machine $ComputerName" -Status "Computer status: $Status" `
                        -PercentComplete $PercentComplete -SecondsRemaining ($Remaining.TotalSeconds)
        }

        switch ($Status) {
            "Running" {
                if ($ping) { break }
                else { 
                    $Status = "Restarting"
                    Write-Verbose "Computer is restarting from $currentTime"
                    $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                    break 
                }
            }

            "Restarting" {
                if ($ping) { 
                    $Status = "Warming"
                    Write-Verbose "Computer is warming from $currentTime"
                    $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                    break
                }
                else {
                    $RestartingTime = $TimeRecord | ? Status -EQ "Restarting" | select -ExpandProperty DateTime
                    if ( ($currentTime - $RestartingTime).TotalMinutes -gt $WaitDuration) {
                        Write-Verbose "Restarting time exceed $WaitDuration minutes! Please troubleshooting!"
                        $Status = "Down"
                        $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                    }
                    else {
                        break 
                    }
                }
            }

            "Warming" {
                if ($ping) {
                    $WarmingTime = $TimeRecord | ? Status -EQ "Warming" | select -ExpandProperty DateTime
                    # To see the warm up time is exceeded.
                    if ( ($currentTime - $WarmingTime).TotalMinutes -gt $WarmUpDuration) {
                        Write-Verbose "Warm up!"
                        $Status = "WarmUp" # set status is "done" that just for end up.
                        $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                    }
                    else { break }
                }
                else {
                    # It should be encounter problem if machine cannot be ping during warm up.
                    $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                    Write-Verbose "Cannot ping $ComputerName during warm up. Please troubleshooting!"
                    $Status = "Down"
                    $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$currentTime;Status=$Status;ComputerName=$ComputerName}
                }
            }

            {"Down","WarmUp" -contains $Status} {
                # Should end up
                $currentTime = $EndTime.AddMinutes(1)
                Write-Verbose ("Monitoring end up...DateTime:[{0}]" -f (Get-Date))
            }

            default {
                # Should first time test
                if ($ping) {
                    $Status = "Running"
                    Write-Verbose "Computer is running"
                }
                else {
                    $Status = "Restarting"
                    Write-Verbose "Computer is restarting"
                }
                $TimeRecord += New-Object -TypeName PSObject -Property @{DateTime=$StartTime;Status=$Status;ComputerName=$ComputerName}
            }
        }
        
    } while ($currentTime -le $EndTime)

    return $TimeRecord 

}

function Watch-CHNCenteralJob {
    
    Param (
        [Parameter(Mandatory=$true,
            ParameterSetName="ID",
            ValueFromPipelineByPropertyName=$true,
            Position=1)]
        [Int] $JobId,
        [Parameter(ParameterSetName="Type",Position=1)]
        [String] $Type,
        [Parameter(ParameterSetName="Type")]
        [Int] $ObjectId,
        [Int[]] $AlertStep,
        [ValidateScript({[System.Array]::sort($AlertStep);$_ -gt $AlertStep[-1]})]
        [Int] $EndStep,
        [ValidateRange(1,24)]
        [Int] $Duration=8, # hours
        [ValidateScript({$_ -le ($Duration * 60 * 60 / 4)})]
        [Int] $Interval=30, # seconds
        [Switch] $WithNotificationMail

    )

    begin {
        
        $CenteralJob = @()
        $CurrentCenteralJob = @()
        $AlertJobHistory = @{}
        $StartTime = Get-Date
        $EndTime = $StartTime.AddHours($Duration)
        $EndUp = $false
        $ShowProperties = "JobId", "Type", "Step", "State", "Owner", "StartTime", "NextTime", "RetryCount", "ObjectId"
        $Separator = "=" * ($Host.UI.RawUI.WindowSize.Width / 2)
        if ($WithNotificationMail.IsPresent) {
            $Sender = "WatchCenteralJob@21vianet.com"
            $To = Get-ADUser $env:USERNAME -Properties EmailAddress | select -ExpandProperty EmailAddress
        }

        function NextRefreshTime {
            Param (
                [Parameter(Mandatory=$true)]
                [Object[]] $CenteralJob,
                [Parameter(Mandatory=$true)]
                [Int] $Interval
            )

            # return interval time as next time
            $RunningCenteralJob = $CenteralJob | ? Owner -NE ""
            if ($RunningCenteralJob) {
                # return interval time as next time
                return (Get-Date).AddSeconds($Interval)
            }
            else {
                # return the nearest next time plus interval
                $MinNextTime = $CenteralJob[0].NextTime
                if ($CenteralJob.Count -gt 1) {
                    foreach ($time in $CenteralJob.NextTime) {
                         if ($MinNextTime.CompareTo($time) -gt 0) {
                            $MinNextTime = $time
                         }
                    }
                }
                return (Get-Date -Date $MinNextTime).AddSeconds($Interval)
            }
        }

        function FilterCurrentAlertJob {
            param (
                [System.Collections.Hashtable] $HistoryJobTable,
                [Object[]] $AlertJob
            )
            $currentAlertJob = @()
            if ($HistoryJobTable.Values -ne $null) {
                foreach ($job in $AlertJob) {
                    $HistoryStep = $HistoryJobTable[$($job.Jobid)]
                    if ($HistoryStep) {
                        if ($HistoryStep -lt $job.Step) {
                            $HistoryJobTable[$($job.Jobid)] = $job.Step
                            $currentAlertJob += $job
                        }
                    }
                    else {
                        $HistoryJobTable.Add($job.JobId, $job.Step)
                        $currentAlertJob += $job
                    }
                }
            }
            else {
                # Create hsitory job table
                foreach ($job in $AlertJob) {
                    $HistoryJobTable.Add($job.JobId, $job.Step)
                    $currentAlertJob += $job
                }
            }
        
            return $currentAlertJob, $HistoryJobTable
        }

    }

    process {
        if ($PSCmdlet.ParameterSetName -eq "ID") {
            $CenteralJob += Get-CenteralJob $JobId
        }

        if ($PSCmdlet.ParameterSetName -eq "Type") {
            if ($ObjectId) {
                $CenteralJob = Get-CenteralJob -Type $Type -ObjectId $ObjectId
            }
            else {
                $CenteralJob = Get-CenteralJob -Type $Type
            }
        }
    }

    end {

        if (!$CenteralJob) {
            Write-Host "Cannot find Centeral job!" -ForegroundColor Red
            break
        }

        Write-Host "Monitor will clean your console! " -ForegroundColor Yellow -NoNewline
        $answer = Read-Host -Prompt "Do you want to proceed?(Yes/No)"
        if ($answer -match "n[o]*") {
            break
        }

        

        do {
            Clear-Host
            Write-Host "`n`n`n`n`n`n`n`n`n"
            Write-Host "Monitoring Centeral Job" -ForegroundColor Green
            Write-Host "Start Time: $StartTime"
            Write-Host "Expired Time: $EndTime"
            Write-Host "Monitor Duratin: $Duration hours"
            Write-Host "Alert Step: $(Convert-ArrayAsString $AlertStep -Quote None)"
            Write-Host "End Step: $EndStep"
            Write-Host "Default refresh interval: $Interval seconds"
            Write-Host "Send notification mail: $($WithNotificationMail.IsPresent)"
            Write-Host ($Separator * 2)

            $CurrentCenteralJob = $CenteralJob | Get-CenteralJob -ErrorAction SilentlyContinue
            if ($CurrentCenteralJob) {

                $CurrentCenteralJob | ft -AutoSize $ShowProperties | Out-Host

                if ($AlertStep) {
                    Write-Host $Separator -ForegroundColor Yellow
                    Write-Host "Hits alert step [$(Convert-ArrayAsString $AlertStep -Quote None)] jobs:" -ForegroundColor Yellow
                    <#
                    $AlertJob = $CurrentCenteralJob | ? Step -EQ $AlertStep | ?{$_.JobId -notin $AlertJobHistory.JobId}
                    if ($AlertJob) {
                        # to be continue
                        Write-Host "$(Convert-ArrayAsString $AlertJob.JobId)" -ForegroundColor Yellow
                        
                        if ($WithNotificationMail.IsPresent) {
                            $Content = Format-HtmlTable -Contents ($AlertJob | select -Property $ShowProperties) `
                                        -Title "Alert Step Jobs:" -Cellpadding 1 -Cellspacing 1
                            $Title = "Job hits alert step [$AlertStep]"
                            Send-Email -To $To -From $Sender -mailbody $Content -BodyAsHtml -mailsubject $Title
                        }

                        $AlertJobHistory += $AlertJob
                    }
                    else {
                        Write-Host "No alert job!" -ForegroundColor Green
                    }
                    #>
                    $AlertJob = $CurrentCenteralJob | ? {$AlertStep -contains $_.Step}
                    $currentAlertJob, $AlertJobHistory = FilterCurrentAlertJob -HistoryJobTable $AlertJobHistory -AlertJob $AlertJob

                    if ($currentAlertJob) {
                        Write-Host "$(Convert-ArrayAsString $currentAlertJob.JobId)" -ForegroundColor Yellow
                        
                        if ($WithNotificationMail.IsPresent) {
                            $Content = Format-HtmlTable -Contents ($currentAlertJob | select -Property $ShowProperties) `
                                        -Title "Alert Step Jobs:" -Cellpadding 1 -Cellspacing 1
                            $currentAlertStep = $currentAlertJob.Step
                            $Title = "Job hits alert step [$(Convert-ArrayAsString $currentAlertStep -Quote None)]"
                            Send-Email -To $To -From $Sender -mailbody $Content -BodyAsHtml -mailsubject $Title
                        }

                    }
                    else {
                        Write-Host "No alert job!" -ForegroundColor Green
                    }
                    Write-Host $Separator -ForegroundColor Yellow
                }

                if ($EndStep) {
                    $EndCenteralJob = $CurrentCenteralJob | ? Step -EQ $EndStep
                    if ($EndCenteralJob) {
                        if ($EndCenteralJob.count -eq 1) {
                            Write-Host "Job [$($EndCenteralJob.JobId)] hits end step!" -ForegroundColor Green
                            $EndUp = $true
                        }
                        elseif ($EndCenteralJob.Count -gt 1) {
                            Write-Host "Jobs [$(Convert-ArrayAsString $EndCenteralJob.JobId)] hit end step!" -ForegroundColor Green
                            $EndUp = $true
                        }

                        if ($WithNotificationMail.IsPresent) {
                            $Content = Format-HtmlTable -Contents ($EndCenteralJob | select -Property $ShowProperties) `
                                        -Title "End Step Jobs:" -Cellpadding 1 -Cellspacing 1
                            $Title = "Job hits end step [$EndStep]"
                            Send-Email -To $To -From $Sender -mailbody $Content -BodyAsHtml -mailsubject $Title
                        }
                        
                    }
                }

                $SuspendedJob = @($CurrentCenteralJob | ? State -EQ "Suspended")
                if ($SuspendedJob) {
                    Write-Host "There are $($SuspendedJob.Count) jobs suspended!" -ForegroundColor Yellow
                    if ($SuspendedJob.Count -eq $CurrentCenteralJob.Count) {
                        Write-Host "All jobs $(Convert-ArrayAsString $CurrentCenteralJob.JobId) suspended!" -ForegroundColor Red
                        if ($WithNotificationMail.IsPresent) {
                            $Content = Format-HtmlTable -Contents ($SuspendedJob | select -Property $ShowProperties) `
                                        -Title "Suspended Jobs:" -Cellpadding 1 -Cellspacing 1
                            $Title = "All job suspended!"
                            Send-Email -To $To -From $Sender -mailbody $Content -BodyAsHtml -mailsubject $Title
                        }
                        break
                        Write-Host ($Separator * 2)
                    }

                }

                # To determine next refresh time
                $NextRefreshTime = NextRefreshTime -CenteralJob $CurrentCenteralJob -Interval $Interval
                Start-SleepProgress -Seconds ($NextRefreshTime - (Get-Date)).TotalSeconds -Activity "Next refresh time: $NextRefreshTime"

            }
            else {
                Write-Host "All jobs $(Convert-ArrayAsString $CenteralJob.JobId) finished!" -ForegroundColor Green
                if ($WithNotificationMail.IsPresent) {
                    $Content = "All jobs $(Convert-ArrayAsString $CenteralJob.JobId) finished!"
                    $Title = "All jobs finished!"
                    Send-Email -To $To -From $Sender -mailbody $Content -mailsubject $Title
                }
                $EndUp = $true
            }

            if ((Get-Date) -gt $EndTime) {
                Write-Host "Monitor duration $Duration hours is expired at $EndTime !" -ForegroundColor Yellow
                if ($WithNotificationMail.IsPresent) {
                    $Content = "Monitor duration $Duration hours is expired at $EndTime !"
                    $Title = "Monitor duration expired!"
                    Send-Email -To $To -From $Sender -mailbody $Content -mailsubject $Title
                }
                $EndUp = $true
            }

        } while (!$EndUp)
        Write-Host ($Separator * 2)
    }


}

Function Invoke-ZhouBapi {
<#
.SYNOPSIS
This function is using for Centeral job monitoring.

.DESCRIPTION
Invoke this function to monitring Centeral jobs by JobId or Type.
Object ID is optional for job type or standalone.
This version default settings is not set error alert but it with finish and suspend alert.
Default refrash interval is 30 seconds and invoke duration is 8 hours.

.PARAMETER Type
Centeral job types.

.PARAMETER ObjectId
Centeral job object id.

.PARAMETER JobId
Centeral job Id.

.PARAMETER Wait
It will wait expect Centeral job appear in duration time if it's turn on.

.PARAMETER RefreashInterval
How long the default Centeral job check interval. Default is 30 seconds.

.PARAMETER Duration
This function running duration. Default is 8 hours.

.PARAMETER AlertSetp
The Centeral job steps which you want to get mail notification.

.PARAMETER FinishAlert
It's a switch to send mail notification mail when job finished if it's on.
Default is on.

.PARAMETER SuspendedAlert
It's a switch to send notification mail when job suspended if it's on.
Default is on.

.PARAMETER ErrorAlert
It's a switch to send no notification mail when job get error if it's on.
Default is off.

.PARAMETER MailTo
The notification mail recipients.
It will send mail to you if this option ommit.

.EXAMPLE
Invoke-ZhouBapi -Type "CenteralFarmIPUOrch","ActivateFarm" -Wait
Monitoring "CenteralFarmIPUOrch" and "ActivateFarm" jobs. It will wait 8 hours for job appearance if
the jobs are not running.

.EXAMPLE
$JobId | Invoke-ZhouBapi
Monitoring jobs which are stored in $JobId.

#>
	Param (
		[Parameter(Mandatory=$true,Position=0,
			ParameterSetName="type")]
		[String[]] $Type,
        [Parameter(Mandatory=$true,
			ParameterSetName="object")]
        [Parameter(ParameterSetName="type")]
        [Int] $ObjectId,
		[Parameter(ParameterSetName="type")]
        [Parameter(ParameterSetName="object")]
		[Switch] $Wait,
		[Parameter(Mandatory=$true,Position=0,
			ValueFromPipelineByPropertyName=$true,ParameterSetName="id")]
		[Int[]] $JobId,
		[Int] $RefreashInterval = 30, # seconds
		[Int] $Duration = 8, # hours
		[Int[]] $AlertStep,
		[Switch] $FinishAlert = $true,
		[Switch] $SuspendAlert = $true,
		[Switch] $ErrorAlert,
		[String[]] $MailTo
	)

	Begin {
		# Convert Time to seconds
		# $RefreashInterval *= 60
		$Duration *= 60 * 60
		$EscapeTime = 0
		$StartTime = Get-Date

		$Identity = @()
		$GetJobCmd = "Get-CenteralJob"
		$JobStateMap = @()
		$JobErrorMap = @()
		$FinishedJob = @()

		$Separator = "=" * $Host.UI.RawUI.WindowSize.Width
		$SingleSeparator = "-" * ($Host.UI.RawUI.WindowSize.Width / 2)
	}

	Process {
		$Identity += $JobId
	}

	End {

		# Generate Get-CenteralJob command expression
        switch ($PSCmdlet.ParameterSetName) {

			"type" {
				$typeStr = Convert-ArrayAsString -array $Type -Separator "," -Quote Double
                if($ObjectId) {
                    $GetJobCmd = $GetJobCmd, "-ObjectId $ObjectId" -join " "
                    $GetJobCmd = $GetJobCmd, "`$_.Type -in $typeStr}" -join " | ?{"
                }
                else {
                    $GetJobCmd = $typeStr, $GetJobCmd -join " | %{"
                    $GetJobCmd = $GetJobCmd, "-Type `$_}" -join " "
                }
			}

            "object" {
                $GetJobCmd = $GetJobCmd, "-ObjectId $ObjectId" -join " "
            }

			"id" {
				$idStr = Convert-ArrayAsString -array $Identity -Separator "," -Quote Double
				$GetJobCmd = $idStr, $GetJobCmd -join " | "

			}

		}
		
		while ($EscapeTime -lt $Duration) {

			
			$CenteralJob = Invoke-Expression $GetJobCmd

			$JobStateMap = UpdateJobStateMap -CenteralJob $CenteralJob -JobStateMap $JobStateMap
			$JobErrorMap = UpdateJobErrorMap -CenteralJob $CenteralJob -JobErrorMap $JobErrorMap

			# Find new finished job
			$newFinishedJobId = $JobStateMap | ? State -EQ "Deleted" | ? PreviousState -NE "Deleted" | select -ExpandProperty JobId
			$newFinishedJob = $newFinishedJobId | %{Get-CenteralJob -Identity $_ -IncludeDeleted}
			$FinishedJob += $newFinishedJob

			# Find suspended job
			$SuspendedJob = $CenteralJob | ? State -EQ "Suspended"
			# Find out new suspend job
			$newSuspendJobId = $JobStateMap | ? State -EQ "Suspended" | ? State -NE "Suspended" | select -ExpandProperty JobId
			$newSuspendJob = $SuspendedJob | ?{$newSuspendJobId -contains $_.JobId}
				
			# Clear-Host

            Write-Host ("`n" * 9)

			ShowBasicInfo -Type $Type -Wait $Wait.IsPresent -JobId $JobId -RefreashInterval $RefreashInterval -Duration $Duration `
							-AlertStep $AlertStep -FinishAlert $FinishAlert -SuspendAlert $SuspendAlert -ErrorAlert $ErrorAlert `
							-MailTo $MailTo
			Write-Host $Separator -ForegroundColor Yellow
			ShowCenteralJobs -CenteralJob $CenteralJob
			Write-Host $SingleSeparator
			ShowJobErrorMap -JobErrorMap $JobErrorMap
			Write-Host $SingleSeparator
			ShowSuspendedJob -SuspendedJob $SuspendedJob
            Write-Host $SingleSeparator
            $JobStateMap | ft -AutoSize
			ShowFinishedJob -FinishJob $FinishedJob

            Write-Host $Separator -ForegroundColor Yellow
				
			if ($JobErrorMap.Count -gt 0 -and $ErrorAlert.IsPresent -and $MailTo.Count -gt 0) {
				FireJobErrorMail -JobErrorMap $JobErrorMap -MailTo $MailTo
			}
				
			if ($SuspendAlert.IsPresent -and $MailTo.Count -gt 0 -and $newSuspendJob) {
				FireSuspendJobMail -SuspendedJob $newSuspendJob -MailTo $MailTo
			}

			if ($FinishAlert.IsPresent -and $MailTo.Count -gt 0 -and $newFinishedJob) {
				FireFinishJobMail -FinishJob $newFinishedJob -MailTo $MailTo
			}

			if ($CenteralJob.Count -eq 0) { 

				if ($Wait.IsPresent) {
					Start-SleepProgress -Seconds $RefreashInterval -Activity "Waiting job appear..." -Status "Waiting"
					$EscapeTime = RefreshEscapeTime -StartTime $StartTime
					continue
				}
				elseif ($JobStateMap.Count -eq 0) {
					Write-Host "No job found!"
					if ($PSCmdlet.ParameterSetName = "id") {
						Write-Host "Add argument 'wait' if you want to wait expect job appear!"
					}
					break
				}
				else {
					Write-Host "All jobs done!"
					break
				}
			}
			else {
                # I think it should wait new job appear even if all existing jobs finished.
                <#
				if ($Wait.IsPresent) {
					$Wait = $false
					$EscapeTime = 0
				}
                #>

				$SleepInterval = CalculateSleepInterval -CenteralJob $CenteralJob -DefaultInterval $RefreashInterval
				# Sleeping interval
				Start-SleepProgress $SleepInterval
				$EscapeTime = RefreshEscapeTime -StartTime $StartTime
			}
		}

		
	}

}
# MyUtilities>

# <SPSHealthChecker
Function Get-SharePointServiceHealth {
<#
.SYNOPSIS
To compare current services with reference services which were saved as templates.

.DESCRIPTION
This function is used for compare serive status with template service. If the service status changed,
it will display tempalte status and current status of this service. In addition, it will find relative 
service out if it has and display relative service template and current status for reference.
Note: This function just check running Centeral VM's service status. Non-running VMs are excluded by design.

.PARAMETER Type
Service type including “OS, IIS, TimerJob”. 

.PARAMETER Farm
Centeral farm ID.

.PARAMETER Role
VM role which include "USR", "BOT" and "DCH".

.PARAMETER ComputerName
VM name that should be in running state in Centeral.

.PARAMETER Name
Service name which will compared. It should be service name in template. All template serivces 
will compared if omit.

.PARAMETER MailTo
Mail recipient.

.PARAMETER TemplatePath
Service template folder path. Default is "\\chn.sponetwork.com\tools\SSHC\ServiceTemplate".

.OUTPUTS
No object return. Just display result on console.

.EXAMPLE
PS C:\> Get-SharePointServiceHealth -Farm 697 -MailTo "o365_spo@21vianet.com"
Compare farm 697 USR OS services with all template services and send result to SPO team.

.EXAMPLE
PS C:\> Get-SharePointServiceHealth -ComputerName BOT07697-001 -Type IIS
Compare IIS service with all template services for BOT07697-001.

.EXAMPLE
PS C:\> Get-SharePointServiceHealth -Farm 695 -Role DCH -Name CircuitBreaker
Just compare serice "CircuitBreaker" for all DCH VMs on farm 695.


#>
    Param (
        
        [Parameter(Mandatory=$true,ParameterSetName="FarmSet")]
        [Int] $Farm,
        [Parameter(Mandatory=$true,ParameterSetName="MachineSet")]
        [Alias("CN","ServerName","MachineName")]
        [String[]] $ComputerName,
        [Parameter(ParameterSetName="FarmSet")]
        [ValidateSet("USR","BOT","DCH")]
        [String] $Role = "USR",
        [ValidateSet("OS","IIS","TimerJob")]
        [String] $Type = "OS",
        [String[]] $Name,
        [String[]] $MailTo,
        [String] $TemplatePath = "\\chn.sponetwork.com\tools\SSHC\ServiceTemplate"
    )

    If ($MailTo) {
        $MailContent = ""
        $DateStr = Get-Date -Format "yyyyMMddhhmmss"
        $MailSubject = "SharePoint Service Health Check by $env:USERNAME on $DateStr"
        $MailSender = "SharePoint Service Health Checker <SSHC@21vianet.com>"
    }

    Write-Warning "This command just check running VMs from Centeral. Non-Running VMs will not be checked!"

    # For FarmSet:
    If ($PSCmdlet.ParameterSetName -eq "FarmSet") {
        $CenteralFarm = Get-CenteralFarm -Identity $Farm
        If (!$CenteralFarm) {
            Write-Host "Cannot find farm [$Farm]." -ForegroundColor Red
            Break
        }

        $CenteralVMs = Get-CenteralVM -Farm $CenteralFarm -Role $Role -State Running
        If (!$CenteralVMs) {
            Write-Host "Cannot find running VMs in farm [$Farm]!" -ForegroundColor Red
            Break
        }
    }
    # For machineSet:
    If ($PSCmdlet.ParameterSetName -eq "MachineSet") {
        $InvalidVMs = @()
        $CenteralVMs = @()
        $ComputerName | %{
            $vm = Get-CenteralVM -Name $_ -State Running
            If (!$vm) { $InvalidVMs += $_ }
            Else { $CenteralVMs += $vm }
        }

        If ($InvalidVMs) {
            Write-Host "Below VMs cannot be found in Centeral:"
            Write-Host (Convert-ArrayAsString $InvalidVMs) -ForegroundColor Yellow 
            If ($InvalidVMs.Count -eq $ComputerName.Count) { 
                Write-Host "No running Centeral VMs!"
                Break 
            }
        }
    }


    # Ping test VMs
    $ValidVMs = Test-MachineConnectivity -ComputerName $CenteralVMs.Name -ReturnGood
    $FailedVMs = Test-MachineConnectivity -ComputerName $CenteralVMs.Name -ReturnBad
    
    If ($FailedVMs) {
        Write-Host "Below VMs cannot be ping:"
        Write-Host (Convert-ArrayAsString $FailedVMs) -ForegroundColor Yellow
        If ($ValidVMs -eq 0) {
            Write-Host "No valid VMs!"
            Break
        }
    }

    # Starting process

    Switch ($Type) {

        "OS" {
            Write-Debug "Process OS services..."
            $RefOsService = Import-HealthTemplate -TemplatePath $TemplatePath -Role $Role -Type "OS"
            If (!$RefOsService) {
                Write-Host "Cannot find reference OS service!"
                Break
            }
            If ($Name) {
                $InvalidName = @($Name | Where-Object { $_ -notin $RefOsService.Name })
                If ($InvalidName) {
                    Write-Host "Name $(Convert-ArrayAsString $invalidName) not available in tempalte!" -NoNewline
                    If ($InvalidName.Count -eq $Name.Count) {
                        Write-Host "All Names are not found in template!" -ForegroundColor Red
                        Break
                    }
                }
                $RefOsService = $RefOsService | Where-Object { $_.Name -in $Name }
            }
            
            $OsService = Get-OSService -ComputerName $ValidVMs -ServiceName $RefOsService.Name
            $CompareResult = Compare-SharePointService -ReferenceService $RefOsService -DifferenceService $OsService
            If ($CompareResult) {
                $MailContent += "<H1>OS Serivce</H1><HR>"
                Write-Host "`n-------------"
                Write-Host "OS Service:" -ForegroundColor Yellow
                Write-Host "-------------`n"
                
                $RedundantChecker = @()

                foreach ($cr in $CompareResult) {
                    
                    # Sometimes, service and relative service state changed both. They will appear in compared result both. 
                    # Add redundant checker to trim redundance outputs.
                    If ($RedundantChecker.RelativeIndex -notcontains $cr.RelativeIndex) {
                        $RedundantChecker += $cr
                        $RelativeService = Get-RelativeOSService -Name $cr.Name -ComputerName $cr.SystemName -ReferenceService $RefOsService
                        Out-OSService -Service $cr -RelativeService $RelativeService -ReferenceService $RefOsService
                    }
                    
                    If ($MailTo) {
                        $MailContent += Write-MailContent -Service $cr -RelativeService $RelativeService -ReferenceService $RefOsService
                    }
                }
            }
            else {
                Write-Host "OS services have not been changed!" -ForegroundColor Green
            }
            break
        }

        "IIS" {
            Write-Host "It is coming soon!"
            break
        }

        "TimerJob" {
            Write-Host "It is coming soon!"
            break
        }

    }


    If ($MailTo) {
        Send-Email -To $MailTo -From $MailSender -mailsubject $MailSubject -mailbody $MailContent -BodyAsHtml
    }

}

Function Get-SharePointTemplateServiceName {
<#
.SYNOPSIS
Return service name as string array.

.DESCRIPTION

.PARAMETER Type
Service type including “OS, IIS, TimerJob”. Default is 'OS'.

.PARAMETER Role
VM role which include "USR", "BOT" and "DCH". Default is "USR".

.PARAMETER TemplatePath
Service template folder path. Default is "\\chn.sponetwork.com\tools\SSHC\ServiceTemplate".

#>

    Param (
        [ValidateSet("USR","BOT","DCH")]
        [String] $Role = "USR",
        [ValidateSet("OS","IIS","TimerJob")]
        [String] $Type = "OS",
        [String] $TemplatePath = "\\chn.sponetwork.com\tools\SSHC\ServiceTemplate"
    )

    $RefOsService = Import-HealthTemplate -TemplatePath $TemplatePath -Role $Role -Type $Type
    $Name = $RefOsService.Name

    return $Name
}
# SPSHealthChecker>

function Get-PAVCResult
{

<#
.SYNOPSIS
Scan patches and vulns from patch label maps and machines.

.DESCRIPTION
Scan patches and vulns from patch label maps and machines.

.PARAMETER $BaselineDays
The past days that which vulns in it will not report. 
The range is 0 to -90. Default is -30 days.

.PARAMETER $Credential
The credential for CHN domain admin.

.PARAMETER $MailTo
Mail recipient. Default is "O365_SPO@21vianet.com".

.PARAMETER PassThru
It will return vuln objects to console if it present.

.EXAMPLE
Get-PAVCResult
Scan PAVC from all machines and send mail to SPO mail group if vulns have.

.EXAMPLE
Get-PAVCResult -MailTo 'xu.jiankun@oe.21vianet.com'
Just send PAVC result to JK if it have.

.EXAMPLE
Get-PAVCResult -PassThru
Send result to SPO group and return result as objects to PowerShell console.
#>

    Param (
        [ValidateScript({$_ -le 0 -and $_ -ge -90})]
        [Int] $BaselineDays = -30,
        [PSCredential] $Credential,
        [String[]] $MailTo = 'o365_spo@21vianet.com',
        [Switch] $PassThru
    )

    $baseline = (get-date).AddDays($BaselineDays)
    $time = get-date
    if (!$Credential) {
        $Credential = Get-Credential -Message "Enter domain admin credential for DSE query"
    }
    $hostVM = $env:COMPUTERNAME
    $DSE = Get-CenteralVM -Role DSE | select -ExpandProperty name
    $Farm = Get-Centeralfarm | ? Role -notlike *DS* | ? Farmid -ne 609 | ? Farmid -ne 618 | ? Role -notlike *AdminDebug* 
    $MOR = Get-CenteralVM -Role MOR | ? name -ne $hostvm | select -ExpandProperty name
    $ISB = 609,618 | %{Get-CenteralVM -Farm $_} | ? name -notlike *DSE* | select -ExpandProperty name
    $spd = Get-SPDFQDN
    $MailSubject = "PAVC Report on $time by $env:username"

    $PMtestpool = 6,7,12,13 | % {Get-CenteralPatchLabelMap -zone $_} | ? State -ne finished
    $vmtestpool = $Farm | % {Get-CenteralPatchLabelMap -Farm $_.FarmId} | ? State -ne finished | ? ObjectType -ne PHYSICALMACHINE | ? State -ne BadStateMachine

    ## Update
    $fullvuln = @()
    $MachinePool = @()
    $MachinePool += $PMtestpool | % {Get-CenteralPMachine -Identity $_.objectid} | select -ExpandProperty Name
    $MachinePool += $vmtestpool | % {Get-CenteralVM -Identity $_.objectid} | ? State -eq Running | select -ExpandProperty Name
    $MachinePool += $spd
    $MachinePool += $MOR

    # Test machine pool
    $fullvuln += Test-WindowsUpdates $MachinePool
    # Test host self
    $fullvuln += Test-WindowsUpdates
    # Test DSE with domain credential
    $fullvuln += Test-WindowsUpdates $DSE -Credential $Credential

    ## Update end

    $uniquevuln = $fullvuln | ? LastDeploymentChangeTime -lt $baseline | select PSComputerName, Title, LastDeploymentChangeTime

    if ($uniquevuln) {
        $htmltype = Format-HtmlTable -Contents $uniquevuln
        Send-Email -To $MailTo -mailbody $htmltype -mailsubject $MailSubject -From 'SPOPAVCCheck@21vianet.com' -BodyAsHtml

        if ($PassThru) { return $uniquevuln }
    }
}

function Test-F5Login{

<#
.Synopsis
This fucntion is using to validate F5 account is working or not when update it's password.

.Description
The F5 account “auto_mgr_chn_spo01” is used for change F5 configuration by Centeral job such as DeployVM. 
This account has been migrated to CME domain. It’s a service account without logon permission. 
Password policy in CME force all of accounts will be expired after 70 days.

.Parameter UserName
The F5 login user name. Default is 'auto_mgr_chn_spo01'.

.Parameter Password
The login user's password.

.Parameter F5IP
F5 devices' IP. Default are '10.41.44.5', '10.41.240.5'.

.Outputs
PSObject include below fields:
    UserName: F5 account.
    RemoteAddress: F5 IP address.
    RemotePort: connectivity testing port.
    TcpTestSucceeded: connectivity testing result.
    F5Login: F5 login testing result.

.Example
Test-F5Login -Password 'yourpassword'
Testing with default user name, F5 IPs and your provided password.
#>

    param (
        [String]$UserName = 'auto_mgr_chn_spo01',
        [String]$Password,    # password must be quoted with single qute mark to avoid escape character.
        [String[]]$F5IP = @('10.41.44.5', '10.41.240.5')
    )

    $Result = @()
    
    Add-PSSnapIn iControlSnapIn

    foreach ($ip in $f5IP) {
       
        $connection = Test-NetConnection -ComputerName $ip -Port 443
        $F5Pass = Initialize-F5.iControl -HostName $ip -Username $UserName -Password $Password

        $Result += New-Object -TypeName PsObject -Property @{ RemoteAddress = $ip; RemotePort = $connection.RemotePort;
                    TcpTestSucceeded = $connection.TcpTestSucceeded; F5Login = $F5Pass; UserName = $UserName }
    }

    return $Result

} 



