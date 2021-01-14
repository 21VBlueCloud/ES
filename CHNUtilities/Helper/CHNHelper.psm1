# <CHNMachineFunctions
Function PingTest {
    Param ([string] $ComputerName)
    # update 1.1.3: add IP address in result.
    # $IPAddress = Resolve-DnsName $ComputerName | select -ExpandProperty ipaddress
    $TestConnection = Test-Connection $ComputerName -Count 1
    if ($TestConnection.IPV4Address) {$Ping = $true}
    else {$Ping = $false}
    $Properties = @{ComputerName=$ComputerName;IPAddress=$TestConnection.IPV4Address;Ping=$Ping}
    $Result = New-Object -TypeName PSObject -Property $Properties
    Return $Result
    }

Function PSsessionTest {
    Param (
    [String] $ComputerName,[System.Management.Automation.PSCredential] $Credential
    )

    $command = "New-PSSession -ComputerName $ComputerName -ErrorAction SilentlyContinue"
    if ($Credential) { $command += " -Credential `$Credential" }

    $session = Invoke-Expression $command
    If ($session) {
        $PSSession = $true
        Remove-PSSession -Session $session
    }
    Else {
        $PSSession = $false
    }

    $Properties = @{ComputerName=$ComputerName;PSSession=$PSSession}
    $Result = New-Object -TypeName PSObject -Property $Properties
    Return $Result
}


Function StartLocalBatchJobs {
<#
.SYNOPSIS
Kick off local PS jobs for connectivity test in batch mode.

#>

    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [ScriptBlock] $ScriptBlock,
        [Parameter(Mandatory=$true,ParameterSetName="Computer")]
        [String[]] $ComputerName,
        [System.Management.Automation.Runspaces.AuthenticationMechanism] $Authentication,
        [PSCredential] $Credential,
        [Int] $ThreadWindow = 20,
        [Int] $TimeOut=120 #seconds
    )

    $command = "Start-Job -ScriptBlock `$ScriptBlock"
    #compose command statement
    # if ($ArgumentList) { $command += " -ArgumentList `$ArgumentList" }
    if ($Authentication) { $command += " -Authentication `$Authentication" }
    if ($Credential) { $command += " -Credential `$Credential" }
    if ($InputObject) { $command += " -InputObject `$InputObject" }
    if ($Name) { $command += " -Name `$Name" }
    if ($PSVersion) { $command += " -PSVersion `$PSVersion" }
    
    $Result = @()

    # start to kick off local jobs for test remote computer
    if ($PSCmdlet.ParameterSetName -eq "Computer") {
        $ThreadQueue = New-Object -TypeName System.Collections.Queue -ArgumentList (,$ComputerName)
        $ActiveThread = @{}
        
        While ($ThreadQueue.Count) {
            While ($ActiveThread.Count -lt $ThreadWindow) {
                # Fullfill active thread to window size and kick off job
                $Value = $ThreadQueue.Dequeue()
                $expression = $command, "-ArgumentList $Value" -join " "
                $job = Invoke-Expression $expression
                $Key = $job.InstanceId
                $ActiveThread.Add($Key,$Value)
                if ($ThreadQueue.Count -eq 0) { break }  # should be break in case of thread queue empty.
            }
            # Watch jobs
            $PreJobs = Get-Job
            $Result += Watch-Jobs -Status "Executing" -NoWait -TimeOut $TimeOut
            $PostJobs = Get-Job
            if ($PostJobs) {
                $RemovedJobs = Compare-Object $PreJobs $PostJobs | select -ExpandProperty InputObject
            }
            else {
                $removedJobs = $PreJobs
            }
            foreach ($job in $RemovedJobs) {
                $ActiveThread.Remove($job.InstanceId)
            }
            
        }

        # To wait jobs end up
        if ($ActiveThread.Count) {
            # Watch jobs
            $Result += Watch-Jobs -Status "Executing"
        }

    }

    return $Result

}
# CHNMachineFunctions>

# <CHNOEAccountFunctions
Function Read-CSV {
    
    $folder = "d:\temp\CSVGallatin"
    $old_folder = "d:\temp\CSV_OLD"
    If (Test-Path $folder) {
        $Files = Get-ChildItem $folder
        $Files = $Files | ? Extension -eq ".csv"
        $accounts = @()
        foreach ( $csv in $Files) {
            $accounts += Import-Csv -Path $csv.FullName
            Move-Item -Path $csv.FullName -Destination $old_folder -Force
        }
    }
    Else {
        Write-Host "$folder does not exist!"
    }
    Return $accounts
}

Function Send-NotificationMail {
    
    Param (
        [System.Net.Mail.MailAddress] $From,
        [System.Net.Mail.MailAddress] $To,
        [String] $Subject,
        [String] $Body
    )

    If ($From -eq $null) {
        $AdUser = Get-ADUser $env:USERNAME -Properties EmailAddress
        $From = New-Object -TypeName System.Net.Mail.MailAddress -ArgumentList @($AdUser.EmailAddress,$AdUser.Name)
    }

    Send-Email -To $To -mailbody $Body -mailsubject $Subject -From $From

}
# CHNOEAccountFunctions>

# <CHNPerformance
function GetOSPerformance {
    Param (
        [ValidateSet("CPU","Memory","Network","Disk")]
        [String[]] $Types
    )

    if (!$Types) {
        $Types = "CPU","Memory","Network","Disk"
    }

    $Result = @{}

    if ($Types -contains "CPU") {
        $CPUWmiObj = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Processor `
                | Where-Object Name -EQ "_Total"
        $Properties = @{
            ComputerName = $env:COMPUTERNAME
            Type = "CPU"
            Utilization = 1 - $CPUWmiObj.PercentIdleTime / 100
        }
        $CPU = New-Object -TypeName PSObject -Property $Properties
        $Result.Add("CPU",$CPU)
    }

    if ($Types -contains "Memory") {
        $PhysicalMemory = Get-WmiObject -Class Win32_PhysicalMemory
        $MemoryWmiObj = Get-WmiObject -Class Win32_PerfFormattedData_PerfOS_Memory
        $Total = 0
        $PhysicalMemory | ForEach-Object {$Total +=$_.Capacity}
        $Available = $MemoryWmiObj.AvailableBytes
        $Utilization = 1 - $Available / $Total
        $Properties = @{
            ComputerName = $env:COMPUTERNAME
            Type = "Memory"
            Total = $Total
            Available = $Available
            Utilization = $Utilization
        }
        $Memory = New-Object -TypeName PSObject -Property $Properties
        $Result.Add("Memory",$Memory)
    }

    if ($Types -contains "Network") {
        $Network = @()
        $NetworkWmiObj = @(Get-WmiObject -Class Win32_PerfFormattedData_Tcpip_NetworkInterface)
        foreach ($obj in $NetworkWmiObj) {
            $Properties = @{
                ComputerName = $env:COMPUTERNAME
                Type = $obj.Name
                Bandwidth = $obj.CurrentBandwidth
                ReceivedBps = $obj.BytesReceivedPersec
                SentBps = $obj.BytesSentPersec
                TotalBps = $obj.BytesTotalPersec
            }
            $Network += New-Object -TypeName PSObject -Property $Properties
        }
        $Result.Add("Network",$Network)
    }

    if ($Types -contains "Disk") {
        $DiskWmiObj = Get-WmiObject -Class Win32_PerfFormattedData_PerfDisk_LogicalDisk `
                        | Where-Object Name -EQ "_Total"
        $Properties = @{
        ComputerName = $env:COMPUTERNAME
            Type = "Disk"
            Activity = 1 - $DiskWmiObj.PercentIdleTime / 100
        }
        $Disk = New-Object -TypeName PSObject -Property $Properties
        $Result.Add("Disk",$Disk)
    }

    return $Result
}

function AddResultInHashTable {
    Param(
        [Parameter(Mandatory=$true)]
        $InputResult,
        $hash
    )

    $Name = $InputResult.Keys -as [string[]]

    if ($hash) {
        foreach ($n in $Name) {
            $valueArray = $hash[$n]
            switch ($n) {
                "CPU" {
                    $value = $InputResult[$n].Utilization
                    $valueArray += $value
                    $hash[$n] = $valueArray
                    break
                }
                "Memory" {
                    $value = $InputResult[$n].Utilization
                    $valueArray += $value
                    $hash[$n] = $valueArray
                    break
                }
                "Disk" {
                    $value = $InputResult[$n].Activity
                    $valueArray += $value
                    $hash[$n] = $valueArray
                    break
                }
                default {
                    $NetworkArray = @($InputResult[$n])
                    foreach ($na in $NetworkArray) {
                        $valueArray = $hash[$na.Type]
                        $valueArray += $na.TotalBps
                        $hash[$na.Type] = $valueArray
                    }
                }
            }
        }
    } # end if hash
    else {
        $hash = [ordered] @{}
        foreach ($n in $Name) {
            switch ($n) {
                "CPU" {
                    $hash.Add($n,@($InputResult[$n].Utilization))
                    break
                }

                "Memory" {
                    $hash.Add($n,@($InputResult[$n].Utilization))
                    break
                }

                "Network" {
                    # We don't know how many network adapter are there.
                    $NetworkArray = @($InputResult[$n])
                    foreach ($na in $NetworkArray) {
                        $hash.Add($na.Type, @($na.TotalBps))
                    }
                    break
                }

                "Disk" {
                    $hash.Add($n,@($InputResult[$n].Activity))
                    break
                }
            }
        }
    }
    return $hash
}

function CalculateOutputs {
    param (
        [Parameter(Mandatory=$true)]
        $hash,
        $Outputs
    )
    $Name = $hash.Keys -as [String[]]
    if ($Outputs) {
        foreach ($n in $Name) {
            $obj = $Outputs | ? Type -EQ $n
            if (!$obj) { continue }
            $valueArray = $hash[$n]
            $lastValue = $valueArray[-1]
            $obj.Current = $lastValue
            if ($valueArray.count - 1 -ne 0) {
                $obj.Avg = ($obj.Avg * ($valueArray.count-1) + $lastValue) / $valueArray.count
                if ($obj.Max -lt $lastValue) { $obj.Max = $lastValue }
                if ($obj.Min -gt $lastValue) { $obj.Min = $lastValue }
            }
            else {
                $obj.Avg = $lastValue
                $obj.Max = $lastValue
                $obj.Min = $lastValue
            }

        
        }
    } # if $outputs end
    else {
        $outputs = @()
        foreach ($t in $Name) {
            $properties = @{
                Type = $t
                Current = $hash[$t][0]
                Max = $hash[$t][0]
                Min = $hash[$t][0]
                Avg = $hash[$t][0]
            }
            $outputs += New-Object -TypeName PSObject -Property $properties
        }
    } # if $outputs else end
    return $Outputs
}

function DisplayOutputs {
    param (
        $Outputs,
        $ComputerInfo,
        [switch] $NoRefresh
    )

    $separator = "=" * ($Host.UI.RawUI.WindowSize.Width / 2)

    $CurrentPercent = @{Name="Current";Expression={$_.Current -as [Single]};FormatString="P0";Alignment="Left";Width=7}
    $CurrentKbps = @{Name="Current";Expression={($_.Current -as [Single]) / 1Kb};FormatString="#,0 'Kbps'";Alignment="Left";Width=7}

    $MaxPercent = @{Name="Maximum";Expression={$_.Max -as [Single]};FormatString="P0";Alignment="Left";Width=7}
    $MaxKbps = @{Name="Maximum";Expression={($_.Max -as [Single]) / 1Kb};FormatString="#,0 'Kbps'";Alignment="Left";Width=7}

    $MinPercent = @{Name="Minimum";Expression={$_.Min -as [Single]};FormatString="P0";Alignment="Left";Width=7}
    $MinKbps = @{Name="Minimum";Expression={($_.Min -as [Single]) / 1Kb};FormatString="#,0 'Kbps'";Alignment="Left";Width=7}

    $AvgPercent = @{Name="Average";Expression={$_.Avg -as [Single]};FormatString="P0";Alignment="Left";Width=7}
    $AvgKbps = @{Name="Average";Expression={($_.Avg -as [Single]) / 1Kb};FormatString="#,0 'Kbps'";Alignment="Left";Width=7}

    $Utilization = "CPU", "Memory"
    $Activity = "Disk"
    
    $returnline = "`n" * 6

    if (!$NoRefresh.IsPresent) {
        Clear-Host
    }
    else {
        Write-Host $separator
    }
    # Write-Host $returnline
    $ComputerInfo | Format-List ComputerName, Processor, Memory, Network*

    foreach ($o in $Outputs) {
        switch ($o.Type) {
            {$o.Type -in $Utilization} {
                Write-Host "$($o.Type)" -ForegroundColor Green -NoNewline
                Write-Host " Utilization"
                $o | Format-Table -AutoSize -Property $CurrentPercent,$MaxPercent,$MinPercent,$AvgPercent
                break
            }

            {$o.Type -in $Activity} {
                Write-Host "$($o.Type)" -ForegroundColor Green -NoNewline
                Write-Host " Activity"
                $o | Format-Table -AutoSize -Property $CurrentPercent,$MaxPercent,$MinPercent,$AvgPercent
                break
            }
            
            default {
                Write-Host "Network ($($o.Type))" -ForegroundColor Green -NoNewline
                Write-Host " Throughput"
                $o | Format-Table -AutoSize -Property $CurrentKbps,$MaxKbps,$MinKbps,$AvgKbps
            }
        }
        
        
    }

}

Function GetOSInfo {
    $processor = Get-WmiObject win32_processor
    $PMemory = Get-WmiObject Win32_PhysicalMemory
    $NetworkInterface = Get-WmiObject Win32_PerfFormattedData_Tcpip_NetworkInterface
    $TotalMemory = 0
    $PMemory | %{$TotalMemory += $_.capacity}
    $Memory = "{0:#,0 'MB'}" -f ($TotalMemory / 1MB)
    $Network = @()
    
    $properties = @{
        ComputerName = $env:ComputerName
        Processor = $processor.Name
        Memory = $Memory
    }

    $BasicInfo = New-Object -TypeName PSObject -Property $properties
    for ($i=0;$i -lt $NetworkInterface.count;$i++) {
        $Network = "Network$i"
        $bandwidth = $NetworkInterface[$i].CurrentBandwidth / (1000 * 1000)
        $NetworkName = $NetworkInterface[$i].Name

        if ($bandwidth -gt 1000) {
            $NetworkName += "({0:#,0.0 'Gbps'})" -f ($bandwidth / 1000)
        }
        else {$NetworkName += "({0:#,0.0 'Mbps'})" -f $bandwidth}
        $BasicInfo | Add-Member -NotePropertyName $Network -NotePropertyValue $NetworkName
    }

    return $BasicInfo
}

function InitializeRootFolder {
    Param (
        [String] $Path,
        [String[]] $Type,
        [String[]] $ComputerName
    )

    # Get the date string from path.
    $DateString = ($Path -split "_")[1]

    # Create host folder
    $HostFolders = @()
    foreach ($computer in $ComputerName) {
        $HostFolders += Join-Path $Path ($Computer, $DateString -join "_")
    }

    try {
        $HostFolderPath = New-Item -ItemType Directory -Path $HostFolders -ErrorAction Stop | select -ExpandProperty FullName
    }
    catch {
        Remove-Item -Path $HostFolders -Recurse -Force -ErrorAction Ignore -Confirm:$false
        Write-Host $_ -ForegroundColor Red
        
    }

    <#
    # Create performance statistic file in each host folder
    try {
        $StatisticFile = New-Item -ItemType File -Path $HostFolders -Name PerformanceStatistic.csv -ErrorAction Stop
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Remove-Item -Path $HostFolders -Recurse -Confirm:$false -Force
    }
    #>

    # Create type folder
    $TypeFolders = @()
    foreach ($t in $Type) {
        $TypeFolders += Join-Path $HostFolders $t
    }
    try {
        $TypeFolderPath = New-Item -ItemType Directory -Path $TypeFolders -ErrorAction Stop | select -ExpandProperty FullName
    }
    catch {
        Remove-Item -Path $HostFolders -Recurse -Confirm:$false -Force -ErrorAction Ignore
        Write-Host $_ -ForegroundColor Red
        
    }

    <#
    # Create record file in each type folder
    try {
        $RecordFile = New-Item -ItemType File -Path $TypeFolders -Name Record.csv -ErrorAction Stop
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        Remove-Item -Path $HostFolders -Recurse -Confirm:$false -Force
    }
    #>

    $result = @{ StatisticPath = $HostFolderPath; RecordPath = $TypeFolderPath }
    return $result

}

function WriteOSInfo {
    Param (
        [String] $Path,
        [Object[]] $OSInfo
    )

    $HostFolders = Get-ChildItem -Path $Path -Directory | select -ExpandProperty FullName
    foreach ($info in $OSInfo) {
        $folder = $HostFolders | ?{$_ -match $info.ComputerName}
        if ($folder) {
            $InfoPath = Join-Path $folder "BasicInfo"
            $info | select ComputerName, Processor, Memory, Network* | Out-File "$InfoPath.txt"
            $info | select ComputerName, Processor, Memory, Network* | Export-Csv "$InfoPath.csv"
        }
        else {
            Write-Host "Cannot find folder for host $($info.ComputerName) in $Path!" -ForegroundColor Yellow
        }
    }
}

function CalculateStatistics {
    Param (
        [Parameter(Mandatory=$true)]
        [Object[]] $InputObject,
        [Object[]] $Statistic
    )

    if ($Statistic) {
        foreach ($obj in $InputObject) {
            $Type = $obj.Keys -as [String[]]
            foreach ($t in $Type) {
                $record = $obj[$t]
                
                switch ($t) {
                    {$t -in "CPU","Memory"} {
                        $s = $Statistic | ? ComputerName -EQ $record.ComputerName | ? Type -EQ $t
                        $value = $record.Utilization
                        $s.Current = $value
                        if ($value -gt $s.Max) {$s.Max = $value}
                        if ($value -lt $s.Min) {$s.Min = $value}
                        $s.Avg = ($s.Count * $s.Avg + $value ) / ++$s.Count
                        break
                    }

                    "Disk" {
                        $s = $Statistic | ? ComputerName -EQ $record.ComputerName | ? Type -EQ $t
                        $value = $record.Activity
                        $s.Current = $value
                        if ($value -gt $s.Max) {$s.Max = $value}
                        if ($value -lt $s.Min) {$s.Min = $value}
                        $s.Avg = ($s.Count * $s.Avg + $value ) / ++$s.Count
                        break
                    }

                    default {
                        foreach ($r in $record) {
                            $s = $Statistic | ? ComputerName -EQ $r.ComputerName | ? Type -EQ $r.Type
                            $value = $r.TotalBps
                            $s.Current = $value
                            if ($value -gt $s.Max) {$s.Max = $value}
                            if ($value -lt $s.Min) {$s.Min = $value}
                            $s.Avg = ($s.Count * $s.Avg + $value ) / ++$s.Count
                        }
                    }
                }
                
            }
        }
    }
    else {
        # Initialize statistic
        foreach ($obj in $InputObject) {
            $Type = $obj.Keys -as [String[]]
            foreach ($t in $Type) {
                $record = $obj[$t]
                switch ($t) {
                    {$_ -in "CPU","Memory"} {
                        $properties = @{
                            ComputerName = $record.ComputerName
                            Type = $_
                            Count = 1
                            Current = $record.Utilization
                            Max = $record.Utilization
                            Min = $record.Utilization
                            Avg = $record.Utilization
                        }
                        $Statistic += New-Object -TypeName PSObject -Property $properties
                        break
                    }
                    "Disk" {
                        $properties = @{
                            ComputerName = $record.ComputerName
                            Type = $_
                            Count = 1
                            Current = $record.Activity
                            Max = $record.Activity
                            Min = $record.Activity
                            Avg = $record.Activity
                        }
                        $Statistic += New-Object -TypeName PSObject -Property $properties
                        break
                    }
                    default {
                        foreach ($subobj in $record) {
                            $properties = @{
                                ComputerName = $subobj.ComputerName
                                Type = $subobj.Type
                                Count = 1
                                Current = $subobj.TotalBps
                                Max = $subobj.TotalBps
                                Min = $subobj.TotalBps
                                Avg = $subobj.TotalBps
                            }
                            $Statistic += New-Object -TypeName PSObject -Property $properties
                        }
                    }
                }
            }
        }
    }

    return $Statistic
}

function WriteRecords {
    Param (
        [String[]] $Path,
        [System.Collections.Hashtable[]] $InputObject
    )

    foreach ($hash in $InputObject) {
        $Type = $hash.Keys -as [string[]]
        foreach ($t in $Type) {
            $record = $hash[$t]
            
            # different treatment for network and others
            if ($t -ne "Network") {
                # Find out store path
                $StorePath = $Path | ?{$_ -match $record.ComputerName}
                $StorePath = $StorePath | ?{$_ -match $record.Type}

                $FilePath = Join-Path $StorePath "$t.csv"
                $record | Export-Csv -Path $FilePath -Append
            }
            else {
                foreach ($r in $record) {
                    $StorePath = $Path | ?{$_ -match $r.ComputerName}
                    $StorePath = $StorePath | ?{$_ -match $t}
                    $FilePath = Join-Path $StorePath "$($r.Type).csv"
                    $r | Export-Csv -Path $FilePath -Append
                }
            }
        }
    }
}

function WriteStatisticRecords {
    Param (
        [String[]] $Path,
        [Object[]] $InputObject
    )

    $GroupedObject = $InputObject | group ComputerName
    foreach ($group in $GroupedObject) {
        $StorePath = $Path | ?{$_ -match $group.Name}
        $FilePath = Join-Path $StorePath "Statistics.csv"
        $group.Group | Export-Csv -Path $FilePath
    }

}

function ShowStatistics {
    Param (
        [Object[]] $Statistics,
        [Object[]] $OsInfo
    )

    foreach ($info in $OsInfo) {
        $statistic = $Statistics | ? ComputerName -EQ $info.ComputerName
        DisplayOutputs -Outputs $statistic -ComputerInfo $info -NoRefresh
    }
}
# CHNPerformance>

# <CHNTenantFunctions
Function Format-TenantUri {
<#
.SYNOPSIS
This function is used for dealing with tenants.
.DESCRIPTION
If the address not contain prefix 'http://' or 'https://', add them.
#>


    Param ([Parameter(Mandatory=$true)][String] $address)
    
    If ($address -match "sharepoint`.cn") {
        
        If ($address -match "^http://") { Return $address -replace "http://", "https://" }

        If ($address -match "^https://") { Return $address }

        Return "https://" + $address

    }

    If ($address -match "^http://" -or $address -match "^https://") {
        Return $address
    }

    Return "http://" + $address

}

Function Watch-TenantStatus {

    Param (
        [Parameter(Mandatory=$true)]
        [Microsoft.SharePointCenteral.TenantMgr.Tenant[]] $tenants,
        [Parameter(Mandatory=$true,ParameterSetName="EnabledTenant")]
        [Switch] $Enabled,
        [Parameter(Mandatory=$true,ParameterSetName="DisabledTenant")]
        [Switch] $Disabled
    )

    $CandidateTenants = @()
    $FunctionStartTime = Get-Date

    # Starting sleep 70 sec for waiting tenants state changed.
    Write-Host "Operation is progressing.." -BackgroundColor Yellow -ForegroundColor DarkGreen
    Start-Sleep -Seconds 70
    

    while ($CandidateTenants.Count -lt $tenants.Count) {

        $LoopStartTime = Get-Date

        # Terminate the loop in case of infinite loop.
        If (($LoopStartTime - $FunctionStartTime).Seconds -gt 180) {
            Write-Host "Waiting window has been expired!`nPlease check tenants status manually." -ForegroundColor Yellow
            Break
        }

        #update tenants status first.
        $tenants = $tenants | %{Get-CenteralTenant -Identity $_}
        If ($PSCmdlet.ParameterSetName -eq "DisabledTenant") {
            $CandidateTenants = $tenants | ?{$_.State -eq "Disabled"}
        }
        If ($PSCmdlet.ParameterSetName -eq "EnabledTenant") {
            $CandidateTenants = $tenants | ?{$_.State -eq "Active"}
        }
        

        # Wait 5 sec for retrial.
        Start-Sleep -Seconds 5
    }

    If ($CandidateTenants.Count -eq $tenants.Count) {
        If ($PSCmdlet.ParameterSetName -eq "DisabledTenant") {
            Write-Host "All of tenants has been disabled!" -ForegroundColor Green
        }

        If ($PSCmdlet.ParameterSetName -eq "EnabledTenant") {
            Write-Host "All of tenants has been Enabled!" -ForegroundColor Green
        }
        
    }
}
# CHNTenantFunctions>

# <CHNTopologyFunctions
Function Format-CHNTopologyHtmlTable {
    
    Param (
        [Parameter(Mandatory=$true,
                    Position=0)]
        [Object[]] $Contents,
        [String] $Title,
        [Int] $TableBorder=1,
        [Int] $Cellpadding=0,
        [Int] $Cellspacing=0,
        [Parameter(Mandatory=$true,
                    ParameterSetName="Item")]
        [Switch] $ColorItems,
        [Parameter(ParameterSetName="Item")]
        [String] $DeletedBgcolor="#FFFF00",
        [Parameter(ParameterSetName="Item")]
        [String] $NewBgcolor="#00FF00",
        [Parameter(Mandatory=$true,
                    ParameterSetName="Property")]
        [Switch] $ColorPorperties

    )
    $TableHeader = "<table border=$TableBorder cellpadding=$Cellpadding cellspacing=$Cellspacing"

    # Add title
    $HTMLContent = $Contents | ConvertTo-Html -Fragment -PreContent "<H1>$Title</H1>"
    #Add table border
    $HTMLContent = $HTMLContent -replace "<table", $TableHeader
    # Strong table header
    $HTMLContent = $HTMLContent -replace "<th>", "<th><strong>"
    $HTMLContent = $HTMLContent -replace "</th>", "</strong></th>"

    # Color Items. Yellow highlight deleted items and Green highlight new items.
    If ($ColorItems) {
        
        for ($i=0;$i -lt $HTMLContent.count;$i++) {
            # Color Deleted items.
            If ($HTMLContent[$i].IndexOf("<td>Deleted</td>") -ge 0) {
                $InsertString = " bgcolor=" + $DeletedBgcolor
                $FindStr = "<tr"
                $Position = $HTMLContent[$i].IndexOf($FindStr) + $FindStr.Length
                $HTMLContent[$i] = $HTMLContent[$i].Insert($Position,$InsertString)
            }

            # Color New items.
            If ($HTMLContent[$i].IndexOf("<td>New</td>") -ge 0) {
                $InsertString = " bgcolor=" + $NewBgcolor
                $FindStr = "<tr"
                $Position = $HTMLContent[$i].IndexOf($FindStr) + $FindStr.Length
                $HTMLContent[$i] = $HTMLContent[$i].Insert($Position,$InsertString)
            }

        }
    }

    # Color Properties.
    If ($ColorPorperties) {
        # Coming soon...
    }

    Return $HTMLContent
}

Function Compare-CHNItems {
    Param (
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceItems,
        [Parameter(Mandatory=$true)]
        [Object[]] $CurrentItems,
        [String] $GroupField = "Role",
        [Parameter(Mandatory=$true)]
        [String] $CompareField
    )

    Write-Debug "Start items comparison ..."

    $Results = @()

    $RefGroups = $ReferenceItems | Group-Object $GroupField | Sort-Object name
    $CurGroups = $CurrentItems | Group-Object $GroupField | Sort-Object Name

    $GrpCompareRlts = Compare-Object $RefGroups $CurGroups -Property Name -IncludeEqual

    foreach ($GrpResult in $GrpCompareRlts) {
        
        # If side indicator is ==, it means this group should appeared in both groups. Continue to comapre items of this group.
        If ($GrpResult.SideIndicator -eq '==') {
            $refItems = $RefGroups | ? Name -EQ $GrpResult.Name | Select-Object -ExpandProperty group
            $curItems = $CurGroups | ? Name -EQ $GrpResult.Name | Select-Object -ExpandProperty group

            $ItemCompareRlts = Compare-Object $refItems $curItems -Property $CompareField

            If ($ItemCompareRlts) {

                foreach ($ItemResult in $ItemCompareRlts) {

                    # It means the Item just appeared in refItems, sould be a deleted one.
                    If ($ItemResult.SideIndicator -eq '<=') {
                        $DeletedItem = $refItems | ? $CompareField -EQ $($ItemResult.$CompareField)
                        $DeletedItem | Add-Member -NotePropertyName RecordAction -NotePropertyValue Deleted
                        $Results += $DeletedItem
                    }

                    # It means the Item just appeared in curItems, sould be a new one.
                    If ($ItemResult.SideIndicator -eq '=>') {
                        $NewItem = $curItems | ? $CompareField -EQ $($ItemResult.$CompareField)
                        $NewItem | Add-Member -NotePropertyName RecordAction -NotePropertyValue New
                        $Results += $NewItem
                    }
                }
            }
        }

        # If side indicator is <=, it means the role just appeared in reference group that is should not in current group. Role must be deleted.
        If ($GrpResult.SideIndicator -eq '<=') {
            $DeletedItems = $RefGroups | ? Name -EQ $GrpResult.Name | Select-Object -ExpandProperty group
            $DeletedItems | Add-Member -NotePropertyName RecordAction -NotePropertyValue Deleted
            $Results += $DeletedItems
        }

        # If side indicator is =>, it means the role just appeared in current group that is should not in reference group. Role must be new.
        If ($GrpResult.SideIndicator -eq '=>') {
            $NewItems = $CurGroups | ? Name -EQ $GrpResult.Name | Select-Object -ExpandProperty group
            $NewItems | Add-Member -NotePropertyName RecordAction -NotePropertyValue New
            $Results += $NewItems
        }
    }
    Write-Debug "Comparison compeleted!"
    Return $Results
}

Function Compare-CHNProperties {
    
    Param (
        [String] $Properties
    )



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
        [Switch] $SnapShot,
        [Parameter(Mandatory=$true,ParameterSetName="FormerComparison",Position=0)]
        [System.Collections.Hashtable] $DifferenceTopology,
        [Parameter(Mandatory=$true,ParameterSetName="FormerComparison",Position=1)]
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

    $Results = @()
    $MailBody = @()

    If ($Types -eq "All") {
            $Types = @("VMs","PMs","Farms","Networks","Zones","NLBs","Vlans","VMIPs")
    }

    $date = Get-Date
    $StrDate = Get-Date -Format "MM/dd/yyyy"

    # Update for 1.2.3
    If ($PSCmdlet.ParameterSetName -eq "CurrentComparison") {
        
        Write-Host "Start current topology comparison..."

        If ($Last -eq 0) {
    
            $Allversion = Get-CHNTopologyVersions | ft -AutoSize
            Out-Host -InputObject $Allversion
            $Last = Read-Host "Please select version ('Last' number)"
        }

        $ReferenceXmls = Get-CHNLastTopologyXmls -Last $Last

        Write-Host "Getting current topology..."
        # 1.2.1 update:
        If ($SnapShot.IsPresent) { $DifferenceTopology = Get-CHNTopology -Types $Types -Export }
        Else { $DifferenceTopology = Get-CHNTopology -Types $Types }
        # 1.2.1 end

    } # end CurrentComparison

    If ($PSCmdlet.ParameterSetName -eq "FormerComparison") {

        Write-Host "Start Comparison by hash table..."

    } # end FormerComparison

    If ($PSCmdlet.ParameterSetName -eq "IndexComparison") {
        Write-Host "Start comparison by index..."
        $ReferenceXmls = Get-CHNLastTopologyXmls -Last $ReferenceIndex
        $DifferenceXmls = Get-CHNLastTopologyXmls -Last $DifferenceIndex
        
    } # end IndexComparison

    If ($PSCmdlet.ParameterSetName -eq "VersionComparison") {
        Write-Host "Start comparison by version..."
        $ReferenceXmls = Get-CHNLastTopologyXmls -Version $ReferenceVersion
        $DifferenceXmls = Get-CHNLastTopologyXmls -Version $DifferenceVersion
        
    } # end VersionComparison
    
    # Start to inspect REF and DIF version.
    # below condition will break comparison.
    # 1. Reference version cannot be found in store.
    # 2. Candidate types cannot be found from refference version.
    # 3. Difference version cannot be found in store.
    # =============================================================
    If (!$ReferenceTopology) {

        If (!$ReferenceXmls) {
            Write-Host "Reference topology version out of range!" -ForegroundColor Red
            Break
        }
        
        # Need to identify loaded xml files include Types.
        $FileNames = $ReferenceXmls.Name
        $continue = $true
        $AbsentTypes = @()
        foreach ($type in $Types) {
            $Matched = $FileNames -match $type
            If (!$Matched) {
                $continue = $false
                $AbsentTypes += $type
            }
        }

        If (!$continue) {
            Write-Host "Types " -NoNewline
            Write-Host (Convert-ArrayAsString $AbsentTypes) -ForegroundColor Red -NoNewline
            Write-Host "cannot be found in $(Convert-ArrayAsString $FileNames)"
            Break
        }

    }

    # Import topology if not did.
    If (!$DifferenceTopology) {
        If (!$DifferenceXmls) {
            Write-Host "Difference topology version out of range!" -ForegroundColor Red
            Break
        }

        # Just compare matched types between REF and DIF if they are not matched.
        # to be continue from here. Change comparison objects from xmls to versions.
        $ComparedTypes = Compare-Object $DifferenceXmls.Name $ReferenceXmls.Name -IncludeEqual
        $MismatchedTypes = $ComparedTypes | ? SideIndicator -NE "="
        $MatchedTypes = $ComparedTypes| ? SideIndicator -EQ "=" | Select -ExpandProperty InputObject
        If ($MismatchedTypes) {
            Write-Host "Reference types and difference types are mismatched!" -ForegroundColor Yellow
            Write-Host "Matched types " -NoNewline
            Write-Host "$(Convert-ArrayAsString -array $MatchedTypes) ." -ForegroundColor Green -NoNewline
            $Types = $MatchedTypes
        }

        Write-Host "Importing difference topology..."
        $DifferenceTopology = Import-CHNTopology -XmlFiles $DifferenceXmls -Types $Types
        
    }


    If (!$ReferenceTopology) {
        Write-Host "Importing reference topology..."
        $ReferenceTopology =  Import-CHNTopology -XmlFiles $ReferenceXmls -Types $Types
    }

    Write-Host "Starting to compare topology for " -NoNewline
    Write-Host (Convert-ArrayAsString -array $Types) -ForegroundColor Green

    Switch ($Types) {
        "VMs" {
            Write-Host "Comparing $_..."
            $VMItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -CompareField VMachineId
            If ($VMItemResult) {
                $Properties = @("RecordAction","DataCenter","NetworkId","FarmRole","FarmState","FarmId","PMachineId","VMachineId","IPAddress","Name","Role","State","*Version","VMImageId")
                $VMItemResult = $VMItemResult | Select-Object -Property $Properties | Sort-Object RecordAction,Role
                $MailBody += Format-CHNTopologyHtmlTable -Contents $VMItemResult -Title "VMs Changes" -ColorItems

            }
            $Results += $VMItemResult
        }

        "PMs" {
            Write-Host "Comparing $_..."
            $PMItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -GroupField Type -CompareField PMachineId
            If ($PMItemResult) { 
                $Properties = @("RecordAction","PMachineId","Name","Type","State","ZoneId","SerialNum","Model","ILoIpAddress","OperatingSystem","PhysicalVlanId","PhysicalZoneId","MacAddress","IpAddress","ImagingRoleName")
                $PMItemResult = $PMItemResult |Select-Object -Property $Properties | Sort-Object RecordAction,Type
                $MailBody += Format-CHNTopologyHtmlTable -Contents $PMItemResult -Title "PMs Changes" -ColorItems 
            }
            $Results += $PMItemResult
        }

        "Networks" {
            Write-Host "Comparing $_..."
            $NetworkItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -GroupField DataCenter -CompareField NetworkId
            If ($NetworkItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $NetworkItemResult -Title "Network Changes" -ColorItems
            }
            $Results += $NetworkItemResult
        }

        "Farms" {
            Write-Host "Comparing $_..."
            $FarmItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -GroupField NetworkId -CompareField FarmId
            If ($FarmItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $FarmItemResult -Title "Network Changes" -ColorItems
            }
            $Results += $FarmItemResult
        }

        "VLANs" {
            Write-Host "Comparing $_..."
            $VLANItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -GroupField ZoneId -CompareField VlanId
            If ($VLANItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $VLANItemResult -Title "Network Changes" -ColorItems
            }
            $Results += $VLANItemResult
        }

        "NLBs" {
            Write-Host "Comparing $_..."
            $NLBItemResult = Compare-CHNItems $ReferenceTopology[$_] $DifferenceTopology[$_] -GroupField Partition -CompareField LoadBalancerId
            If ($NLBItemResult) {
                $MailBody += Format-CHNTopologyHtmlTable -Contents $NLBItemResult -Title "Network Changes" -ColorItems
            }
            $Results += $NLBItemResult
        }

        "VMIPs" {
            # Compare properties only. Coming soon...
        }
    }

    $TypeStr = "`"$($Types -join ',')`""

    If (!$Results) {
        Write-Warning -Message "$TypeStr no any changes in version  $ReferenceXmls"
        $MailBody = "<h1 Style=`"color:Green`">No any changes!</h1>"
    }
    Else {
        Write-Host "Comaprison has been completed!" -ForegroundColor Green
    }

    If ($MailTo) {
        Write-Host "Send mail to $MailTo" -ForegroundColor Green
        $version = (Get-CHNTopologyVersions -Last $Last).Version
        # Deprecated
        # Send-TopologyMail -To $MailTo -Subject "CHN topology comparison for $TypeStr with version [$version] " -Body $MailBody
        $From = "TopologyReporter@21vianet.com"
        Send-Email -To $MailTo -mailsubject "CHN topology comparison for $TypeStr with version [$version] " -mailbody $MailBody -From $From -BodyAsHtml
    }
    Else {
        Write-Host "Didn't send result as mail! If you want, please re-run this script with parameter 'MailTo'." -ForegroundColor Yellow
    }    

    Return $Results

}

Function InspectTypes {
    Param (
        [String[]] $Types,
        [String[]] $RefTypes,
        [String[]] $DifTypes
    )

    # Compare ref and dif types for mismatch.
    $MismatchedTypes = Compare-Object -ReferenceObject $RefTypes -DifferenceObject $DifTypes
    If ($MismatchedTypes) {
        Write-Host "Refernce and difference types are mismatched!" -ForegroundColor Red
        Return $false
    }

    $AbsentTypes = $Types | ? {$_ -notin $RefTypes}
    If ($AbsentTypes) {
        Write-Host "Query types [$(Convert-ArrayAsString $AbsentTypes)] are not found in refernece topology." -ForegroundColor Red
        Return $false
    }

    Return $true
}
# CHNTopologyFunctions>

# <MyUtilities

# MyUtilities>

# <SPSHealthChecker
Function Import-HealthTemplate {

    Param (
        [Parameter(Mandatory=$true)]
        [String] $TemplatePath,
        [Parameter(Mandatory=$true)]
        [String] $Role,
        [Parameter(Mandatory=$true)]
        [String] $Type
    )

    $FileName = $Type,$Role,"Template.xml" -join "_"
    If (!(Test-Path $TemplatePath)) {
        Write-Host "Template path `"$TemplatePath`" invalid!" -ForegroundColor Red
        Return
    }

    $FileFullPath = Join-Path $TemplatePath $FileName
    If (!(Test-Path $FileFullPath)) {
        Write-Host "Template file `"$FileFullPath`" does not exist!" -ForegroundColor Red
        Return
    }

    $objects = Import-Clixml -Path $FileFullPath

    Return $objects
}

Function Get-OSService {
    
    Param (
        [Parameter(Mandatory=$true,ParameterSetName="FarmSet")]
        [Int] $Farm,
        [Parameter(ParameterSetName="FarmSet")]
        [ValidateSet("USR","BOT","DCH")]
        [String] $Role = "USR",
        [Parameter(Mandatory=$true,ParameterSetName="ComputerSet")]
        [Alias("CN","ServerName","MachineName")]
        [String[]] $ComputerName,
        [Parameter(Mandatory=$true)]
        [String[]] $ServiceName
    )

    # For farm set:
    If ($PSCmdlet.ParameterSetName -eq "FarmSet") {
        $VMs = Get-CenteralVM -Farm $Farm -Role $Role
        If (!$VMs) { Write-Host "No [$Role] VM in farm [$Farm]" -ForegroundColor Red; Return }
        $RunningVMs = $VMs | ? State -EQ "Running" | Select -ExpandProperty Name
        If (!$RunningVMs) { Write-Host "No running [$Role] VM in farm [$Farm]" -ForegroundColor Red; Return }
        $OtherStateVMs = $VMs | ? State -NE "Running"
        If ($OtherStateVMs) {
            Write-Host "Below VM(s) will not retrieve service:" -ForegroundColor Yellow
            $OtherStateVMs | Format-Table -AutoSize VMachineId, PMachineId, NetworkId, `
                                    FarmId, Name, Role, State, Version, ProductVersion, VMImageId, Heartbeat | Out-Host
        }
    }

    # For computer set:
    If ($PSCmdlet.ParameterSetName -eq "ComputerSet") {
        # Skip Centeral validation. Validation should be done before invoke this function.
        $RunningVMs = $ComputerName
    }

    $sb = {
        Param ([string[]] $ServiceName)
        
        $Services = @()
        foreach ($s in $ServiceName) {
            $Services += Get-WmiObject -Class Win32_Service -Filter "Name='$s'" 
        }

        Return $Services
    }

    # Retrieve services for every VM.
    If ($RunningVMs.Count -eq 1) {
        $Services = Get-WmiObject -Class Win32_Service -ComputerName $RunningVMs | ?{ $_.Name -in $ServiceName } | Select-Object `
                                    -Property SystemName,Name,DisplayName,Description,PathName,StartMode,StartName,State
    }
    Else {
        $Jobs = Invoke-Command -ScriptBlock $sb -ComputerName $RunningVMs -ArgumentList (,$ServiceName) -AsJob | Out-Null
        $Services = Watch-Jobs -Id $Jobs -Activity "Retrieving services ..." | Select-Object `
                                    -Property SystemName,Name,DisplayName,Description,PathName,StartMode,StartName,State
    }
                                
    Return $Services
}

Function Compare-SharePointService {
    
    Param (
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceService,
        [Parameter(Mandatory=$true)]
        [Object[]] $DifferenceService,
        [String] $Property = "State"
    )

    $Result = @()

    $GroupedDiffSrv = $DifferenceService | Group-Object Name

    foreach ($GDS in $GroupedDiffSrv) {
        $templateService = $ReferenceService | ? Name -EQ $GDS.Name
        foreach ($s in $GDS.Group) {
            If ($s.$Property -ne $templateService.$Property) { $Result += $s }
        }
    }

    # Add 'RelativeIndex' into result.
    $Result | Add-Member -NotePropertyName RelativeIndex -NotePropertyValue 0
    foreach ($r in $Result) {
        $r.RelativeIndex = $ReferenceService | ? Name -EQ $r.Name | Select-Object -ExpandProperty RelativeIndex
    }

    Return $Result

}

Function Get-RelativeOSService {
    
    Param (
        [Parameter(Mandatory=$true)]
        [String] $Name,
        [Parameter(Mandatory=$true)]
        [String] $ComputerName,
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceService
    )

    $Service = $ReferenceService | ? Name -EQ $Name
    If (!$Service) {
        Write-Host "Cannot find $Name in reference!" -ForegroundColor Yellow
        Return
    }

    # find out relative service from template.
    $RefRelativeService = $ReferenceService | ? RelativeIndex -EQ $Service.RelativeIndex | ? Name -NE $Name

    # Retrieve real service from machine.
    If ($RefRelativeService) {
        $RelativeService = Get-OSService -ComputerName $ComputerName -ServiceName $RefRelativeService.Name
    }

    Return $RelativeService
}

Function Get-MaxLengthString {
    Param (
        [Object[]] $InputObject,
        [String] $Property
    )
    $maxString = ""

    # Should be covert array in force because it will not return array if 'InputObject' just have one element.
    $Field = @($InputObject | Select-Object -ExpandProperty $Property)
    If (!$Field) { Write-Host "Cannot find property [$Property]!" -ForegroundColor Red; Return }
    If ($Field[0].GetType().Name -ne "String") {
        Write-Host "Property [$Property] is not string!" -ForegroundColor Red
        Return
    }

    foreach ($object in $Field) {
        If ($object.Length -gt $maxString.Length) {
            $maxString = $object
        }
    }

    Return $maxString
}

Function Get-PropertyMaxLength {
    Param (
        [Object[]] $InputObject,
        [String] $Property
    )

    $maxString = Get-MaxLengthString -InputObject $InputObject -Property $Property
    If ($Property.Length -gt $maxString.Length) {
        Return $Property.Length
    }
    Else { Return $maxString.Length }
}

Function Out-OSService {
    
    Param (
        [Parameter(Mandatory=$true)]
        [Object] $Service,
        [Object] $RelativeService,
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceService
    )
    Write-Host "`n===============`n"
    Write-Host "Changed Service:"
    $ComputerNameLength = Get-PropertyMaxLength -InputObject $Service -Property SystemName
    If ("ComputerName".Length -gt $ComputerNameLength) { $ComputerNameLength = "ComputerName".Length }
    $DisplayNameLength = Get-PropertyMaxLength -InputObject $Service -Property DisplayName
    If ($DisplayNameLength -gt 25) { $DisplayNameLength = 25 }
    $NameLength = Get-PropertyMaxLength -InputObject $Service -Property Name
    $StateLength = "TemplateState".Length

    $Service | Format-Table -Wrap -Property @{Name="ComputerName";Expression="SystemName";Width=$ComputerNameLength;Alignment="Left"}, 
                                      @{Name="ServiceName";Expression="Name";Width=$NameLength;Alignment="Left"}, 
                                      @{Name="DisplayName";Expression="DisplayName";Width=$DisplayNameLength;Alignment="Left"},
                                      @{Name="TemplateState";Expression={($ReferenceService | ? Name -EQ $_.Name).State};Width=$StateLength;Alignment="Left"}, 
                                      @{Name="CurrentState";Expression={$_.State};Width=$StateLength;Alignment="Left"}, 
                                      Description | Out-Host

    

    If ($RelativeService) {
    Write-Host "Relative Service:"
    $ComputerNameLength = Get-PropertyMaxLength -InputObject $RelativeService -Property SystemName
    If ("ComputerName".Length -gt $ComputerNameLength) { $ComputerNameLength = "ComputerName".Length }
    $DisplayNameLength = Get-PropertyMaxLength -InputObject $Service -Property DisplayName
    If ($DisplayNameLength -gt 25) { $DisplayNameLength = 25 }
    $NameLength = Get-PropertyMaxLength -InputObject $RelativeService -Property Name

    $RelativeService | Format-Table -Wrap -Property @{Name="ComputerName";Expression="SystemName";Width=$ComputerNameLength;Alignment="Left"}, 
                                      @{Name="ServiceName";Expression="Name";Width=$NameLength;Alignment="Left"}, 
                                      @{Name="DisplayName";Expression="DisplayName";Width=$DisplayNameLength;Alignment="Left"},
                                      @{Name="TemplateState";Expression={($ReferenceService | ? Name -EQ $_.Name).State};Width=$StateLength;Alignment="Left"}, 
                                      @{Name="CurrentState";Expression={$_.State};Width=$StateLength;Alignment="Left"}, 
                                      Description | Out-Host
    }

    
}

Function Write-MailContent {
    
    Param (
        [Parameter(Mandatory=$true)]
        [Object] $Service,
        [Object] $RelativeService,
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceService
    )

    $OutputService = $Service | select -Property @{Name="ComputerName";Expression="SystemName"}, 
                                                @{Name="ServiceName";Expression="Name"}, 
                                                DisplayName,
                                                @{Name="TemplateState";Expression={($ReferenceService | ? Name -EQ $_.Name).State}}, 
                                                @{Name="CurrentState";Expression="State"}, 
                                                Description

    $Content = Format-HtmlTable -Contents $OutputService -Title "Changed Service:"

    If ($RelativeService) {
        $OutputRService = $RelativeService | select -Property @{Name="ComputerName";Expression="SystemName"}, 
                                                @{Name="ServiceName";Expression="Name"}, 
                                                DisplayName,
                                                @{Name="TemplateState";Expression={($ReferenceService | ? Name -EQ $_.Name).State}}, 
                                                @{Name="CurrentState";Expression="State"}, 
                                                Description    
        $Content += Format-HtmlTable -Contents $OutputRService -Title "Relative Service:"
    }
    Else { $Content += "<H1>Relative Service:</H1>NA" }

    $Content += "<hr>"

    Return $Content
}
# SPSHealthChecker>

# <Invoke-ZhouBapi
Function UpdateJobStateMap {
<#
Job state map: To fire alert mail, if State not equals PreviousState
JobId, State, PreviousState
#>
	[CmdletBinding()]
	Param (
		[Object[]] $CenteralJob,
		[Object[]] $JobStateMap
	)

	Write-Verbose "Start Function [UpdateJobStateMap]"

	if ($CenteralJob.Count -gt 0 -and $JobStateMap.Count -gt 0) {
		# Compare JobId
		$DiffObject = Compare-Object $CenteralJob $JobStateMap -Property JobId -IncludeEqual

		$NewJobId = $DiffObject | ? SideIndicator -EQ "<=" | select -ExpandProperty JobId
		$FinishJobId = $DiffObject | ? SideIndicator -EQ "=>" | select -ExpandProperty JobId
		$ExistJobId = $DiffObject | ? SideIndicator -EQ "==" | select -ExpandProperty JobId

		if ($NewJobId) {
			# Create new job state map
			$NewJob = $CenteralJob | ? JobId -In $NewJobId
			foreach ($job in $NewJob) {
				$JobStateMap += CreateJobStateObject -CenteralJob $job
			}
		}

		if ($FinishJobId) {
			# set state as 'Deleted'
			$FinishStateMap = $JobStateMap | ? JobId -In $FinishJobId
			foreach ($map in $FinishStateMap) {
				$map.PreviousState = $map.State
				$map.State = 'Deleted'
			}
		}

		if ($ExistJobId) {
			# Update job state in map if it's changed.
			foreach ($jobid in $ExistJobId) {
				$job = $CenteralJob | ? JobId -EQ $jobid
				$statemap = $JobStateMap | ? JobId -EQ $jobid
				if ($job.State -ne $statemap.State) {
					$statemap.PreviousState = $statemap.State
					$statemap.State = $job.State
				}
			}
		}

	}
	elseif ($CenteralJob.Count -gt 0 -and $JobStateMap.Count -eq 0) {
		# Create Job state map
		foreach ($job in $CenteralJob) {
			$JobStateMap += CreateJobStateObject -CenteralJob $job
		}
	}
	elseif ($CenteralJob.Count -eq 0 -and $JobStateMap.Count -gt 0) {
		# All job finish. Update state as 'Deleted' in map.
		foreach ($StateMap in $JobStateMap) {
			$StateMap.PreviousState = $StateMap.State
			$StateMap.State = "Deleted"
		}
	}
	
	Write-Verbose "End Function  [UpdateJobStateMap]"
	return $JobStateMap
}

# Consider to depracte this, becuase monitor may loss some step durtion wating interval.
Function UpdateJobStepMap {

}

Function UpdateJobErrorMap {
<#
Job error map: To record job errors for each step. Just retain 10 latest errors.
JobId, Step, ErrorCount, NewErrorCount, LatestReportId, Errors
#>

	[cmdletbinding()]
	Param (
		[Object[]] $CenteralJob,
		[Object[]] $JobErrorMap
	)

	Write-Verbose "Start Function [UpdateJobErrorMap]"
	if ($CenteralJob.Count -gt 0 -and $JobErrorMap.Count -gt 0) {
		# Compare objects
		$DiffObject = Compare-Object -ReferenceObject $CenteralJob -DifferenceObject $JobErrorMap -Property JobId -IncludeEqual

		# New job appear
		$NewJobId = $DiffObject | ? SideIndicator -EQ "<=" | select -ExpandProperty JobId
		if ($NewJobId) {
			$NewJob = $CenteralJob | ?{$NewJobId -contains $_.JobId}
			# Create Error map
			foreach ($job in $NewJob) {
				# To limit max error count, we just get errors less than 500.
				$jobError = Get-CenteralJobError -Identity $job -Step $job.Step -Limit 500 -ErrorAction Ignore
                if ($jobError) {
				    $JobErrorMap += CreateJobErrorObject -CenteralJobError $jobError
                }
			}
		}

		# Job end
		$EndJobId = $DiffObject | ? SideIndicator -EQ "=>" | select -ExpandProperty JobId 
		if ($EndJobId) {
			$EndJobErrorMap = $JobErrorMap | ?{$EndJobId -contains $_.JobId}
			$EndJobErrorMap | %{$_.NewErrorCount = 0}
		}

		# exist jobs
		$ExistJobId = $DiffObject | ? SideIndicator -EQ "=" | select -ExpandProperty JobId
		if ($ExistJobId) {
			foreach ($jobid in $ExistJobId) {
				$job = $CenteralJob | ? JobId -EQ $jobid
				$jobError = Get-CenteralJobError -Identity $job -Step $job.Step -Limit 500 -ErrorAction Ignore
				$errorMap = $JobErrorMap | ? JobId -EQ $jobid
				# Update error map
				if ($errorMap.Step -ne $job.Step) {
					$errorMap.Step = $job.Step
					$errorMap.ErrorCount = $jobError.ErrorCount
				}
				else {
					$errorMap.ErrorCount += $jobError.ErrorCount
				}
				$errorMap.NewErrorCount = $jobError.ErrorCount
				$errorMap.LatestReportId = $jobError | select -First 1 -ExpandProperty ReportId
				$errorMap.Errors = $jobError | select -First 10
			}
		}

	}
	elseif ($CenteralJob.Count -gt 0 -and $JobErrorMap.Count -eq 0) {
		# Create JobErrorMap
		foreach ($job in $CenteralJob) {
			# To limit max error count, we just get errors less than 500.
			$jobError = Get-CenteralJobError -Identity $job -Step $job.Step -Limit 500 -ErrorAction Ignore
            if ($JobError) {
			    $JobErrorMap += CreateJobErrorObject -CenteralJobError $jobError
            }
		}
	}
	elseif ($CenteralJob.Count -eq 0 -and $JobErrorMap.Count -gt 0) {
		# Update NewErrorCount = 0
		$JobErrorMap | %{$_.NewErrorCount = 0}
	}

	return $JobErrorMap
}


Function ShowBasicInfo {
	Param(
		[String[]] $Type,
		[Boolean] $Wait,
		[Int[]] $JobId,
		[Int] $RefreashInterval,
		[Int] $Duration,
		[Int[]] $AlertStep,
		[Boolean] $FinishAlert,
		[Boolean] $SuspendAlert,
		[Boolean] $ErrorAlert,
		[String[]] $MailTo
	)

	$Text = @()

	if ($Type.Count -gt 0) {
		$Text += "Type: $(Convert-ArrayAsString -array $Type -Separator ',' -Quote Double)"
		$Text += "Wait: $Wait"
	}

	if ($JobId.Count -gt 0) {
		$Text += "JobId: $(Convert-ArrayAsString -array $JobId -Separator ',' -Quote None)"
	}

	$Text += "Default Inverval: $RefreashInterval seconds"
	$Text += "Duration: $($Duration / (60 * 60)) hours"
	if ($AlertStep.Count -gt 0) {
		$Text += "Alert Step: $(Convert-ArrayAsString -array $AlertStep -Separator ',' -Quote None)"
	}

	$Text += "Finish Alert: $($FinishAlert.ToString())"
	$Text += "Suspend Alert: $($SuspendAlert.ToString())"
	$Text += "Error Alert: $($ErrorAlert.ToString())"
	if (($FinishAlert -or $SuspendAlert -or $ErrorAlert -or $AlertStep.Count -gt 0) -and $MailTo.Count -gt 0) {
		$Text += "Mail To: $(Convert-ArrayAsString -array $MailTo -Separator ',' -Quote Double)"
	}

	Format-List -InputObject $Text | Out-Host

}

Function ShowCenteralJobs {
<#
Centeral job:
JobId, Type, StartTime, NextTime, State, Step, Owner, Retry, ObjectId
#>

	Param ([Object[]] $CenteralJob)

	Write-Host "Job(s):"
    if ($CenteralJob.Count -gt 0) {
        $CenteralJob | Format-Table -AutoSize -Property @{Name="JobId";Expression="JobId";Alignment="Left"}, 
                    @{Name="Type";Expression="Type";Alignment="Left"}, 
                    @{Name="ObjectId";Expression="ObjectId";Alignment="Left"}, 
                    @{Name="State";Expression="State";Alignment="Left"},
                    @{Name="Step";Expression="Step";Alignment="Left"},
                    @{Name="NextTime";Expression="NextTime";Alignment="Left"},
                    @{Name="Owner";Expression="Owner";Alignment="Left"},
                    @{Name="RetryCount";Expression="RetryCount";Alignment="Left"} | Out-Host
    }
    else {
        Write-Host "NA"
    }
	

}

Function ShowJobErrorMap {
	Param ([Object[]] $JobErrorMap)

	Write-Host "Job error map:"
	if ($JobErrorMap) {
		$JobErrorMap | Format-Table -AutoSize -Property JobId, Step,
				@{Name="IncreaseErrorCount";Expression="NewErrorCount"}, 
				@{Name="StepErrorCount";Expression="ErrorCount"} | Out-Host
	}
	else {
		Write-Host "NA"
	}
	
}

Function ShowSuspendedJob {
	Param ([Object[]] $SuspendedJob)

	Write-Host "Suspended Job(s):" -ForegroundColor Yellow
	if ($SuspendedJob.Count -gt 0) {
		$SuspendedJob | Format-Table -AutoSize -Property JobId, Type, Step, State, LockTime, StartTime | Out-Host
	}
	else {
		Write-Host "NA" -ForegroundColor Green
	}
}

Function ShowFinishedJob {
	Param ([Object[]] $FinishJob)

	Write-Host "Finish Job(s):"
	if ($FinishJob.Count -gt 0) {
		$FinishJob | Format-Table -AutoSize -Property JobId, Type, Step, State, LockTime, StartTime | Out-Host
	}
	else {
		Write-Host "NA"
	}
}


Function FireJobErrorMail {
<#
Job error map: To record job errors for each step. Just retain 10 latest errors.
JobId, Step, ErrorCount, NewErrorCount, LatestReportId, Errors
#>
	Param (
		[Parameter(Mandatory=$true)]
		[Object[]] $JobErrorMap,
		[Parameter(Mandatory=$true)]
		[String[]] $MailTo
	)

	foreach ($map in $JobErrorMap) {
		if ($map.NewErrorCount -gt 0) {
			$MailBody = $map.Errors | Out-String
			$MailSubject = "Job [$map.JobId] got new error(s)"
		}
		FireAlertMail -MailBody $MailBody -MailSubject $MailSubject -MailTo $MailTo
	}

}

Function FireSuspendJobMail {
<#
Job state map: To fire alert mail, if State not equals PreviousState
JobId, State, PreviousState
#>
	Param (
		[Parameter(Mandatory=$true)]
		[Object[]] $SuspendedJob,
		[Parameter(Mandatory=$true)]
		[String[]] $MailTo
	)

	$MailSubect = "Job [{0}] suspended!" -f (Convert-ArrayAsString -array $SuspendedJob.JobId)
	$job = $SuspendedJob | select -Property JobId, Step, StartTime, LockTime, RetryCount, State
	$MailBody = Format-HtmlTable -Contents $job -Title "Suspended Job(s):" -TitleColor $CommonColor['Yellow']

	FireAlertMail -MailBody $MailBody -MailSubject $MailSubject -MailTo $MailTo -BodyAsHtml

}

Function FireFinishJobMail {
	Param (
		[Parameter(Mandatory=$true)]
		[Object[]] $FinishJob,
		[Parameter(Mandatory=$true)]
		[String[]] $MailTo
	)

	$MailSubject = "Job [{0}] finished!" -f (Convert-ArrayAsString -array $FinishJob.JobId)
	$job = $FinishJob | select -Property JobId, Step, StartTime, LockTime, RetryCount, State
	$MailBody = Format-HtmlTable -Contents $job -Title "Finished Job(s):" -TitleColor $CommonColor['Yellow']

	FireAlertMail -MailBody $MailBody -MailSubject $MailSubject -MailTo $MailTo -BodyAsHtml

}

Function FireAlertMail {
	Param (
		[Parameter(Mandatory=$true)]
		[String[]] $MailBody,
		[Parameter(Mandatory=$true)]
		[String] $MailSubject,
		[Parameter(Mandatory=$true)]
		[String[]] $MailTo,
		[Switch] $BodyAsHtml
	)

	Send-Email -To $MailTo -mailsubject $MailSubject -mailbody $MailBody -From "Zhoubapi_SPO@21vianet.com" -BodyAsHtml:$($BodyAsHtml.IsPresent)

}


Function CalculateSleepInterval {
<#
	To calculate sleeping interval accroding to current job is running and the next run time.
#>
	Param (
		[Parameter(Mandatory=$true)]
		[Object[]] $CenteralJob,
		[Parameter(Mandatory=$true)]
		[Int] $DefaultInterval
	)
	$RunningCenteralJob = $CenteralJob | ? Owner -NE ""
	if ($RunningCenteralJob) {
		$Interval = $DefaultInterval
	}
	else {
		$SleepingCenteralJob = $CenteralJob | ? Owner -EQ ""
	
		if ($SleepingCenteralJob) {
			$NextTime = $SleepingCenteralJob[0].NextTime
			foreach ($job in $SleepingCenteralJob) {
				if ($job.NextTime -lt $NextTime) {
					$NextTime = $job.NextTime
				}
			}
		}
		$Interval = ($NextTime - (Get-Date) | select -ExpandProperty TotalSeconds) + $DefaultInterval -as [Int]
	}

	return $Interval
}


Function RefreshEscapeTime {

	Param (
		[Parameter(Mandatory=$true)]
		[DateTime] $StartTime
	)

	return ((Get-Date) - $StartTime | select -ExpandProperty TotalSeconds) -as [Int]
}

Function CreateJobStateObject {
	[CmdletBinding()]
	Param ([Object] $CenteralJob)

	Write-Verbose "Start Function [CreateJobStateObject]"
	$Property = @{
		JobId = $CenteralJob.JobId
		State = $CenteralJob.State
		PreviousState = ""
	}

	Write-Verbose "End Function [CreateJobStateObject]"
	return New-Object -TypeName PSObject -Property $Property
}

Function CreateJobErrorObject {
	[CmdletBinding()]
	Param ([Object[]] $CenteralJobError)

	Write-Verbose "Start Function [CreateJobErrorObject]"
	$jobError = $CenteralJobError | select -First 1
	$Property = @{
		JobId =  $jobError.JobId
		Step = $jobError.Step
		LatestErrorId = $jobError.ReportId
		ErrorCount = $CenteralJobError.Count
		NewErrorCount = $CenteralJobError.Count
		Errors = @($CenteralJobError | select -First 10)
	}

	Write-Verbose "End Function [CreateJobErrorObject]"
	return New-Object -TypeName PSObject -Property $Property
}

#>