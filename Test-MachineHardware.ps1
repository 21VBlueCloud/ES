<# 
   此脚本可以实现对域环境的多台服务器，通过使用hpssacli.exe抓取HP 服务器的ILO 日志信息，来分析判断服务器的物理健康情况。
   此脚本可以对多台服务器使用，从而节省运维时间。
#>

Param
(
    [Parameter(Position=0,Mandatory=$False, HelpMessage='Servername or name filter to run script against.')]
    [string]$ServerName,

    [Parameter(Position=1,Mandatory=$false, HelpMessage='Determines how many previous hours of IML logs to return entries for.')]    
    [int] $RecentHours = 24,

    [Parameter(Position=2,Mandatory=$False, HelpMessage='CA Provisioning state to run against.')]
    [string]$ProvisioningState,

    [Parameter(Position=3,Mandatory=$False, HelpMessage='Throw an error if failures are found. Useful if script is running as a CA Operation')]
    [switch]$ThrowErrorOnFailures,

    [Parameter(Position=4,Mandatory=$False, HelpMessage='Optional filename to save an HTML formatted table for easier viewing')]
    [string]$HtmlOutputFilename,

    [Parameter(Position=5,Mandatory=$False, HelpMessage='Send an email')]
    [switch]$SendEmail,

    [Parameter(Position=6,Mandatory=$False, HelpMessage='Email recipients')]
    [string]$Recipients,

    [Parameter(Position=7,Mandatory=$False, HelpMessage='Include CSV result')]
    [switch]$AttachCSV,

    [Parameter(Position=8,Mandatory=$False, HelpMessage='Machine definition regex to run script against in local site.')]
    [string]$ServerDefinition 
)


if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
}
else
{
    $Script:basePath = "."
}




# Reference to the datacenter ILO library and deployment library  
$RoleDatacenterPath = "C:\LocalFiles\Exchange\Datacenter"
. (Join-Path $RoleDatacenterPath 'FFODatacenterDeploymentCommonLibrary.ps1')
. (Join-Path $RoleDatacenterPath 'DatacenterDiskCommonLibrary.ps1')
. (Join-Path $RoleDatacenterPath 'FfoDatacenterIloLibrary.ps1')
. D:\Operation\Tools\Library\LogHelper.ps1
. D:\Operation\Tools\Library\CommonHelper.ps1
. D:\Operation\Tools\Library\RecordHelper.ps1
# Increase MaximumFunctionCount from 4096 to 8192 to avoid function loading failure
$MaximumFunctionCount  = 8192

# Load Management Session
Load-ManagementShell
# Choose which machines to target
if ($ProvisioningState -and $Servername) 
{
    $mailBody = "Machines with recent ILO IML entries or failures in HP Array config found on servers matching '$Servername' in '$ProvisioningState' State.`r`n`r`n"
    Write-POELog "Using $env:ComputerName to check machine ILO IML entries and failures in HP Array config on servers matching '$Servername' in '$ProvisioningState' State"
    $allMachines = Get-CentralAdminMachine -Filter "name -like '$Servername' -and ProvisioningState -like '$ProvisioningState'"
}
elseif ($ProvisioningState)
{
    $mailBody = "Machines with recent ILO IML entries or failures in HP Array config found on servers in '$ProvisioningState' State.`r`n`r`n"
    Write-POELog "Using $env:ComputerName to check machine ILO IML entries and failures in HP Array config on servers in '$ProvisioningState' State"
    $allMachines = Get-CentralAdminMachine -Filter "ProvisioningState -like '$ProvisioningState'"
}
elseif ($Servername)
{
    $mailBody = "Machines with recent ILO IML entries or failures in HP Array config found on servers matching '$Servername'.`r`n`r`n"
    Write-POELog "Using $env:ComputerName to check machine ILO IML entries and failures in HP Array config on servers matching '$Servername'"
    $allMachines = Get-CentralAdminMachine -Identity $Servername 
}
elseif ($ServerDefinition)
{
    $mailBody = "Machines with recent ILO IML entries or failures in HP Array config found on servers matching '$ServerNameRegex'.`r`n`r`n"
    Write-POELog "Using $env:ComputerName to check machine ILO IML entries and failures in HP Array config on servers matching '$ServerNameRegex'"
    $allMachines = Get-CentralAdminMachine | Where{ $_.actualmachinedefinition -match $ServerDefinition } 
}
else
{
    $localLocation = Get-CentralAdminLocation -Local | Select -First 1
    $mailBody = "Machines with recent ILO IML entries or failures in HP Array config found in location:'$($localLocation.Name)'.`r`n`r`n"
    Write-POELog "Using $env:ComputerName to check machine ILO IML entries and failures in HP Array config on all servers in location:'$($localLocation.Name)'."
    $allMachines = Get-CentralAdminMachine -Location $localLocation.Name
}

$allMachines = $allMachines | where {$_.DigiIp -ne '::' -and $_.DigiIP -ne "1.1.1.1"}

if ($allMachines.Count -eq 0)
{
    return "No servers matched the specified criteria" 
}

#Get Forest Admin credentials
$ForestAdminCredentials = @{}
$forestMachines = $allMachines | group Forest
foreach ($forestMachine in $forestMachines)
{
    $Username = Get-MachineSingleParameter -Machine ($forestMachine.Group | select -first 1).Name -ParameterDefinition testAdminaccount
    $Password = Get-MachineSingleParameter -Machine ($forestMachine.Group | select -first 1).Name -ParameterDefinition testAdminPWD
    $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ("$($forestMachine.Name)\$Username", $SecurePassword)
    $ForestAdminCredentials.Add($forestMachine.Name, $Credential)
}

# initialize variables
$ServerErrors = @()
$machinesWithIloConnectionFailures = @{}
$machinesWithIloEntries = @{}
$machinesWithHpArrayFailures = @{}

foreach ($machine in $allMachines)
{
    Write-POELog "Testing $($machine.Name)..."
    #Run a simple login to the server ILO to determine whether it is online and the credentials in CA are correct
    try
    {
        $iloTest = Test-Ilo -MachineName $machine.Name 
    }
    finally
    {
        if ($iloTest[0] -eq $False) 
        {
            $machinesWithIloConnectionFailures.Add($machine.Name, $iloTest[1])
            $iloTestResult = $iloTest[1]
            $iloImlResult = "Unable to Test"
        }
        else
        { 
            $iloTestResult = "Successful"
            # ILO login is successful so pull the IML log and return any events indicating hardware issues
            $currentMachineIloEntries = Get-IloIMLIssues -MachineName $machine.Name -RecentHours $RecentHours -ErrorAction SilentlyContinue
            if ($currentMachineIloEntries)
            {
                Write-POELog "Machine $($machine.Name) has ILO entries"
                $iloImlText = $currentMachineIloEntries | Out-String
                $machinesWithIloEntries.Add($machine.Name, $iloImlText)
                $iloImlResult = $iloImlText
            }
            else
            {
                Write-POELog "Machine $($machine.Name) has no ILO entries"
                $iloImlResult = "No IML Entries Found"
            }
        }
    }
    # Pull the HP array config from the server and report if any failures are found
    try
    {
        # use HP SSA if its HP Gen9
        if ($global:nextGenHpSku -contains $(Get-CentralAdminMachine -Identity $machine).SKU)
        {
            $hpacuCmd = "C:\LocalFiles\HPTools\HPSSA\bin\hpssacli.exe"
        }
        else
        {
            $hpacuCmd = "C:\LocalFiles\HPTools\HPACUCLI\bin\hpacucli.exe"
        }
        # This isn't vulnerable to PowerShell injection or isn't interesting to exploit. Justification: No User Input
        $scriptblock = [scriptblock]::create("If (Test-Path $hpacuCmd) {$hpacuCmd ctrl all show config}")
        $hpconfigerrors = Invoke-Command -ScriptBlock $scriptblock -ComputerName $machine.Name -Credential $ForestAdminCredentials[$machine.Forest] -ErrorAction SilentlyContinue | select-string "fail" | select-string -notmatch "Activate on drive failure" | Out-String
    }
    finally
    {
        if ($hpconfigerrors)
        {
            Write-POELog "Machine $($machine.Name) has HP array config failures"
            $machinesWithHpArrayFailures.Add($machine.Name.ToString(), $hpconfigerrors)
            $hpArrayResult = $hpconfigerrors
        }
        else
        {
            Write-POELog "Machine $($machine.Name) has no HP array config failures"
            $hpArrayResult = "No HP Array Config Issues Found"
        }
    }

    if (($iloTestResult -ne "Successful") -or ($iloImlResult -ne "No IML Entries Found") -or ($hpArrayResult -ne "No HP Array Config Issues Found"))
    {
        $ServerErrorProperties = @{
                                      'Server Name' = $machine.Name;
                                      'SerialNumber' = $machine.SerialNumber;
                                      'AssetNumber'=$machine.AssetNumber;
                                      'Rack'= $machine.Rack;
                                      'DigiIP'=$machine.DigiIP;
                                      'ILO Test' = $iloTestResult;
                                      'IML Entries' = $iloImlResult;
                                      'HP Array Config Failures' = $hpArrayResult            
                                  }
        $ServerError = New-Object PSObject -Property $ServerErrorProperties
        $ServerErrors += $ServerError
    }
}

#If any errors were found, add them to the mailbody variable
if ($machinesWithIloConnectionFailures.Count -gt 0) 
{
    # We have at least one machine where we cannot log into the ILO
    $mailBody += "`r`nMachines with ILO login issues:`r`n"
    $machinesWithIloConnectionFailures.GetEnumerator() | ForEach-Object {
                                                                            Write-POELog "Machine with ILO login issues: $_.Key"
                                                                            Write-POELog $_.Value
                                                                            $mailBody += [string]::Concat("`r`n", "MachineName: ", $_.Key, "`r`n")
                                                                            $mailBody += [string]::Concat("`r`n", $_.Value, "`r`n")
                                                                        }
}
else
{
    Write-POELog "No machines tested have ILO login issues."   
}

if ($machinesWithIloEntries.Count -gt 0) 
{
    # We have at least one machine which is has recent ILO entries
    $mailBody += "`r`nMachines with recent ILO IML entries:`r`n"
    $machinesWithIloEntries.GetEnumerator() | ForEach-Object {
                                                                 Write-POELog "Machine with recent ILO IML entries: $_.Key"
                                                                 Write-POELog $_.Value
                                                                 $mailBody += [string]::Concat("`r`n", "MachineName: ", $_.Key, "`r`n")
                                                                 $mailBody += [string]::Concat("`r`n", $_.Value, "`r`n")
                                                             }
}
else
{
    Write-POELog "No machines tested have recent ILO IML entries."   
}

if ($machinesWithHpArrayFailures.Count -gt 0) 
{
    # We have at least one machine which is has HP array config failures
    $mailBody += "`r`nMachines with HP array config failures:`r`n"
    $machinesWithHpArrayFailures.GetEnumerator() | ForEach-Object {
                                                                      Write-POELog "Machine with HP array config failures: $_.Key"
                                                                      Write-POELog $_.Value
                                                                      $mailBody += [string]::Concat("`r`n", "MachineName: ", $_.Key, "`r`n")
                                                                      $mailBody += [string]::Concat("`r`n", $_.Value, "`r`n")
                                                                  }   
}
else
{
    Write-POELog "No machines tested have HP array config failures."   
}

# If any errors are found, write them to output and terminate the workflow so monitoring probe can pick up and fire escalation if -ThrowErrorOnFailures switch was set
if (($machinesWithIloEntries.Count -gt 0) -or ($machinesWithHpArrayFailures.Count -gt 0) -or ($machinesWithIloConnectionFailures.Count -gt 0))
{
    Write-Output $mailBody

    #Create html header for table formatting
    $htmlhead = "<style>"
    $htmlhead += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $htmlhead += "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
    $htmlhead += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;}"
    $htmlhead += "</style>"
    #Create HTML output
    $html = $ServerErrors | select 'Server Name', 'SerialNumber','AssetNumber','Rack','DigiIP','ILO Test', 'IML Entries', 'HP Array Config Failures' | ConvertTo-HTML -head $htmlhead
    #Preserve original text formatting
    $html = $html.replace("<td>", "<td><pre>")
    $html = $html.replace("</td>", "</pre></td>")
    
    if ($HtmlOutputFilename)
    {
        $html | Out-File $HtmlOutputFilename
    }

    if ($SendEmail)
    {
        
      
         $Recipients = Get-MachineSingleParameter -Machine $env:COMPUTERNAME -ParameterDefinition FfoHardwareAlertEmail
         Send-Mail -option 6 -Subject "[POE]Hardware Test Results (test.gbl)" -Body $html 
    }

    if ($ThrowErrorOnFailures)
    {
        $RecurrentLogToken = "[FFO-MANAGED-ERROR] "
        Write-Host "$RecurrentLogToken $mailBody"
        throw "$RecurrentLogToken - Issues have been found, detail can be found in the Log Trace."  
    }
}
