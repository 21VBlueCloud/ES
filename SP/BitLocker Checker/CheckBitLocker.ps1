<#
.Synopsis
This script is used to check bitlocker services all relative key points.

.Description
This script will query all key points for drive V Bitlocker encryption and check Hyper-V
service status.

It's better to save result into variable. Because this script will run a bit longer.

Note: The parameter Credential is different with common. This credential just used to query
DSE's PMs. It will not take effect, if the input PM array not contains DSE's PM.

Script can accept pipe input by zone or PM name. Please refer in exmaples.
It's querying in batch mode by PowerShell job. It will return PS object array as result. 
The returned object include below customed properties:

    ComputerName: The physical machine name.
    hasDriveV: Return true if drive V can be access.
    BitlockerExist: Return true if Bitlocker volume drive V exists.
    LockStatus: Indicate the drive V lock status. Good status is "unlocked".
    ProtectionStatus: Indicate the drive V protection status. Good status is "On".
    KeyProtector: Show drive V key protector.
    CertExist: Return true if certificate 'bitlocker-Gallatin.microsoft.com' exists on this machine.
    Cert: Return Gallatin bitlocker certificate.
    KeyName: Return temprary key name for bitlocker encrpytion.
    UnlockServiceExist: Indicate the bitlocker unlock service exists on this machine.
    UnlockServiceVersion: Bitlocker unlock service version.
    UnlockServiceState: Bitlocker unlock service state.
    UnlockServiceFileExist: The Bitlocker unlock service bit file with full path.
    V35Exist: Indicate Net framework 3.5 exists on this machine. It just shown if unlock service version is 1.
    V4Exist: Indicate Net framework 4.7 exists on this machine and version >= 461808 . It just shown if unlock 
             service version is 2.
    HyperVServiceExist: Indicate the service 'vmms' exists on this machine.
    HyperVServiceState: The 'vmms' state.
    HyperVServiceDependOn: Show 'vmms' depending service. Normally it should include bitlocker unlock service.


.Parameter ComputerName
PMachine name array. It's accept pipeline input.
Alias are "Cn" and "MachineName".

.Parameter ZoneId
Zone Id. It's accept pipeline input by property name.

.Parameter CheckType
Accept array input including "TestDriveV", "CheckKeyProtector", "CheckCert", "CheckUnlockService" and "CheckHyperVService".
Alias is "Type".


.Parameter Credential
Your ada account credential. It's just for DSE's PM querying.
It will not take effect if the PM list not contains DSE's PM.

.Example
$result = 6,7 | CheckBitLocker.ps1 -Credential $ada_cred
This example is querying all checking types for all PMs in zone 6 and 7 both.

.Example
$result = $PMName | CheckBitLocker.ps1 -Credential $ada_cred -Type "TestDriveV", "CheckKeyProtector"
This example is querying checking types "TestDriveV", "CheckKeyProtector" for PMs stored in variable PMName.

.Outputs
PSObject[]

#>

param (
    
    [Parameter(Mandatory=$true, ParameterSetName="cn",
        ValueFromPipeline=$true)]
    [Alias("Cn", "MachineName")]
    [String[]]$ComputerName,
    [Parameter(Mandatory=$true, ParameterSetName="zone",
        ValueFromPipelineByPropertyName=$true,
        ValueFromPipeline=$true)]
    [Int]$ZoneId,
    [ValidateSet("TestDriveV", "CheckKeyProtector", "CheckCert", "CheckUnlockService", "CheckHyperVService")]
    [Alias("Type")]
    [String[]]$CheckType = @("TestDriveV", "CheckKeyProtector", "CheckCert", "CheckUnlockService", "CheckHyperVService"),
    [PSCredential]$Credential
)

begin {

$PMachines = @()
$CheckResult = @()
$Properties = @()

#region scriptblock

$TestDriveV = {
    $hasV = Test-Path V:\
    return New-Object -TypeName PSObject -Property @{hasDriveV=$hasV; ComputerName=$env:COMPUTERNAME}
}

$CheckKeyProtector = {
    $v = Get-BitlockerVolume V -ErrorAction Ignore
    if ($v) {
        return New-Object -TypeName PSObject -Property @{
            Exist = $true
            LockStatus=$v.LockStatus.ToString(); 
            ProtectionStatus=$v.ProtectionStatus.ToString(); 
            KeyProtector=$v.KeyProtector;
            ComputerName=$env:COMPUTERNAME
        }
    }
    else {
        return New-Object -TypeName PSObject -Property @{
            Exist = $false
            LockStatus=$null 
            ProtectionStatus=$null
            KeyProtector=$null
            ComputerName=$env:COMPUTERNAME
        }
           
    }
}

$CheckCert = {
    $Cert = Get-Item Cert:\LocalMachine\My\BD8F9B5FDD339C0383B35CE89F0745720533DC74 -ErrorAction Ignore
    if ($Cert) {

        #Add check cert key value from registry.
        $key_path = 'HKLM:\SOFTWARE\Policies\Microsoft\FVE'
        $ip = (Get-ItemProperty -Path $key_path -ErrorAction Ignore)

        if ($ip) {
            $KeyName =  @($ip | gm | ? Name -match '^v2' | select -ExpandProperty Name)
        }
        return New-Object -TypeName PSObject -Property @{Exist=$true; Cert=$Cert; KeyName=$KeyName; ComputerName=$env:COMPUTERNAME}
    }
    else { return New-Object -TypeName PSObject -Property @{Exist=$false; Cert=$null; KeyName=$null; ComputerName=$env:COMPUTERNAME} }
}

$CheckUnlockService = {
    
    $Version = 2
    $Service = Get-WmiObject -Query "select * from Win32_service where Name='BitLockerUnlockService'"

    if (!$Service) {
        $Service = Get-WmiObject -Query "select * from Win32_service where Name='BitLockerUnlockingService'"
        $Version = 1
    }

    if ($Service) {
        $PathName = $Service.PathName
        $FileExisted = Test-Path $PathName
        return New-Object -TypeName PSObject -Property @{
            Exist = $true
            Version = $Version
            Name = $Service.Name
            State = $Service.State.ToString()
            FilePath = $PathName
            FileExisted = $FileExisted
            ComputerName = $env:COMPUTERNAME
        }
    }
    else {
        return New-Object -TypeName PSObject -Property @{
            Exist = $false
            Version = $null
            Name = $null
            State = $null
            FilePath = $null
            FileExisted = $null
            ComputerName = $env:COMPUTERNAME
        }
    }
}

$CheckHyperVService = {
    $Service = Get-Service vmms -ErrorAction Ignore

    if ($Service) {
        $KeyProperty = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\VMMS"
        return New-Object -TypeName PSObject -Property @{
            Exist = $true
            Name = $Service.Name
            State = $Service.Status.ToString()
            DependOnService = $KeyProperty.DependOnService
            ComputerName = $env:COMPUTERNAME
        }
    }
    else {
        return New-Object -TypeName PSObject -Property @{
            Exist = $false
            Name = $null
            State = $null
            DependOnService = $null
            ComputerName = $env:COMPUTERNAME
        }
    }
}

$CheckNetFrameworkV4 = {
    $V4 = Get-ItemProperty -Path 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction Ignore
    if ($V4 -and $V4.Release -ge 461808) {
        return New-Object -TypeName PSObject -Property @{
            Exist = $true
            ReleaseVerion = $V4.Release
            ComputerName = $env:COMPUTERNAME
        }
    }
    else {
        return New-Object -TypeName PSObject -Property @{
            Exist = $false
            ReleaseVerion = $null
            ComputerName = $env:COMPUTERNAME
        }
    }
}

$CheckNetFrameworkV35 = {

    $V35 = Get-ItemProperty -Path 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5' -ErrorAction Ignore

    if ($V35) {
        return New-Object -TypeName PSObject -Property @{
            Exist = $V35.Installed
            ComputerName = $env:COMPUTERNAME
        }
    }
    else {
        return New-Object -TypeName PSObject -Property @{
            Exist = $false
            ReleaseVerion = $null
            ComputerName = $env:COMPUTERNAME
        }
    }
    
}


#endregion

#region Nested function
function ExecuteTesting {

    param (
        [String[]]$PMachines,
        [PSCredential]$Credential
    )

    Write-Host ("`nPM count: {0}" -f $PMachines.Count) -ForegroundColor Green
    

    # Test the PMs are pingable.
    $BadPMs = @(Test-MachineConnectivity -ComputerName $PMachines -ReturnBad)

    if ($BadPMs.Count -gt 0) {
        Write-Host ("PMs: {0} cannot be reachable.`n" -f (Convert-ArrayAsString $BadPMs)) -ForegroundColor Red
    }

    $PMachines = @($PMachines | Where-Object { $BadPMs -notcontains $_ })

    # Starting to check
    if ($PMachines.Count -gt 0) {
    
        $Properties = @()
        foreach ($pm in $PMachines) {
            $Properties += @{ComputerName=$pm}
        }

        #region step 1: Test drive V.
    
        if ($CheckType -contains "TestDriveV") {

            Write-Host "Starting to test drive V ... " -NoNewline

            if (!$Credential) {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $TestDriveV -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Testing drive V ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $TestDriveV -ComputerName $PMachines
                }
            }
            else {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $TestDriveV -ComputerName $PMachines -AsJob -Credential $Credential | Watch-Jobs -Activity "Testing drive V ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $TestDriveV -ComputerName $PMachines -Credential $Credential
                }
            }

            foreach ($r in $Result) {
                
                $Properties | Where-Object {$_.ComputerName -eq $r.ComputerName} | foreach-object {
                    $_.Add("hasDriveV", $r.hasDriveV)
                    
                }

            }

            Write-Host "Done" -ForegroundColor Green

        }
    
        #endregion


        #region step 2: Check lock state.

        if ($CheckType -contains "CheckKeyProtector") {

            Write-Host "Starting to check key protector ... " -NoNewline

            if (!$Credential) {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckKeyProtector -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking Key Protector ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckKeyProtector -ComputerName $PMachines
                }
            }
            else {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckKeyProtector -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking Key Protector ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckKeyProtector -ComputerName $PMachines -Credential $Credential
                }
            }

            foreach ($r in $Result) {

                $Properties | Where-Object {$_.ComputerName -eq $r.ComputerName} | foreach-object {
                    $_.Add("BitlockerExist", $r.Exist)
                    $_.Add("LockStatus", $r.LockStatus)
                    $_.Add("ProtectionStatus", $r.ProtectionStatus)
                    $_.Add("KeyProtector", $r.KeyProtector)
                }
                
            }

            Write-Host "Done" -ForegroundColor Green

        }
     

        #endregion


        #region step 3: Check cert.

        if ($CheckType -contains "CheckCert") {

            Write-Host "Starting to check certificate ... " -NoNewline

            if (!$Credential) {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckCert -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking certificate ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckCert -ComputerName $PMachines
                }
            }
            else {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckCert -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking certificate ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckCert -ComputerName $PMachines -Credential $Credential
                }
            }

            foreach ($r in $Result) {

                $Properties | Where-Object {$_.ComputerName -eq $r.ComputerName} | foreach-object {
                    $_.Add("CertExist", $r.Exist)
                    $_.Add("Cert", $r.Cert)
                    $_.Add("KeyName", $r.KeyName)
                }
            }

            Write-Host "Done" -ForegroundColor Green

        }

        #endregion


        #region step 4: Cert key value in registry.

        # Merged into "check cert".


        #endregion


        #region Step 5: Check bitlocker unlocking service 

        # Including framework 3.5 & 4 checking.

        if ($CheckType -contains "CheckUnlockService") {

            Write-Host "Starting to check unlock service ... " -NoNewline

            if (!$Credential) {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckUnlockService -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking Unlock Service ..."
                    $V35Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV35 -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking Net framework V3.5 ..."
                    $V4Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV4 -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking Net framework V4 ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckUnlockService -ComputerName $PMachines
                    $V35Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV35 -ComputerName $PMachines
                    $V4Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV4 -ComputerName $PMachines
                }
            }
            else {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckUnlockService -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking Unlock Service ..."
                    $V35Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV35 -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking Net framework V3.5 ..."
                    $V4Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV4 -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking Net framework V4 ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckUnlockService -ComputerName $PMachines -Credential $Credential
                    $V35Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV35 -ComputerName $PMachines -Credential $Credential
                    $V4Result = Invoke-Command -ScriptBlock $CheckNetFrameworkV4 -ComputerName $PMachines -Credential $Credential
                }
            }

            foreach ($r in $Result) {

                $Properties | Where-Object {$_.ComputerName -eq $r.ComputerName} | foreach-object {
                    $_.Add("UnlockServiceExist", $r.Exist)
                    $_.Add("UnlockServiceVersion", $r.Version)
                    $_.Add("UnlockServiceState", $r.State)
                    $_.Add("UnlockServiceFileExist", $r.FileExisted)
                    if ($r.Version -eq 1) { 
                        $V35Exist = $V35Result | Where-Object ComputerName -eq $r.ComputerName | Select-Object -ExpandProperty Exist
                        $_.Add("V35Exist", $V35Exist) 
                    }
                    elseif ($r.Version -eq 2) { 
                        $V4Exist = $V4Result | Where-Object ComputerName -eq $r.ComputerName | Select-Object -ExpandProperty Exist
                        $_.Add("V4Exist", $V4Exist) 
                    }
                }
            }

            

            Write-Host "Done" -ForegroundColor Green

        }

        #endregion


        #region step 6: check service 'vmms'

        if ($CheckType -contains "CheckHyperVService") {

            Write-Host "Starting to check Hyper-V service ... " -NoNewline

            if (!$Credential) {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckHyperVService -ComputerName $PMachines -AsJob | Watch-Jobs -Activity "Checking Hyper-V Service ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckHyperVService -ComputerName $PMachines
                }
            }
            else {
                if ($PMachines.Count -gt 1) {
                    $Result = Invoke-Command -ScriptBlock $CheckHyperVService -ComputerName $PMachines -Credential $Credential -AsJob | Watch-Jobs -Activity "Checking Hyper-V Service ..."
                }
                else {
                    $Result = Invoke-Command -ScriptBlock $CheckHyperVService -ComputerName $PMachines -Credential $Credential
                }
            }

            foreach ($r in $Result) {

                $Properties | Where-Object {$_.ComputerName -eq $r.ComputerName} | foreach-object {
                    $_.Add("HyperVServiceExist", $r.Exist)
                    $_.Add("HyperVServiceState", $r.State)
                    $_.Add("HyperVServiceDependOn", $r.DependOnService)
                }
            }

            Write-Host "Done" -ForegroundColor Green

        }

        #endregion

    }
    else {
        Write-Host "No any PM is reachable." -ForegroundColor Red
    }

    return $Properties

}

function GetDSEPMs {
    # query DSE's PMs from AD.
    $OU = "OU=VUTDCs,OU=pMachines,OU=Servers,DC=chn,DC=sponetwork,DC=com"
    $ADPMs = Get-ADComputer -SearchBase $OU -Filter *
    return $ADPMs.Name
}

function FindOutDSEPM {
# to find out DSE's PM from list.
    param ([String[]]$PMachines)

    $Result = @()
    $DSEPMs = GetDSEPMs
    foreach ($pm in $PMachines) {
        if ($pm -in $DSEPMs) { 
            $Result += $pm 
            if ($DSEPMs.Count -eq $Result.Count) { break }
        }
    }

    return $Result
}
#endregion

} # end begin

process {

    if ($PSCmdlet.ParameterSetName -eq "zone") {
        $PMachines = Get-CenteralPMachine -Zone $ZoneId -ExcludeType ChassisManager | Select-Object -ExpandProperty Name
        Write-Host "`n`nStarting to check PMs for zone [$ZoneId]." -ForegroundColor Green
        Write-Host ("Check Type: [{0}]" -f (Convert-ArrayAsString $CheckType)) -ForegroundColor Green

        # To find out DSE's PM and ask credential if input is not provide.
        $DSEPMs = FindOutDSEPM -PMachines $PMachines
        if ($DSEPMs.Count -gt 0) {
            
            if (!$Credential) {
                Write-Host "PM list include DSE's PM. Please provide ada account credential!" -ForegroundColor Yellow
                $Credential = Get-Credential -Message "Your ada account:"
            }

            $Properties += ExecuteTesting -PMachines $DSEPMs -Credential $Credential

            $PMachines = $PMachines | Where-Object {$DSEPMs -notcontains $_}
        }
       
        $Properties += ExecuteTesting -PMachines $PMachines
    }

    if ($PSCmdlet.ParameterSetName -eq "cn") {
        $PMachines += $ComputerName
    }

} # end process

end {

    if ($PSCmdlet.ParameterSetName -eq "cn") {
        Write-Host "`n`nStarting to check PM." -ForegroundColor Green
        Write-Host ("Check Type: [{0}]" -f (Convert-ArrayAsString $CheckType)) -ForegroundColor Green

        # To find out DSE's PM and ask credential if input is not provide.
        $DSEPMs = FindOutDSEPM -PMachines $PMachines
        if ($DSEPMs.Count -gt 0) {
            
            if (!$Credential) {
                Write-Host "PM list include DSE's PM. Please provide ada account credential!" -ForegroundColor Yellow
                $Credential = Get-Credential -Message "Your ada account:"
            }

            Write-Host "DSE's PM checking:" -ForegroundColor Green 
            $Properties += ExecuteTesting -PMachines $DSEPMs -Credential $Credential

            $PMachines = $PMachines | Where-Object {$DSEPMs -notcontains $_}
        }

        Write-Host "PMs checking:" -ForegroundColor Green 
        $Properties += ExecuteTesting -PMachines $PMachines
    }

    # Make objects for return
    if ($Properties.Count -gt 0) {
        foreach ($Property in $Properties) {
            $CheckResult += New-Object -TypeName PSObject -Property $Property
        }
    }

    return $CheckResult

} # end end


