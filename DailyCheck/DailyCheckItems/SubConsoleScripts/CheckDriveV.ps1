#####################################################################################
#
#File: CheckDriveV.ps1
#Author: Wind Song
#Version: 1.0
#
##  Revision History:
##  Date          Version    Alias       Reason for change
##  --------      -------   --------    ---------------------------------------
##  10/30/2019    1.0        Wind        This script is using to check all drive V
##                                       on PMs to see Bitlocker is working or not.
#####################################################################################


Param (
    [Parameter(ParameterSetName="Computer")]
    [Alias("CN", "PMachine", "PM")]
    [String[]]$ComputerName,
    [Parameter(ParameterSetName="Zone")]
    [Int[]]$ZoneId,
    [PSCredential]$Credential,
    [Switch]$FormatHTML
)

# Initialize HTML table
$HtmlBody = "<TABLE Class='Cookie Age' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
$TableHeader = "SNC Cookie Age Check"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"

# Main code
# =====================================================================================================================

$ScriptBlock = {
    $Property = @{ComputerName=$env:COMPUTERNAME}
    $Volume = Get-Volume V -ErrorAction Ignore
    if ($Volume) { 
        $Property.Add("VolumeV", $true)
        $Property.Add("TestPathV", (Test-Path V:\))
        $BitlockerV = Get-BitlockerVolume V
        $Property.Add("LockStatus", $BitlockerV.LockStatus)
        $Property.Add("ProtectionStatus", $BitlockerV.ProtectionStatus)
    }
    else { 
        $Property.Add("VolumeV", $false)
        $Property.Add("TestPathV", $false)
        $Property.Add("LockStatus", $null)
        $Property.Add("ProtectionStatus", $null)
    }

    return New-Object -TypeName PSObject -Property $Property

}

if ($ComputerName.Count -eq 0 -and $ZoneId.Count -eq 0) {
    # To get all PMs from Centeral includeing "Running" and "Probation".
    $Zone = Get-CenteralZone -State Active
    $PMs = $Zone | ForEach-Object {Get-CenteralPMachine -Zone $_ -ExcludeType ChassisManager} | `
            Where-Object {$_.State -eq "Running" -or $_.State -eq "Probation"} | Select-Object -ExpandProperty Name

}

if ($ComputerName.Count -gt 0) {
    $PMs = $ComputerName #| ForEach-Object {Get-CenteralPMachine -Name $_} | Select-Object -ExpandProperty Name
}

if ($ZoneId.Count -gt 0) {
    $PMs = $ZoneId | ForEach-Object {Get-CenteralPMachine -Zone $_ -ExcludeType ChassisManager} | `
            Where-Object {$_.State -eq "Running" -or $_.State -eq "Probation"} | Select-Object -ExpandProperty Name
}

# To find out DC VUT.
$SearchBase = "ou=VUTDCs,ou=pMachines,ou=Servers,dc=chn,dc=sponetwork,dc=com"
$DCVUTs = @(Get-ADComputer -Filter * -SearchBase $SearchBase | Select-Object -ExpandProperty Name)
$DCPMs = @($PMs | Where-Object {$DCVUTs -contains $_})

$PMs = $PMs | Where-Object {$DCPMs -notcontains $_}

<#
# Test machine online or not.
$BadPMs = @(Test-MachineConnectivity $PMs -ReturnBad)
if ($BadPMs.Count -eq 0) {
    $GoodPMs = $PMs
}
else {
    $GoodPMs = @($PMs | Where-Object {$_ -notin $BadPMs})
}
#>


    $Job = Invoke-Command -ScriptBlock $ScriptBlock -ComputerName $PMs -AsJob
    $PMResult = @($Job | Watch-Jobs -Activity "Testing drive V" -Status Running)



# Deal with DC PMs
if ($DCPMs.Count -gt 0) {

    if (!$Credential) {
        $Credential = Get-Credential -Message "Your CHN domain admin:"
    }

    $BadDCPMs = @(Test-MachineConnectivity $DCPMs -ReturnBad -Credential $Credential)
    if ($BadDCPMs.Count -eq 0) {
        $GoodDCPMs = $DCPMs
    }
    else {
        $GoodDCPMs = $DCPMs | Where-Object {$_ -notin $BadDCPMs}
    }

    if ($GoodDCPMs.Count -gt 0) {
        $Job = Invoke-Command -ScriptBlock $ScriptBlock -ComputerName $GoodDCPMs -AsJob -Credential $Credential
        $DCPMResult = @($Job | Watch-Jobs -Activity "Testing drive V" -Status Running)
    }
}

if ($FormatHTML.IsPresent) {

}
else {
    return ($PMResult + $DCPMResult)
}

# =====================================================================================================================
$HtmlBody += "</table>"

