Param (
    [Parameter(Mandatory=$true)]
    [String] $xmlConfig,
    [String[]] $Module,
    [String[]] $Helper,
    [System.Management.Automation.PSCredential] $Credential
    
)

#Pre-loading
##==============================================================================================================
# Import module
if ($Module) {
    Import-Module $Module
}

#Load helper
$helper | %{ . $_ }

#Load xml configuraion file
$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load($XmlConfig)

If ($Xml.HasChildNodes -eq $false) {
    Write-Host "Can not load config file `"$XmlConfig`"!"
    Break
}

$separator = "=" * $Host.UI.RawUI.WindowSize.Width

# Initialize HTML table
$HtmlBody = "<TABLE Class='VMDISK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

Write-Host $separator

$VMs = @()
$SPDs = @()
$DSEs = @()
$BadVMs = @()

$Roles = $Xml.DailyCheck.VMDisk.Threshold.Role
$WarningTable = @{}
$AlertTable = @{}
$currNode = $Roles.FirstChild

While ($currNode) {
    if ($currNode.Warning) { $WarningTable.Add($currNode.Name,[float] $currNode.Warning) }
    if ($currnode.Alert) { $AlertTable.Add($currNode.Name,[float] $currNode.Alert) }
    $currNode = $currNode.NextSibling
}

$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='center'>VMs disk Usage</td></TR>"
$HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold><TD colspan='2'>ServerName</TD><TD>DeviceID</TD><TD>DeviceSize</TD><TD colspan='2'>Used(%)</TD></TR>"

$VMs += Get-CenteralNetwork | %{Get-CenteralVM -Network $_} `
                | ? pmachineid -ne -1 | ? name -NotLike "SPD*" | ? name -NotLike "DSE*" | Select-Object -ExpandProperty Name
$SPDs = Get-SPDFQDN
$VMs += $SPDs
$DSEs += Get-CenteralVM -Role DSE | Select-Object -ExpandProperty Name

$job = Get-WmiObject -Class Win32_logicaldisk -Filter "drivetype=3" -ComputerName $VMs -AsJob
$DSE_job = Get-WmiObject -Class Win32_logicaldisk -Filter "drivetype=3" -ComputerName $DSEs -AsJob -Credential $Credential

While ($job.state -eq "Running" -or $DSE_job.state -eq "Running") { Start-Sleep 3 }


$Results = $job | receive-job -ErrorAction SilentlyContinue
$Results += $DSE_job | receive-job -ErrorAction SilentlyContinue

If ($job.state -eq "Failed") {
    $badjobs = $job.Childjobs | ? state -eq "failed"
    $BadVMs += $badjobs.Location
}

If ($DSE_job.state -eq "Failed") {
    $badjobs = $DSE_job.Childjobs | ? state -eq "failed"
    $BadVMs += $badjobs.Location
}

foreach($Result in $Results) {
    [Float] $AlertThreshold = $Xml.DailyCheck.VMDisk.Threshold.DefaultAlert
    [Float] $WarningThreshold = $Xml.DailyCheck.VMDisk.Threshold.DefaultWarning
    $Usage = 1 - $Result.FreeSpace / $Result.Size
    $role = ($Result.PSComputerName ).SubString(0,3)
    
    if ($role -in $AlertTable.Keys) { $AlertThreshold = $AlertTable[$role] }

    if ($role -in $WarningTable.Keys) { $WarningThreshold = $WarningTable[$role] }

    If ($Usage -gt $WarningThreshold) {
        $size = "{0:N2}" -f ($Result.Size/1GB)
        $servername = $Result.SystemName
        $deviceID = $Result.DeviceID
        $percentage = "{0:P2}" -f $Usage
        if ($Usage -gt $AlertThreshold) { $color = "Red" }
        else { $color = "Yellow" }
        $HtmlBody += "<TR><TD colspan='2'>$servername</TD><TD>$deviceID</TD><TD>$size</TD><TD style=background-color:$color colspan='2'>$percentage</TD></TR>"
    }

}


If ($BadVMs.count -gt 0) {
    Write-Host "Below VMs are fail:" -ForegroundColor Red
    Write-Host $BadVMs
    $HtmlBody += "<TR style=background-color:#EA0000;font-weight:bold;font-size:16px;color:white><td colspan='6' align='center'>Below VMs can't be fetch disks:</td></TR>"
    $HtmlBody += "<TR><td colspan='6' align='Left'>$BadVMs</td></TR>"

}

Remove-Job -Job $job
Remove-Job -Job $DSE_job

Write-Host "Checking for 'VMDisk' done." -ForegroundColor Green
Write-Host $separator

# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================