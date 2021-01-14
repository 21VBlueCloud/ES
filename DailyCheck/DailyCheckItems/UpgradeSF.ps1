#####################################################################################
#
#File: UpgradeSF.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##
##  12/5/2019    1.1       Wind       1. Distinguish 2 SF farm.
##                                    2. Change monitoring Centeral job type to 'UpgradeCenteralFarm'.
#####################################################################################

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
$HtmlBody = "<TABLE Class='UPGRADESF' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

Write-Host $separator
$Steps = @(
    "Wait Centeral package installed"
    "Set farm queue"
    "Upgrade farms"
    "Create VM images and upgrade vms"
    "Clear job"
)

$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='center'>ServiceFabric Upgrade Status</td></TR>"

$DSFarms = Get-CenteralFarm -Role DS
$PRFarms = $DSFarms | ? RecoveryFarmId -NE 0
#-----------------------------------------------------------------------------

foreach ($PRFarm in $PRFarms) {
## Update 1.1: change monitoring job to 'UpgradeCenteralFarm'.
$UpgradeJobs = @(Get-CenteralJob -Type UpgradeCenteralFarm -ObjectId $PRFarm.FarmId )
$RunningJob = $UpgradeJobs | ? State -eq Executing
$SuspendedJobs = $UpgradeJobs | ? State -eq Suspended | Sort-Object StartTime

if ($RunningJob) {
    $data = $RunningJob.Data
    $TargetVersion = (-split $data)[5]
    Write-Host "Target version is " -NoNewline
    Write-Host $TargetVersion -ForegroundColor Green
    $Step = $RunningJob.Step
    Write-Host "Upgrade SF in step " -NoNewline
    Write-Host $Step -ForegroundColor Green
    $HtmlBody += "<tr><td colspan='6'>"
    $HtmlBody += "<Table border=0 width=900>"
    $HtmlBody += "<tr><td style=background-color:$($CommonColor['Green']);font-weight:bold;font-size:16px;color:$($CommonColor['White'])>"
    $HtmlBody += "UpgradeSF in step [$Step/5] for target version $TargetVersion</td></tr>"

    if ($SuspendedJobs.Count -gt 0) {
        
        # To compare with current version. If it is equals or less than current version, it means upgrade done by successive job.
        if ($PRFarm.Version -lt $SuspendedJobs[-1].Version) {
            $message = "UpgradeSF job $($SuspendedJobs.JobId) suspended at step [$Step]"
            Write-Host $message -ForegroundColor Yellow
            $HtmlBody += "<tr><td style=background-color:$($CommonColor['Yellow'])>$message</td></tr>"
        }
        
    }

    ## Update 1.1: UpgradeSPDVM job was deprecated.
    <#
    # We should check UpgradeSPDVM jobs if step at 2.
    if ($Step -eq 2) {
        $UpgradeVMJobs = Get-CenteralJob -Type UpgradeSPDVM
        $suspendJobs = $UpgradeVMJobs | ? State -eq "Suspended"
        
        if ($suspendJobs) {
            $message = "UpgradeSPDVM job ({0}) suspended!" -f (Convert-ArrayAsString $suspendJobs.JobId)
            Write-Host $message -ForegroundColor Yellow
            $HtmlBody += "<tr><td style=background-color:$($CommonColor['Yellow'])>$message</td></tr>"
        }
    }
    #>

    $HtmlBody += "</Table>"
    $HtmlBody += "</td></tr>"
}
else {

    $message = "No upgrade Jobs.PR DS farm [$($PRFarm.FarmId)] version is [$($PRFarm.Version)]"
    $HtmlBody += "<TR style=background-color:#00A600;font-weight:bold;font-size:17px><TD colspan='6' align='center' style=color:#FAF4FF>$message</TD></TR>"
    
    Write-Host $message -ForegroundColor Green
}

}

#-----------------------------------------------------------------------------

Write-Host "UpgradeSF check done!" -ForegroundColor Green
Write-Host $separator

# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================