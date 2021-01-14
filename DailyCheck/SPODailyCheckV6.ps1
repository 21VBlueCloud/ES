<#
#####################################################################################
#
#File: SPODailyCheckv6.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  9/19/2016   1.0       Wind        V6 edition.
##                                    
##  5/16/2017   1.1       Wind        Add helper file to transfer into check items.
##
##  6/14/2017   1.2       Wind        Add attachments which retrive path from configuration file.
##  
##  6/16/2017   1.3       Wind        Change separator width to fit console width.
##
##  6/20/2017   1.4       Wind        Add mark in mailbody if check item return null.
## 
##  6/21/2017   1.5       Wind        Add 'Enabled' attribute in configuration file as a switch for
##                                    replacing 'displayorder' as validation. 'displayorder' just using
##                                    for display order.
## 
##  7/10/2017   1.5.1    Wind         Bug fix for Function MakeUpHTMLFileName:
##                                        change filename replacement for '", "',"_".
##
##  5/8/2018    1.5.2    JK           Add type SQLDistribution.
##
##  5/16/2018   1.5.3    Wind         Add "SubscriptionCheck".
##
##  7/3/2018    1.5.4    Wind         Add "SubPSConsole". That is using to invoke new PowerShell console to 
##                                    execute new job which no outputs with mail.
#####################################################################################
#>


<#
.SYNOPSIS
This script is used for SPO daily check.

.DESCRIPTION
There are 10 types checking "GMOrch","DeployBuildOrch","Tenants","SPOIPU","SearchIPU","DFRQuotas","HardwareCheck","VMDisk","SPD","VMCapacity" and "OEADAccount". 
If you are not specify types, it will check them all.
Credential must be your domain admin account becuase it's used for checking DSE. Credential diaglog box pop up, if you are not provide it
and it must be used in check.

.EXAMPLE
PS D:\>.\SPODailyCheckV5-3.6.ps1 -Types "GMOrch","DeployBuildOrch","Tenants"
Check Centeralmanager upgrade orch, Deploy build orch and tenants only.

.EXAMPLE
PS D:\>.\SPODailyCheckV5-3.6.ps1 -Credential $cred
Check all of types as default with credential $cred.

.EXAMPLE
PS D:\>.\SPODailyCheckV5-3.6.ps1 -MailTo "song.wende@oe.21vianet.com"
Check all of types and mail to Wind.

#>

Param (
    [ValidateSet("GMOrch","DeployBuildOrch","Tenants","UpgradeSF","SPOIPU","SearchIPU","VMDisk","HardwareCheck","ADCheck","VMCapacity","OEADAccount","SQLDistribution","SubPSConsole")]
    [Parameter(ParameterSetName="types")]
    [String[]] $Types,
    [Parameter(Mandatory=$true,ParameterSetName="withoutSubConsole")]
    [Switch] $WithoutSubConsole,
    [System.Management.Automation.PSCredential]$Credential,
    [String] $MailTo,
    [String] $XmlConfig = "\\chn\tools\DailyCheck\config-dailycheck.xml",
    [String] $Export
)

# Function
#=============================================================================
Function SplitResultIntoHashTable {
    Param ([string[]] $InputObject, [string[]] $Keys)
    $HashTable = @{}
    foreach ($string in $InputObject) {
        foreach ($key in $keys) {
            $key = $key.ToUpper()
            if ($string.IndexOf("$key") -ge 0) { $HashTable.Add($key,$string); Break }
        }
    }
    return $HashTable
}

function MakeUpHTMLFileName {
    
    param(
        [string] $InputString,
        [string] $Path
    )

    $filename = $InputString -replace "/",""
    $filename = $filename -replace '", "',"_"
    $filename = $filename -replace '","',"_"
    $filename = $filename -replace " ","_"
    $filename = $filename -replace '"',""
    if($Path) {
        $filename = Join-Path $Path $filename
    }
    $filename += ".html"

    return $filename

}

#=============================================================================

$StartTime = Get-Date
$CurrentPath = Split-Path $MyInvocation.MyCommand.Path -Parent
# $XmlPath = Join-Path $CurrentPath $XmlConfig
$XmlPath = $XmlConfig

# Load configuration.
$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load($XmlPath)

If ($Xml.HasChildNodes -eq $false) {
    Write-Host "Can not load config file `"$XmlPath`"!"
    Break
}

# Load common functions
$Requiredmodule = $xml.DailyCheck.Common.RequiredModule
IF ($Requiredmodule) {

    $module = Import-Module $Requiredmodule -PassThru
    If (!$module) {
        Write-Warning "Can not load required module!"
        Break
    }  
}

if (!$Export) {
    $Export = $Xml.DailyCheck.Common.Export
}

# Get display order
$DisplayOrder = @{}
$NodeDailyCheck = $xml.DailyCheck
$CurrentNode = $NodeDailyCheck.FirstChild
$EnabledItems = @()
$DisabledItems = @()

do {
    $Enabled = $CurrentNode.GetAttribute("Enabled")
    
    switch ($Enabled) {
        "True" {
            $EnabledItems += $CurrentNode.Name
            $order = $CurrentNode.GetAttribute("DisplayOrder")
            if ($order) {
                $DisplayOrder.Add($order,$CurrentNode.Name)
            }
            break
        }

        "False" {
            if(!$Types -or $CurrentNode.Name -in $Types) {
            $DisabledItems += $CurrentNode.Name
            }
        }
    }
    $CurrentNode = $CurrentNode.NextSibling
} while ($CurrentNode)

# Get check items
if (!$Types) {
    [string[]]$Types = $EnabledItems
}
else {
    # check overlap between Types and EnabledItems
    
    $CandidateItems = @()
    foreach ($t in $Types) {
        if ($t -in $EnabledItems) { $CandidateItems += $t }
    }

    if ($CandidateItems) {
        $Types = $CandidateItems
    }
    else {
        Write-Host "You selected check types are disabled!" -ForegroundColor Red
        break
    }

}


$ItemsPath = Join-Path $CurrentPath "DailyCheckItems"
$CheckItems = $Types | %{Join-Path $ItemsPath "$_.ps1"} | Get-ChildItem
#Update 1.1: Load helper 
$LibraryPath = Join-Path $ItemsPath Library
if (!$LibraryPath) { Write-Host "Cannot find library '$LibraryPath'!" -ForegroundColor Red; break }
$helper = @(Get-ChildItem -Path $LibraryPath -Filter *helper.ps1 | Select-Object -ExpandProperty FullName)
#Load helper
$helper | %{ . $_ }


$separator = "=" * $Host.UI.RawUI.WindowSize.Width

Write-Host "`n"
if ($DisabledItems) {
    Write-Host "Warning: check type(s) $(Convert-ArrayAsString $DisabledItems) was/were disabled!" -ForegroundColor Yellow
}
Write-Host $Types -ForegroundColor DarkGreen -NoNewline -Separator ","
Write-Host " will be check. Start from $StartTime"
Write-Host $separator

#Emial Properties
$MailServer = $Xml.DailyCheck.Common.SmtpServer

$MailSubject = "SPO Daily Check V6 $(Convert-ArrayAsString $Types) on $($StartTime.ToString("yyyyMMddhhmmss")) by $env:username"
$MailSender = $Xml.DailyCheck.Common.MailSender
If ($MailTo -eq "") {
    $MailRecevier = $Xml.DailyCheck.Common.MailRecevier
}
Else {
    $MailRecevier = $MailTo
}

# To inspect the 'Credential' is whether used in sub script or not. If yes, cred should be input.
$sls = Select-String -Pattern '\$Credential' -Path $CheckItems.FullName
$groupedsls = $sls | Group-Object filename | ? Count -gt 1
If ($groupedsls -and !$Credential) {
    $Credential = Get-Credential
}

# Kick off sub scripts

foreach ($chkItem in $CheckItems) {
    $arguments = @($xmlPath,$RequiredModule,$helper)
    if($chkItem.BaseName -eq "SubPSConsole") {
        if($PSCmdlet.ParameterSetName -eq "withoutSubConsole") { 
            $KickoffSub = $false 
            Write-Host "It will not kick off sub PS console jobs!" -ForegroundColor Yellow
        }
        else { $KickoffSub = $true }
        $SubPsConsole = $chkItem.FullName
        continue
    }
    if ($chkItem.name -in $groupedsls.Name) { $arguments += $Credential }
    Write-Host "Kick off job " -NoNewline; Write-Host "'$($chkItem.BaseName)'..." -NoNewline
    Start-Job -FilePath $chkItem.FullName -ArgumentList $arguments -Name $chkItem.BaseName | Out-Null
    Write-Host "done" -ForegroundColor Green
}

Write-Host $separator

# Kick off sub PS console
if($KickoffSub) {

    $SubPsConsole += " -XmlConfig $XmlPath"
    $SubPsConsole += " -Module $(Convert-ArrayAsString $Requiredmodule)"
    $SubPsConsole += " -Helper $(Convert-ArrayAsString $helper)"

    Invoke-Expression $SubPsConsole | Out-Null

}
# Watch background jobs
Write-Host "Watching jobs"
Write-Host $separator
$result = @(Watch-Jobs -Activity "Monitoring job $(Convert-ArrayAsString $Types)" -Status Running -TimeOut 240)

# fill result in hash table
$ResultHashTable = SplitResultIntoHashTable -InputObject $result -Keys $CheckItems.BaseName

# order outputs
$MailBody = ""
for ($i=1;$i -le $DisplayOrder.Count;$i++) {
    $key = $DisplayOrder["$i"]
    if($ResultHashTable["$key"]) {
        $MailBody += $ResultHashTable["$key"]
    }
    elseif ($key -in $Types) {
        $EmptyBody = "<TABLE Class='$($key.ToUpper())' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
        $EmptyBody += "<TR style=background-color:$($CommonColor['Red']);font-weight:bold;font-size:17px>"
        $EmptyBody += "<TD colspan='6' align='center' style=color:$($CommonColor['White'])>No result for '$key'."
        $EmptyBody += "<br>Please re-run this type later or report error to Wind.</TD>"
        $EmptyBody += "</TR></TABLE>"
        $MailBody += $EmptyBody
    }
}

if ($MailBody) {
    Send-Email -To $MailRecevier -mailsubject $MailSubject -From $MailSender -mailbody $MailBody -SmtpServer $MailServer -BodyAsHtml -XmlConfig $XmlPath
}

# Save outputs
If (Test-Path $Export) {
    
    $file = MakeUpHTMLFileName -InputString $MailSubject -Path $Export
    $MailBody | Out-File $file
}
Else {
    Write-Host "Invalid path `"$Export`"!" -ForegroundColor Red
    Write-Host "Can not export result !" -ForegroundColor Red
}

$EndTime = Get-Date
Write-Host "All Checking $(Convert-ArrayAsString $Types) were finished. Ended at $EndTime"
$escapeTime = $EndTime - $StartTime
Write-Host "Total cost time: $escapeTime"