#####################################################################################
#
#File: TestDailyCheckItem.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  6/20/2017   1.0       Wind        First version.
##                                    
##
#####################################################################################

<#
.SYNOPSIS
To test runing daily check item individually.

.DESCRIPTION
This script just allow run only one check item and return html result for debug.
It will open HTML result in IE if run it with parameter 'Open'.

.PARAMETER Type
Check item name.

.PARAMETER Credential
The admin crednetial which some check item needed.

.PARAMETER Open
Open result in default Internet browser.

.EXAMPLE
PS C:\> TestDailyCheckItem.ps1 -Type GMOrch -Open
Run 'GMorch' check item and display result in browser.
#>


Param (
    [Parameter(Mandatory=$true,Position=1,ParameterSetName="Type")]
    [String] $Type,
    [Parameter(Mandatory=$true,ParameterSetName="Path")]
    [String] $Path,
    [System.Management.Automation.PSCredential]$Credential,
    [String] $XmlConfig = "\\chn\tools\DailyCheck\config-dailycheck.xml",
    [Switch] $Open
)

# pre-loading
#========================================================================
$StartTime = Get-Date

# Load configuration.
$Xml = New-Object -TypeName System.Xml.XmlDocument
$Xml.PreserveWhitespace = $false
$Xml.Load($XmlConfig)

If ($Xml.HasChildNodes -eq $false) {
    Write-Host "Can not load config file `"$XmlConfig`"!"
    Break
}

# Load common functions
$Requiredmodule = $xml.DailyCheck.Common.RequiredModule

$CurrentPath = Split-Path $MyInvocation.MyCommand.Path -Parent


try{
    $ItemPath = Join-Path $CurrentPath "DailyCheckItems" -Resolve
    
    $LibraryPath = Join-Path $ItemPath Library -Resolve
}
catch {
    Write-Host $_ -ForegroundColor Red
    break
}

# 
    
try {
    if ($PSCmdlet.ParameterSetName -eq "Type") {
        $CheckItem = Join-Path $ItemPath "$Type.ps1" -Resolve
    }
    else {
        $CheckItem = Resolve-Path $Path | select -ExpandProperty Path
        $Type = Get-Item $CheckItem | select -ExpandProperty BaseName
    }
}
catch { 
    Write-Host $_ -ForegroundColor Red
    break 
}




$helper = @(Get-ChildItem -Path $LibraryPath -Filter *helper.ps1 | Select-Object -ExpandProperty FullName)
#Load helper
# $helper | %{ . $_ }

$arguments = "-XmlConfig $XmlConfig -Module $(Convert-ArrayAsString $Requiredmodule) -Helper $(Convert-ArrayAsString $helper)"

$sls = @(Select-String -Pattern '\$Credential' -Path $CheckItem)
if (!$Credential -and $sls.Count -gt 1) {
    
    $Credential = Get-Credential -Message "Please enter your admin credential"
    if (!$Credential) {
        Write-Host "Error: Admin credential must be provide!" -ForegroundColor Red
        break
    }
    
}

if ($Credential) { $arguments += "-Credential `$Credential" }


#Main code
#========================================================================
Write-Host "Daily Check Item ($Type) Test" -ForegroundColor Green
$expression = $CheckItem, $arguments -join " "
$result = Invoke-Expression -Command $expression

$EndTime = Get-Date
$duration = $EndTime - $StartTime
$properties = @{
    StartTime = $StartTime
    EndTime = $EndTime
    Duration = $duration
}
$cost = New-Object -TypeName PSObject -Property $properties

Write-Host "Time cost" -ForegroundColor Yellow
$cost | ft -AutoSize StartTime, EndTime, Duration | Out-Host

if ($Open.IsPresent) {
    $filename = $type,"html" -join "."
    $htmlfile = Join-Path $env:TEMP $filename
    $result | Out-File $htmlfile
    start -Verb Open $htmlfile
}

return $result