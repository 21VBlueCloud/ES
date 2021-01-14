#####################################################################################
#
#File: MonitorFarmNLBrotation.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  1/1/2017   1.0       Wind        First version.
##                                    
##  1/8/2017   1.1       Wind        Checking 'USR' and 'BOT' both if omit Role.
##
##  1/15/2017  1.2       Wind        1. Add User alias in mail subject.
##                                   2. Add Centeral VM list in suspicous list.
#####################################################################################

<#
.SYNOPSIS
Monitor NLB rotation state.

.DESCRIPTION
This script is used for monitoring NLB rotation state which VMs out.
Default monitoring object is BOT if you don't specify parameter "Role".
This script will not return object and just display monitoring result.
If the VMs are keeping out of rotation in the monitoring duration, it will display at end of the output as red warning.

.PARAMETER Role
Option are "BOT" and "USR". Checking both if omit.

.PARAMETER FarmId
The primary content farm ID. Script will find it if you not specify it.

.PARAMETER Duration
The monitoring duration. Default is 60 minutes.

.PARAMETER Interval
The check interval. Default is 5 minutes.

.PARAMETER MailTo
The monitoring result will send, if you specify recipients.

.EXAMPLE
PS C:\> MonitorFarmNLBrotation.ps1 -Duration 30 -Interval 10 -MailTo "song.wende@oe.21vianet.com"
Minitor BOT VMs rotation of primary content farm every 10 minutes in 30 minutes and send result to Wind.

.EXAMPLE
PS C:\> MonitorFarmNLBrotation.ps1 -FarmId 697 -MailTo "o365_spo@21vianet.com"
Monitor BOT VMs rotation of farm 697 every 5 minutes in 60 minutes and send result to SPO team.

.EXAMPLE
PS C:\> MonitorFarmNLBrotation.ps1 -Role USR 
Monitor USR rotation for primary content farm every 5 minutes in 60 minutes.
#>

Param (
    [ValidateSet("BOT","USR")]
    [String] $Role,
    [Int] $FarmId,
    [Int] $Duration = 60,  # Minutes
    [Int] $Interval = 5,   # Minutes
    [String[]] $MailTo
)

# Load common functions
$moduleName = "CommonFunctions"
$module = Get-Module $moduleName
IF (!$module) {
    Write-Warning "Can not find common functions module!"
    $module = Import-Module $moduleName -PassThru
    If (!$module) {
        Write-Warning "Can not load common functions module!"
        Write-Warning "Trying loading from Wind's script folder..."
        $WindModule = "D:\UserData\oe-songwende\MyScripts\Release\Module\CommonFunctions"
        $module = Import-Module $WindModule -PassThru
        If (!$module) {
            Write-Warning "Can not load common functions from Wind's path!"
            Break
        }
    }  
}

$CommonColor = @{
    Red = "#EA0000"
    Cyan = "#00FFFF"
    Blue = "#0066CC"
    purple = "#800080"
    Yellow = "#FFFF00"
    Lime = "#00FF00"
    Magenta = "#FF00FF"
    White = "#FFFFFF"
    Silver = "#C0C0C0"
    Gray = "#808080"
    Black = "#000000"
    Orange = "#FFA500"
    Brown = "#A52A2A"
    Maroom = "#800000"
    Green = "#00A600"
    Olive = "#808000"
    Pink = "#FFC0CB"
    DarkBlue = "#00008B"
    DarkCyan = "#008B8B"
    DarkGray = "#A9A9A9"
    DarkGreen = "#006400"
    DarkMagenta = "#8B008B"
    DarkOrange = "#FF8C00"
    DarkRed = "#8B0000"
    DeepPink = "#FF1493"
    LightBlue = "#ADD8E6"
    LightCyan = "#E0FFFF"
    LightGray = "#D3D3D3"
    LightGreen = "#90EE90"
    LightYellow = "#FFFFE0"
    MediumBlue = "#0000CD"
    MediumPurple = "#9370DB"
    YellowGreen = "#9ACD32"
}

# Get PR content farm.
If (!$FarmId) {
    $ContentFarm = Get-CenteralFarm -Role Content | ? RecoveryFarmId -NE 0
    $FarmId = $ContentFarm.FarmId
}

# Initialize variable
$StartTime = Get-Date
$EscapeSeconds = 0
$Round = 0
$Records = @()

$MonitorSummary = New-Object -TypeName PSObject -Property @{
            StartTime = $StartTime
            EndTime = $null
            Duration = $Duration
            Interval = $Interval
            RoundCount = 0
        }



# Convert minutes to seconds
$Duration *= 60
$Interval *= 60

# Main code
#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
While ($true) {

    $Round ++
    Write-Host "Round:$Round" -ForegroundColor Magenta
    $CurrentNLBState = Get-CenteralFarmNLBStats -Identity $FarmId -FromCBOnly -WarningAction Ignore
    $CurrentTime = Get-Date
    $EscapeSeconds = ($CurrentTime - $StartTime).TotalSeconds -as [Int]
    $OutRotationRecords = $CurrentNLBState | ? CBRotationState -EQ Out
    If ($Role) {
        $OutRotationRecords = $OutRotationRecords | ? Name -Match $Role
    }
    $OutRotationRecords | %{ Add-Member -InputObject $_ -NotePropertyName Round -NotePropertyValue $Round
                             Add-Member -InputObject $_ -NotePropertyName Time -NotePropertyValue $CurrentTime }
    $OutRotationRecords | ft -a Name, Round, Time, CBServerState, AvgHealthScore, ReasonFromCB
    $Records += $OutRotationRecords

    If ($EscapeSeconds -gt $Duration) {
        Write-Host "Times up!" -ForegroundColor Yellow
        Break
    }
    Start-SleepProgress -Seconds $Interval -Activity "Round:$Round"
}

$EndTime = Get-Date
$MonitorSummary.EndTime = $EndTime
$MonitorSummary.RoundCount = $Round

# Find out which VMs have been existing in the monitoring duration.
$VMs = $Records | group Name | ? Count -EQ $Round | Select -ExpandProperty Name

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Output
Write-Host "`n===============================================================`n"
Write-Host "Monitor Summary:" -ForegroundColor Green
$MonitorSummary | ft -a StartTime, EndTime, Duration, Interval, RoundCount

Write-Host "Out of Rotation Records:" -ForegroundColor Green
$Records | sort Name, Round | ft -a Name, Round, Time, CBServerState, AvgHealthScore, ReasonFromCB

If ($VMs) {
    $SuspiciousWarning = $VMs | %{ $Records | ? Name -EQ $_ | ? Round -EQ $MonitorSummary.RoundCount }
    $VMString = Convert-ArrayAsString $VMs
    Write-Host "VM: $VMString are still out of rotation during this check!" -ForegroundColor Red
    $SuspiciousWarning | fl Name,AvgHealthScore,CBServerState,ReasonFromCB 
    $CenteralVMs = $VMs | %{Get-CenteralVM -Name $_}
    $CenteralVMs | ft -AutoSize VMachineId, PMachineId, NetworkId, Name, Role, State, Version
}
Else {
    Write-Host "No suspicious VMs!" -ForegroundColor Green
}

If ($MailTo) {
    $Contents = @()
    If ($SuspiciousWarning) {
        $Contents += Format-HtmlTable -Contents ($SuspiciousWarning | Select-Object Name,AvgHealthScore,CBServerState,ReasonFromCB) `
                        -Title "VM: $VMString are still out of rotation during this check!" -Titlecolor $CommonColor["Red"]
        $Contents += Format-HtmlTable -Contents ($CenteralVMs | Select-Object VMachineId, PMachineId, NetworkId, Name, Role, State, Version)
    }
    Else {
        $Contents += "<H1 style=color:$($CommonColor["Green"])>No Suspicious VMs!</H1>"
    }
    $Contents += Format-HtmlTable -Contents ($MonitorSummary | Select-Object StartTime, EndTime, Duration, Interval, RoundCount) `
                    -Title "Monitor Summary:" -Titlecolor $CommonColor["Green"]
    $Contents += Format-HtmlTable -Contents ($Records | sort Name, Round | Select-Object Name, Round, Time, CBServerState, AvgHealthScore, ReasonFromCB) `
                    -Title "Out of Rotation Records:" -Titlecolor $CommonColor["Green"]
    
    $DateString = $StartTime.ToString("yyyyMMddhhmmss")
    Send-Email -To $MailTo -mailbody $Contents -mailsubject "NLB Rotation monitoring on $DateString by $env:username" -From "NLBMonitoring@21vianet.com" -BodyAsHtml
}
