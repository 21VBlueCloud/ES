# This script is used to monitor sub PowerShell console.
# Send report Email if sub consoles got error.

<#
.PARAMETER Path
Path is point to a csv file which store the process info.
The csv include fields: 
    PID: The PowerShell console process ID.
    errorFile: The error file path.

.PARAMETER Duration
Duration is this script's maximum running time. Default is 60 seconds.
#>

Param (
    [Parameter(Mandatory=$true)]
    [string] $Path,
    [Int] $Duration = 60  #minutes
)

$StartTime = Get-Date
Write-Host "Start at [$StartTime]"
$DurationSeconds = $Duration * 60

$thisScript = $MyInvocation.MyCommand.Path

# change console title
$Host.UI.RawUI.WindowTitle = "Sub PS console",(Split-Path $thisScript -Leaf) -join " - "


if(!(Test-Path $Path)) {
    $msg = "File [$Path] does not exist!"
    Write-Error $msg
    Write-Host $msg
    exit 1
}

$InfoTable = Import-Csv -Path $Path -ErrorAction Ignore

if(!$InfoTable) {
    $msg = "File [$Path] is empty!"
    Write-Host $msg
    Write-Error $msg
    exit 2
}

$Id = $InfoTable.PID
Write-Host "Find $($Id.Count) process info. [Id = $(Convert-ArrayAsString $Id)]"

$process = @(Get-Process -Id $Id -ErrorAction Ignore)
Write-Host "Find $($process.Count) processes."

# to check the process is exist or not.
if($Id.Count -ne $process.Count) {
    $NoNPID = $Id | ? {$_ -notin $process.Id}
    $msg = "Process [$(Convert-ArrayAsString $NoNPID -Quote None)] not exist!"
    Write-Error $msg
    Write-Host $msg
    
}

# To check all process is PowerShell
$validProcess = @()
foreach($p in $process) {
    if($p.Name -ne "powershell") {
        $msg = "The process [Id=$($p.Id)] is not PowerShell!"
        Write-Error $msg
        Write-Host $msg
    }
    else {
        $validProcess += $p
    }
}

$finishedProcess = @()

while($true) {

    $runningProcess = @()
    $currentDate = Get-Date
    $currentDateStr = $currentDate.ToString('[yyyy/MM/dd hh:mm:ss]')

    foreach($p in $validProcess) {
        $CurrentProcess = $p | Get-Process
        if($CurrentProcess.HasExited) {
            $finishedProcess += $CurrentProcess
        }
        else {
            $runningProcess += $CurrentProcess
        }
    }

    # calculate elapse time.
    $elapse = $currentDate - $StartTime | select -ExpandProperty TotalSeconds
    
    if($runningProcess.Count -eq 0) {
        Write-Host "$currentDateStr No process is running!" -ForegroundColor Green
        break
    }
    else {
        Write-Host "$currentDateStr $($runningProcess.Count) processes are still running! [Elapse: $($elapse.toString('.00'))]"
        $validProcess = $runningProcess
    }

   

    if($elapse -ge ($DurationSeconds)) {
        Write-Host "$currentDateStr Monitoring period expired!" 
        Write-Error "$currentDateStr Monitoring period expired at [$currentDate]!"
        Write-Host "There are still $($runningProcess.Count) process running!" -ForegroundColor Yellow
        $runningProcess | ft -AutoSize
        break
    }

    Start-SleepProgress -Seconds 5 -Activity "Waiting for refresh..." -Status Waiting

}

$endDate = Get-Date

if($finishedProcess) {
    
    $to = Get-ADUser $env:USERNAME -Properties mail | select -ExpandProperty mail
    
    # check error files
    $ProcessId = $finishedProcess.Id
    $errorFilePath = $InfoTable | ? PID -In $ProcessId | select -ExpandProperty ErrorFile
    $errorFile = Get-Item -Path $errorFilePath -ErrorAction Ignore

    foreach($p in $errorFile) {
        # to check every error file. Read it and send error content, if its length greater than 0.
        if($p.Length -gt 0) {
            $errorContent = $p | Get-Content
            Send-Email -To $to -mailbody $errorContent -mailsubject "$($p.BaseName)" -From SPODailyCheck@21vianet.com 
        }
    }
}
else {
    # if no finish process after the checking period, it means there no valid process to monitor.
    $msg = "No any valid process is running!"
    Write-Error $msg
    Write-Host $msg
    exit 1
}
