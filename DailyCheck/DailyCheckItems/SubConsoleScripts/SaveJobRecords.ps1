

Param (
    [Parameter(Mandatory=$true)]
    [String] $ResourcePath,
    [Parameter(Mandatory=$true)]
    [String] $RecordPath
)

$StartTime = Get-Date

$thisScript = $MyInvocation.MyCommand.Path
$ScriptName = Split-Path $thisScript -Leaf

$Host.UI.RawUI.WindowTitle += " - $ScriptName"


$RecordFile = Join-Path $RecordPath "jobrecord.csv"
$HistoryFile = Join-Path $RecordPath "History.txt"
$StorePath = Join-Path $RecordPath "Stores"

# Write a empty line.
Write-Host "`n`r"

#---------------------------------------------------------------------------------------------
# Test path

if(!(Test-Path $ResourcePath)) {
    Write-Host "Can not find job recource path [$ResourcePath]!`n`r" -ForegroundColor Red
    Break
}

if(!(Test-Path $RecordPath)) {
    Write-Host "Can not find record path [$RecordPath]!`n`r" -ForegroundColor Red
    Break
}

if(!(Test-Path $StorePath)) {
    # Create store path
    Write-Host "Can not find store path [$StorePath]! Create it.`n`r" -ForegroundColor Yellow
    New-Path $StorePath -Type Directory
}
#---------------------------------------------------------------------------------------------

# Backup record file if it more than 300MB.
$date = Get-date -Format yyyyMMddhhmmss
$recordsize = (Get-Item -Path $RecordPath -ErrorAction Ignore).Length 
if ($recordsize -gt 300MB){
    Write-Host "Record size greater than 300MB. Backup it!`n`r"
    $BackupFile = Join-Path $StorePath "jobrecord_$date.csv"
    try {
        Move-Item -path $RecordFile -Destination $BackupFile -ErrorAction Stop
    }
    catch {
        Write-Error "The file [$RecordFile] cannot be backup!"
        Write-Error $_
    }

    Write-Host "Backup done.`n`r"

}


# backup history file before flush history.
$HistoryBackupFile = Join-Path $StorePath "History_$date.txt"
try {
    Copy-Item $HistoryFile $HistoryBackupFile -ErrorAction Stop
}
catch {
    Write-Error "Backup history file failed!"
    Write-Error $_
}

Write-Host "Starting get job ID from [$ResourcePath].`n`r"
$jobid = @(Get-ChildItem -Path $ResourcePath -Directory | select -ExpandProperty Name)

Write-Host "Getting job history from [$HistoryFile].`n`r"
$Historyjobid = Get-Content $HistoryFile -ErrorAction Ignore

Write-Host "Filtering new jobs.`n`r"
if($Historyjobid) {
    
    $queryjobid = $jobid | ? {$Historyjobid -NotContains $_}
}
else {
    $queryjobid = $jobid
}

Write-Host "Find $($queryjobid.Count) new jobs.`n`r"
Write-Host "Querying new job records.`n`r"
$newJob = @()
for ($i=0;$i -lt $queryjobid.Count;$i++) {
    Write-Host "Querying is in progressing [$($i+1) of $($queryjobid.Count)]"
    $newJob += $queryjobid[$i] | Get-CenteralJob -IncludeDeleted -ErrorAction Ignore
}


# Flush history
Write-Host "Saving job history...`n`r"
$jobid | Out-File $HistoryFile

# Save new job
Write-Host "Saving job records...`n`r"
$newJob |  Export-Csv $RecordFile -Append

Write-Host "Done.`n`r"

$EndTime = Get-Date
$TimeCost = $EndTime - $StartTime
Write-Host "Script time cost: $TimeCost`n`r"