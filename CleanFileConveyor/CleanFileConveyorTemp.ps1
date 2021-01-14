# This script is using to delete old files in file conveyor temp folder to free disk space for GFEs.

param (
    [Parameter(Mandatory=$true)]
    [String]$ComputerName,  # The GFE's host name.
    [Int]$DaysOffSet=7,  # The folder creation time is more than 7 days.
    [Int]$Size=2    # the folder size which is greater than 2GB.
)

$ErrorActionPreference = "stop"

# Nested funtions
function deleteTempFolder {
    # delete folder including sub-folders.
    param (
        [Parameter(Mandatory=$true)]
        [String]$TempFolder
    )

    # create result object
    $properties = @{
        ComputerName = $env:COMPUTERNAME
        Success = $false
        Errors = ""
    }

    $ResultObject = New-Object -TypeName PSObject -Property $properties

    if(!(Test-Path $TempFolder)) {
        $ResultObject.Errors = "Cannot find folder '$TempFolder'."
        return $ResultObject
    }

    try {
        Remove-Item -Path $TempFolder -Recurse
    }
    catch {
        $ResultObject.Errors += "$_`n"
    }

    if($?) { $ResultObject.Success = $true }
    return $ResultObject
}

$SizeInBytes = $Size * 1024 * 1024 * 1024
$Today = Get-Date

# To check the GFE is valid or not.
# If not valid, end the script up.
$GFE = Get-CenteralVM -Name $ComputerName -Role GFE
if(!$GFE) {
    Write-Error "The server $ComputerName is not GFE or valid!"
    $LASTEXITCODE = 1
    break
}


# To get all one level sub-folders' size in the FileConveyorTemp folder.
# Need to update Get-FolderSize to add a column 'ModifiedDate'.
Write-Host "Checking folder size ..." -ForegroundColor Green
$SubFolder = Get-FolderSize -ComputerName $ComputerName -Path C:\FileConveyorTemp -SubFolder
Write-Host "Done."

# To filter out which folder is greater than 2GB.
$CandidateFolder = $SubFolder | Where-Object Size -GT $SizeInBytes

# To filter out which folder's modified date is older than days off set.
$CandidateFolder = $CandidateFolder | Where-Object {($Today - (Get-Date -Date $_.LastWriteTime)).TotalDays -gt $DaysOffSet}

if($CandidateFolder.Count -eq 0) {
    Write-Host "No folder is greater than $($Size)GB more than $DaysOffSet days!" -ForegroundColor Green
    break
}
else {
    # show folder info.
    Write-Host "Below folders will be cleaned up:"
    $CandidateFolder | Format-Table -AutoSize Size, ComputerName, FullPath, LastWriteTime
    $continue = Read-Host "Are you sure to continue?(Default: No)"
    if($continue -notmatch "^y[e]?[s]?$") {
        Write-Host "User abort!" -ForegroundColor Red
        break
    }
}

<#
# To delete folder remotely.
foreach ($folder in $CandidateFolder) {
    Write-Host "`n Start to remove folder $($folder.FullPath) on $($folder.ComputerName) ..."
    $result = Invoke-Command -ScriptBlock $Function:deleteTempFolder -ArgumentList $folder.FullPath `
        -ComputerName $folder.ComputerName
    Write-Host "Done."
    $result | formate-Table -AutoSize
    Write-Host ("-" * 20)    
}
#>