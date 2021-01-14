#####################################################################################
#
#File: RecordHelper.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  9/19/2016   1.0       Wind        First version.
##                                    
##
##  7/10/2018   1.2       Wind        Bug fix:
##                                      Update function  GetLastdayRecord due to it
##                                      cannot find the right last day records.
#####################################################################################

Function AbstractDate {
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [DateTime[]] $InputObject
    )
    $Date = @()
    foreach ($obj in $InputObject) {
        $Date += $obj.ToShortDateString()
    }

    return $Date
}

Function AddRecord {
    Param (
        [Parameter(Mandatory=$true)]
        [String] $CsvFullPath,
        [Parameter(Mandatory=$true)]
        [Object] $InputObject
    )

    If (!(Test-Path $CsvFullPath)) {
        Write-Host "Creating '$CsvFullPath'." -ForegroundColor Cyan
        New-Item -Path $CsvFullPath -ItemType File
        $InputObject | Export-Csv -Path $CsvFullPath
    }

    # Record once every day.
    $HistoryRecord = @(Import-Csv -Path $CsvFullPath)
    $AddRecord = $true
    if ($HistoryRecord) {
        $Today = Get-Date
        $TodayShortDate = $Today.ToShortDateString()
        $HistoryShortDate = AbstractDate $HistoryRecord.DateTime
        if ($HistoryShortDate -contains $TodayShortDate) {
            $AddRecord = $false 
            Write-Host "Will not record due to it done today!" -ForegroundColor Yellow
        }
    }

    If ($AddRecord) {
        Write-Host "Export record!" 
        $InputObject | Export-Csv -Path $CsvFullPath -Append
    }
}

Function GetRecord {
    Param(
        [Parameter(Position=1)]
        [Object[]] $InputObject,
        [Int] $Dayoffset = 0
    )

    if (!$InputObject) { return }

    $today = Get-Date
    $offsetday = $today.AddDays($Dayoffset)
    $offsetdayShortDate = $offsetday.ToShortDateString()
    $result = @()
    foreach ($obj in $InputObject) {
        $ObjShortDate = AbstractDate $obj.DateTime
        if ($offsetdayShortDate -eq $ObjShortDate) { $result += $obj }
    }
    
    return $result
}

Function GetYesterdayRecord {
    Param(
        [Parameter(Position=1)]
        [Object[]] $InputObject
    )

    return (GetRecord -InputObject $InputObject -Dayoffset -1)
}

Function GetLastdayRecord {
    Param(
        [Parameter(Position=1)]
        [Object[]] $InputObject
    )

    $today = Get-Date
    $todayShortDateStr = $today.ToShortDateString()
    $output = @()
    
    # The input objects are sorted by DateTime descending.
    # we check the date time from rear to find out the last date.

    for($i = $InputObject.Length - 1;$i -ge 0;$i--) {
        $currentDate = Get-Date -Date $InputObject[$i].DateTime
        $currDateStr = $currentDate.ToShortDateString()
        if($currDateStr -eq $todayShortDateStr) {
            continue
        }
        else {
            if($output) {
                if($currentDate.Equals($outputDateTime)) {
                    $output += $InputObject[$i]
                }
                else {
                    break
                }
            }
            else {
                $output += $InputObject[$i]
                $outputDateTime = Get-Date -Date $InputObject[$i].DateTime
            }
        }
    }

    return $output

}

Function GetTodayRecord {
    Param(
        [Parameter(Position=1)]
        [Object[]] $InputObject
    )

    return (GetRecord -InputObject $InputObject)
}

