#####################################################################################
#
#File: cookieage.ps1
#Author: Wind Song
#Version: 1.0
#
##  Revision History:
##  Date        Version    Alias       Reason for change
##  --------    -------   --------    ---------------------------------------
##  2/19/2019    1.0      Wind          First version.
##     
#####################################################################################


Param (
    [Parameter(Mandatory=$true)]
    [Int] $Threshold,  # cookie age delayed duration (Unit: minutes).
    [Parameter(Mandatory=$true)]
    [String] $MailTo      
)

# Initialize HTML table
$HtmlBody = "<TABLE Class='Cookie Age' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
$TableHeader = "SNC Cookie Age Check"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"

# kick off cookie age job
$SPD = Get-CenteralVM -Role SPD | select -First 1
$Job = Get-SharePointSNCCookieAgeFromSPODS -TargetSPD $SPD -NoWait
Write-Host "Kick off cookie age job [$($Job.JobId)]." -ForegroundColor Green


if($Job) {
    # Checking cookie age result
    $CheckMaxCount = 20
    while ($CheckMaxCount -gt 0) {
        # Write-Host "Waiting Cookie Age job ... round $(11 - $CheckMaxCount)" -ForegroundColor Green
        $CheckMaxCount --
        $Job = Get-CenteralJob $Job -IncludeDeleted
        if($Job.State -eq "Deleted") { break }
        Start-SleepProgress -Seconds 10 -Status "Waiting" -Activity "Cookie age job [$($Job.JobId)] is running ... round $(11 - $CheckMaxCount)"
    }

    if($Job.State -eq "Deleted") {
        Get-CenteralJobResult -Identity $Job -ProgressVariable CookieResult

        $pattern = "SNC Cookie Value for DsFqdn \[SPODS00300022.MsoCHN.msft.net\] is \[(\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{1,2}:\d{1,2} [AP]M)\]"
        $regex = [regex]::Match($CookieResult, $pattern)
        $CookieAge = get-date $regex.Groups[1].Value
        $CookieAge = [TimeZoneInfo]::ConvertTimeFromUtc($CookieAge, [TimeZoneInfo]::Local)
        
    }
    else {
        Write-Host "Cookie age job [$($Job.JobId)] have not been finished!" -ForegroundColor Red
        # Adding html warning.
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:$($CommonColor['White']);font-weight:bold;font-size:16px><td colspan='6' align='left'>Cookie age job [$($Job.JobId)] have not been finished!</td></TR>"
    }

    if($CookieAge) {

        $CurrentDate = Get-Date

        Write-Host "Cookie age is [$CookieAge] PST."
        Write-Host "Current Time is [$CurrentDate] PST."

        $TimeShow = "[Cookie Age(PST):$CookieAge] "
        $TimeShow += "[Current Time(PST):$CurrentDate]"

        if(($CurrentDate - $CookieAge).TotalMinutes -gt $Threshold) {
            Write-Host "Cookie age is older than $Threshold minutes." -ForegroundColor Red
            $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:$($CommonColor['White']);font-weight:bold;font-size:16px><td colspan='6' align='left'>$TimeShow Cookie Age is older than 30 minutes!</td></TR>"
        }
        else {
            Write-Host "Cookie age is freshness." -ForegroundColor Green
            $HtmlBody += "<TR style=background-color:$($CommonColor['Green']);color:$($CommonColor['White']);font-weight:bold;font-size:16px><td colspan='6' align='left'>$TimeShow Cookie Age is frashness!</td></TR>"
        }
        
    }
    else {
        Write-Host "Cannot get Cooike age!" -ForegroundColor Red
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:$($CommonColor['White']);font-weight:bold;font-size:16px><td colspan='6' align='left'>No Cooike Age found! Please run `"Get-SharePointSNCCookieAgeFromSPODS`" manually!</td></TR>"
    }
}
else {
    Write-Host "Kick off cookie age job failed!" -ForegroundColor Red
    $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:$($CommonColor['White']);font-weight:bold;font-size:16px><td colspan='6' align='left'>Kick off cookie age job failed! Please run `"Get-SharePointSNCCookieAgeFromSPODS`" manually!</td></TR>"
}

$HtmlBody += "</table>"

Send-Email -To $MailTo -mailbody $HtmlBody -mailsubject $TableHeader -BodyAsHtml -From "SPODailyCheck@21vianet.com"
