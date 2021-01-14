#####################################################################################
#
#File: CheckSubscription.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  11/28/2018   1.1      Emma        Change mail sending policy to send mail every
##                                    time even if no subscription will expire.
#####################################################################################



Param (
    [Parameter(Mandatory=$true)]
    [string] $InfoFile,
    [Parameter(Mandatory=$true)]
    [Int] $ExpireDays,
    [Parameter(Mandatory=$true)]
    [Int] $EndDays,
    [Parameter(Mandatory=$true)]
    [string] $MailTo
)

Write-Host "Starting check subscription.`n`r"

try {
    $File = Import-Csv $InfoFile -ErrorAction Stop
}
catch {
    Write-Host "Cannot load file [$InfoFile]!`n`r" -ForegroundColor Red
    Write-Error "Cannot load file [$InfoFile]!`n`r"
    Write-Error $_ 
}

$ExpireBody = ""
$EndBody = ""

$HtmlBody = "<TABLE Class='SUBSCRIPTIONCHECK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
$TableHeader = "SUBSCRIPTION CHECK"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"


Foreach($line in $File)
{
    
    $EDays = (New-TimeSpan -End $line.ExpirationDate).Days
    
    if($EDays -le $ExpireDays)
    {
         
         $ExpireBody += "<TR style=font-weight:bold;font-size:17px><TD align='left' style=color:#000000>$($line.DisplayName)</TD>`
                        <TD align='center' style=color:#000000>$($line.ExpirationDate)</TD>`
                        <TD align='center' style=color:#000000>$EDays</TD>`
                        <TD align='center' style=color:#000000>$($line.Operation)</TD></TR>"
         Write-Host ("{0} will expire in {1} days!`n`r" -f $line.DisplayName, $EDays) -ForegroundColor Yellow

    }

    $DDays = (New-TimeSpan -End $line.EndingOfLife).Days
    if($DDays -le $EndDays) {
        
        $EndBody += "<TR style=font-weight:bold;font-size:17px><TD align='left' style=color:#000000>$($line.DisplayName)</TD>`
                        <TD align='center' style=color:#000000>$($line.ExpirationDate)</TD>`
                        <TD align='center' style=color:#000000>$DDays</TD>`
                        <TD align='center' style=color:#000000>$($line.Operation)</TD></TR>"

        Write-Host ("{0} will end in {1} days!`n`r" -f $line.DisplayName, $DDays) -ForegroundColor Red
    }

}

# To assemble html body
if ($ExpireBody) {
    $HtmlBody += "<TR><TD>"
    $HtmlBody += "<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
    $HtmlBody += "<TR style=background-color:$($CommonColor['Yellow']);font-weight:bold;font-size:17px><TD colspan=4 style=color:$($CommonColor['Red'])>Below Accounts will be expired!</TD></TR>"
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold;font-size:17px>`
                    <TD align='center'>Display Name</TD><TD align='center'>Expiration Date</TD>`
                    <TD align='center'>Left days</TD><TD align='center'>Operator</TD></TR>"

    $HtmlBody += $ExpireBody
    $HtmlBody += "</TABLE>"
    $HtmlBody += "</TD></TR>"
}
else {
    $HtmlBody += "<TR><TD>"
    $HtmlBody += "<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
    $HtmlBody += "<TR style=background-color:#00A600;font-weight:bold;font-size:15px><TD align='center' style=color:#FAF4FF>There is no Service Tenant will be expired in 30 days</TD></TR>"
    $HtmlBody += "</TABLE>"
    $HtmlBody += "</TD></TR>"
}

if ($EndBody) {
    $HtmlBody += "<TR><TD>"
    $HtmlBody += "<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
    $HtmlBody += "<TR style=background-color:#EA0000;font-weight:bold;font-size:17px><TD colspan=4 style=color:#FAF4FF>Below Accounts will be ended!</TD></TR>"
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold;font-size:17px>`
                    <TD align='center'>Display Name</TD><TD align='center'>End Date</TD>`
                    <TD align='center'>Left days</TD><TD align='center'>Operator</TD></TR>"

    $HtmlBody += $EndBody
    $HtmlBody += "</TABLE>"
    $HtmlBody += "</TD></TR>"
}
else {
    $HtmlBody += "<TR><TD>"
    $HtmlBody += "<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
    $HtmlBody += "<TR style=background-color:#00A600;font-weight:bold;font-size:15px><TD align='center' style=color:#FAF4FF>There is no Service Tenant will be end in 90 days</TD></TR>"
    $HtmlBody += "</TABLE>"
    $HtmlBody += "</TD></TR>"
}

$HtmlBody += "</table>"

# Send email.

#Write-Host "Got some subscription will expire!`n`r" -ForegroundColor Yellow
#Write-Host "Sending notification E-mail to $MailTo"
Send-Email -To $MailTo -mailbody $HtmlBody -mailsubject $TableHeader -BodyAsHtml -From "SPODailyCheck@21vianet.com"
