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
$HtmlBody = "<TABLE Class='SUBSCRIPTIONCHECK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

# Functions

# Main code

# header
$TableHeader = "SUBSCRIPTION CHECK"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"


$File = Import-Csv $xml.DailyCheck.SubscriptionCheck.InfoFile

$ExpireDays = $xml.DailyCheck.SubscriptionCheck.ExpireDays -as [Int]
$EndDays = $xml.DailyCheck.SubscriptionCheck.EndDays -as [Int]


$ExpireBody = ""
$EndBody = ""

Foreach($line in $File)
{
    
    $EDays = (New-TimeSpan -End $line.ExpirationDate).Days
    
    if($EDays -le $ExpireDays)
    {
         
         $ExpireBody += "<TR style=font-weight:bold;font-size:17px><TD align='left' style=color:#000000>$($line.DisplayName)</TD>`
                        <TD align='center' style=color:#000000>$($line.ExpirationDate)</TD>`
                        <TD align='center' style=color:#000000>$EDays</TD>`
                        <TD align='center' style=color:#000000>$($line.Operation)</TD></TR>"
         Write-Host ("{0} will expire in {1} days!" -f $line.DisplayName, $EDays) -ForegroundColor Yellow

    }

    $DDays = (New-TimeSpan -End $line.EndingOfLife).Days
    if($DDays -le $EndDays)
    {
        
        $EndBody += "<TR style=font-weight:bold;font-size:17px><TD align='left' style=color:#000000>$($line.DisplayName)</TD>`
                        <TD align='center' style=color:#000000>$($line.ExpirationDate)</TD>`
                        <TD align='center' style=color:#000000>$DDays</TD>`
                        <TD align='center' style=color:#000000>$($line.Operation)</TD></TR>"

        Write-Host ("{0} will end in {1} days!" -f $line.DisplayName, $DDays) -ForegroundColor Red
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
    $HtmlBody += "<TR style=font-weight:bold;font-size:17px><TD style=color:$($commonColor['Green'])>No subscription will expire!</TD></TR>"
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
    $HtmlBody += "<TR style=font-weight:bold;font-size:17px><TD style=color:$($commonColor['Green'])>No subscription will end of life!</TD></TR>"
}

Write-Host "Checking for '$TableHeader' done." -ForegroundColor Green
Write-Host $separator
# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================