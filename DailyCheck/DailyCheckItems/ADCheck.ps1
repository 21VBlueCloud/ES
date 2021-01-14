#####################################################################################
#
#File: ADCheck(SPD).ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  5/22/2017   1.1       Wind        Simplifying outputs. Just show test result if get error. 
##                                    
##  7/18/2017   1.2       Wind        Bug fix: 
##                                      1. Change AD replication check result from CSV to string.
##                                         Because CSV mode detect AD unavailable is hard.
##                                      2. Change AD DNS test from quite to normal mode.
##                                         Because quite mode doesn't check server name unavailable.
##
##  7/19/2017   2.0       Wind        Change name from SPD.ps1 to ADCheck.ps1 and add DSE check.
#####################################################################################

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

# Functions
Function TrimBlankElement {
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [Object[]] $InputString
    )

    [String[]] $result = @()

    foreach ($line in $InputString) {
        if (![string]::IsNullOrWhiteSpace($line)) {
            $result += $line
        }
    }

    return $result

}

# Initialize HTML table
$HtmlBody = "<TABLE Class='ADCHECK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================
    
    Write-Host $separator

    Write-Host "Creating PS session for SPD,DSE...  " -NoNewline
    $SPD = Get-SPDFQDN -First 1
    $DSE = Get-CenteralVM -Role DSE -State Running | Select -First 1 -ExpandProperty Name
    try{ 
        $Session = @(New-PSSession -ComputerName $SPD)
        $Session += @(New-PSSession -ComputerName $DSE -Credential $Credential)
    }
    catch { 
        Write-Host $_ -ForegroundColor Red 
        break 
    }

    Write-Host "Done" -ForegroundColor Green

    # Script block
    <#
    $ShowReplSB = { repadmin /showrepl * }
    $TestDNSSB = { DCDiag /Test:DNS}
    $PollADSB = { DFSRDiag PollAD }
    $ReplStateSB = { DFSRDiag ReplicationState }
    #>
    $sb = {
        $Properties = @{}
        $ShowReplSB = repadmin /showrepl
        $TestDNSSB = DCDiag /Test:DNS
        $PollADSB = DFSRDiag PollAD
        $ReplStateSB = DFSRDiag ReplicationState
        $Properties.Add("ComputerName", $env:COMPUTERNAME)
        $Properties.Add("ShowRepl", $ShowReplSB)
        $Properties.Add("TestDNS", $TestDNSSB)
        $Properties.Add("PollAD", $PollADSB)
        $Properties.Add("ReplState", $ReplStateSB)
        $result = New-Object -TypeName PSObject -Property $Properties
        return $result
    }

    Write-Host "Checking...  " -NoNewline
    Invoke-Command -Session $Session -ScriptBlock $sb -AsJob | Out-Null

    $result = Watch-Jobs -Activity "AD checking..." -TimeOut 180
    Write-Host "Done" -ForegroundColor Green

    Write-Host "Closing Session...  " -NoNewline
    Remove-PSSession -Session $Session -ErrorAction SilentlyContinue
    Write-Host "Done" -ForegroundColor Green

    $SPDResult = $result | ? ComputerName -EQ ($SPD -split "\.")[0]
    $DSEResult = $result | ? ComputerName -EQ $DSE


    # Display results
    Write-Host "Generating result...  " -NoNewline
    $HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:white><td colspan='6' align='center'>AD Health Check</td></TR>"

    # SPD result
    $HtmlBody += "<TR style=background-color:$($CommonColor['Green']);font-weight:bold;color:white><TD colspan='6'>Check from $SPD</TD></TR>"

    # 1. AD replication
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>AD Replication"
    # Update 1.2 Deprecated
    <#
    # pick up the 'Number of Failures' is not 0
    $errorResult = $SPDShowReplResult | ? "Number of Failures" -NE 0
    if ($errorResult) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold>"`
                    + "<TD>Source</TD><TD>Destination</TD><TD>Naming Context</TD>"`
                    + "<TD>Last Success Time</TD><TD>Last Failure Time</TD><TD>Number of Failures</TD>"`
                    + "</TR>"

        foreach ($e in $errorResult) {
            $source = $e."Source DSA"
            $destination = $e."Destination DSA"
            $context = $e."Naming Context"
            $successTime = $e."Last Success Time"
            $failureTime = $e."Last Failure Time"
            $failures = $e."Number of Failures"
            $HtmlBody += "<TR>"`
                    + "<TD>$source</TD><TD>$destination</TD><TD>$context</TD>"`
                    + "<TD>$successTime</TD><TD>$failureTime</TD><TD>$failures</TD>"`
                    + "</TR>"
        }

        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>repadmin /showrepl *</span>`" to get details.</TD></TR>"
    }
    #>

    $SPDShowReplResult = TrimBlankElement $SPDResult.ShowRepl

    if ($SPDShowReplResult -match "error" -or $SPDShowReplResult -match "failure" -or $SPDShowReplResult -match "unavailable") {
        $HtmlBody += "<span style=background-color:$($CommonColor['Red'])> got error!</span></TD></TR>"
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>repadmin /showrepl</span>`" to test again.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"
        foreach ($line in $SPDShowReplResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
        
    }
    else {
        $HtmlBody += "</TD></TR>"
        $HtmlBody += "<TR style=Color:$($CommonColor['Green'])><TD colspan='6'>No Error</TD></TR>"
    }

    #2. DNS test
    $SPDTestDNSResult = TrimBlankElement $SPDResult.TestDNS

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>AD DNS Test"
    if ($SPDTestDNSResult -match "failure" -or $SPDTestDNSResult -match "unavailable" -or $SPDTestDNSResult -match "error") {
        $HtmlBody += "<span style=background-color:$($CommonColor['Red'])> got error!</span></TD></TR>"
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>dcdiag /test:dns</span>`" to test again.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"
        foreach ($line in $SPDTestDNSResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
    }
    else {
        $HtmlBody += "</TD></TR>"
        $HtmlBody += "<TR style=Color:$($CommonColor['Green'])><TD colspan='6'>No Error</TD></TR>"
    }

    #3. DFR pollAD
    $SPDPollADResult = TrimBlankElement $SPDResult.PollAD

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>DFSR Poll AD</TD></TR>"
    $PollADSucceeded = $false
    foreach($line in $SPDPollADResult) {
        if($line.contains('Succeeded')) {
            $PollADSucceeded = $true
        }
    }

    if ($PollADSucceeded) {
        $HtmlBody += "<TR style=color:$($CommonColor['Green'])><TD colspan='6'>Operation Succeeded</TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>DFSRDiag PollAD</span>`" to get details.</TD></TR>"
        $HtmlBody += "<TR style=color:$($CommonColor['Red'])><TD colspan='6'>Operation Failed!</TD></TR>"
    }

    #4. DFR replication state
    $SPDReplStateResult = TrimBlankElement $SPDResult.ReplState

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>DFSR Replication State</TD></TR>"
    $ReplicationStateSucceeded = $false
    foreach($line in $SPDReplStateResult) {
        if($line.contains('Succeeded')) {
            $ReplicationStateSucceeded = $true
        }
    }

    if ($ReplicationStateSucceeded) {
        $HtmlBody += "<TR style=color:$($CommonColor['Green'])><TD colspan='6'>Operation Succeeded</TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>DFSRDiag ReplicationState</span>`" to get details.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"
        foreach ($line in $SPDReplStateResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
        $HtmlBody += "<TR style=color:$($CommonColor['Red'])><TD colspan='6'>Operation Failed!</TD></TR>"
    }

    # DSE result
    $HtmlBody += "<TR style=background-color:$($CommonColor['Green']);font-weight:bold;color:white><TD colspan='6'>Check from $DSE</TD></TR>"

    # 1. AD replication
    $DSEShowReplResult = TrimBlankElement $DSEResult.ShowRepl

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>AD Replication"
    
    if ($DSEShowReplResult -match "error" -or $DSEShowReplResult -match "failure" -or $DSEShowReplResult -match "unavailable") {
        $HtmlBody += "<span style=background-color:$($CommonColor['Red'])> got error!</span></TD></TR>"
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>repadmin /showrepl</span>`" to test again.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"
        foreach ($line in $DSEShowReplResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
        
    }
    else {
        $HtmlBody += "</TD></TR>"
        $HtmlBody += "<TR style=Color:$($CommonColor['Green'])><TD colspan='6'>No Error</TD></TR>"
    }

    #2. DNS test
    $DSETestDNSResult = TrimBlankElement $DSEResult.TestDNS

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>AD DNS Test"
    if ($DSETestDNSResult -match "failure" -or $DSETestDNSResult -match "unavailable" -or $DSETestDNSResult -match "error") {
        $HtmlBody += "<span style=background-color:$($CommonColor['Red'])> got error!</span></TD></TR>"
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>dcdiag /test:dns</span>`" to test again.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"
        foreach ($line in $DSETestDNSResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
    }
    else {
        $HtmlBody += "</TD></TR>"
        $HtmlBody += "<TR style=Color:$($CommonColor['Green'])><TD colspan='6'>No Error</TD></TR>"
    }

    #3. DFR pollAD
    $DSEPollADResult = TrimBlankElement $DSEResult.PollAD

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>DFSR Poll AD</TD></TR>"
    $PollADSucceeded = $false
    foreach($line in $DSEPollADResult) {
        if($line.contains('Succeeded')) {
            $PollADSucceeded = $true
        }
    }

    if ($PollADSucceeded) {
        $HtmlBody += "<TR style=color:$($CommonColor['Green'])><TD colspan='6'>Operation Succeeded</TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>DFSRDiag PollAD</span>`" to get details.</TD></TR>"
        $HtmlBody += "<TR style=color:$($CommonColor['Red'])><TD colspan='6'>Operation Failed!</TD></TR>"
    }

    #4. DFR replication state
    $DSEReplStateResult = TrimBlankElement $DSEResult.ReplState

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;color:white><TD colspan='6'>DFSR Replication State</TD></TR>"
    $ReplicationStateSucceeded = $false
    foreach($line in $DSEReplStateResult) {
        if($line.contains('Succeeded')) {
            $ReplicationStateSucceeded = $true
        }
    }

    if ($ReplicationStateSucceeded) {
        $HtmlBody += "<TR style=color:$($CommonColor['Green'])><TD colspan='6'>Operation Succeeded</TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);color:white>"`
                    + "<TD colspan='6'>Run command `"<span style=background-color:$($CommonColor['Yellow']);"`
                    + "sfont-weight:bold;color:$($CommonColor['Black'])>DFSRDiag ReplicationState</span>`" to get details.</TD></TR>"
        $HtmlBody += "<tr><td><table border=0>"

        foreach ($line in $DSEReplStateResult) {
            $HtmlBody += "<TR><TD colspan='6'>$line</TD></TR>"
        }
        $HtmlBody += "</table></td></tr>"
        $HtmlBody += "<TR style=color:$($CommonColor['Red'])><TD colspan='6'>Operation Failed!</TD></TR>"
    }

    Write-Host "Done" -ForegroundColor Green
    Write-Host "Checking for 'ADCheck' done." -ForegroundColor Green
    Write-Host $separator


# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================