#####################################################################################
#
#File: PMDisk.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  6/26/2017   1.1       Wind        Send mail notification to IM if disk failure.
##                                    
##  7/13/2017   1.2       Wind        Add array controller check.
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

# Initialize HTML table
$HtmlBody = "<TABLE Class='PMDISK' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

    Write-Host $separator

    $HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF>"`
                + "<td colspan='6' align='center'>PM Disk Check</td></TR>"

    $PMsName = (Get-Centeralzone | %{Get-CenteralPMachine -Zone $_} | ? State -NE Decommissioned).name
    $job = Get-WmiObject -Class HPSA_DiskDrive -Namespace root\hpq -ComputerName $PMsName -AsJob

    While ($job.state -eq "Running") { Start-Sleep 2 }

    $process = $true
    # If job hasn't more data, it means can't fetch disks info for all servers.
    If ($job.HasMoreData -eq $false) { 
        Write-Host "Cannot fetch disk info for all PMs." -ForegroundColor Red 
        $process = $false
    }

    $Results = $job | receive-job -ErrorAction SilentlyContinue  # Get result.


    # It means a part of PMs can't be fetch info. Get their name.
    If ($job.state -eq "Failed") {
        $badjobs = $job.Childjobs | ? state -eq "failed"
        $BadPMs = $badjobs.Location
    }

    $BadDisks = @($Results | ?{$_.OperationalStatus -eq '6' })
    $PredictiveDisks = @($Results | ?{$_.OperationalStatus -eq '5' })

    Remove-Job -Job $job

    # Update 1.2: Add array controller check
    if ($process) {
        $GoodPMs = $PMsName | ?{$_ -notin $BadPMs}

        $job = Get-WmiObject -Namespace root\hpq -Class HPSA_ArrayController -ComputerName $GoodPMs -AsJob | Out-Null

        $ACResult = $job | receive-job -ErrorAction SilentlyContinue

        If ($job.state -eq "Failed") {
            $badjobs = $job.Childjobs | ? state -eq "failed"
            $BadPMs += $badjobs.Location
            Write-Host "Below PMs cannot be fetch info." -ForegroundColor Yellow
            Write-Host $BadPMs
        }

        $BadArrayController = @($ACResult | ?{$_.OperationalStatus -eq '6' })
    }

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px>"`
                + "<td colspan='6' align='center'>Error Disk</td></TR>"
    

    if($BadDisks.count -gt 0)
    {
        Write-Host "Have bad disk!" -ForegroundColor Red
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold>"`
                    + "<TD colspan='2'>ElementName</TD><TD colspan='2'>ServerName</TD><TD colspan='2'>SerialNumber</TD></TR>"
        for($i=0;$i -lt $BadDisks.count;$i++)
        {


            # No need to test PM connectivity
            #if(Test-Connection -ComputerName $BadDisks[$i].PSComputerName -BufferSize 16 -count 1 -ea 0 -Quiet)
            #{
            $bay=$BadDisks[$i].ElementName
            $server=$BadDisks[$i].PSComputerName
            $SerialNum = Get-CenteralPMachine -Name $server | Select-Object -ExpandProperty SerialNum
            $HtmlBody += "<TR style=background-color:Red><TD colspan='2'>$bay</TD><TD colspan='2'>$server</TD><TD colspan='2'>$SerialNum</TD></TR>"
            #}
          }
   
    }
    else {
        $HtmlBody += "<TR><td colspan='6' align='left' style=color:$($CommonColor['Green'])>None</td></TR>"
    }

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px>"`
                + "<td colspan='6' align='center'>Predictive Failure Disk</td></TR>"
    

    if($PredictiveDisks.Count -gt 0) {
        Write-Host "Have predictive disk!" -ForegroundColor Yellow
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold>"`
                    + "<TD colspan='2'>ElementName</TD><TD colspan='2'>ServerName</TD><TD colspan='2'>SerialNumber</TD></TR>"
        for($i = 0;$i -lt $PredictiveDisks.count;$i++) {
            $Pbay = $PredictiveDisks[$i].ElementName
            $Pserver = $PredictiveDisks[$i].PSComputerName
            $SerialNum = Get-CenteralPMachine -Name $Pserver | Select-Object -ExpandProperty SerialNum
            $HtmlBody += "<TR style=background-color:Yellow><TD colspan='2'>$Pbay</TD><TD colspan='2'>$Pserver</TD><TD colspan='2'>$SerialNum</TD></TR>"
        }
    }
    else {
        $HtmlBody += "<TR><td colspan='6' align='left' style=color:$($CommonColor['Green'])>None</td></TR>"
    }

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px>"`
                + "<td colspan='6' align='center'>Error Array Controller</td></TR>"

    if($BadArrayController.Count -gt 0) {
        Write-Host "Have bad array controller!" -ForegroundColor Red
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold>"`
                    + "<TD colspan='2'>ElementName</TD><TD colspan='2'>ServerName</TD><TD colspan='2'>SerialNumber</TD></TR>"
        for($i = 0;$i -lt $BadArrayController.count;$i++) {
            $Pbay = $BadArrayController[$i].ElementName
            $Pserver = $BadArrayController[$i].PSComputerName
            $SerialNum = Get-CenteralPMachine -Name $Pserver | Select-Object -ExpandProperty SerialNum
            $HtmlBody += "<TR style=background-color:Yellow><TD colspan='2'>$Pbay</TD><TD colspan='2'>$Pserver</TD><TD colspan='2'>$SerialNum</TD></TR>"
        }
    }
    else {
        $HtmlBody += "<TR><td colspan='6' align='left' style=color:$($CommonColor['Green'])>None</td></TR>"
    }

    # adding bad PMs into mail.
    If ($BadPMs.count -gt 0) {
        $HtmlBody += "<TR style=background-color:#EA0000;font-weight:bold;font-size:16px;color:white><td colspan='6' align='center'>Below PMs can't be fetch disks:</td></TR>"
        $HtmlBody += "<TR><td colspan='6' align='Left'>$BadPMs</td></TR>"
        Write-Host "Bad PMs:"
        Write-Host "$(Convert-ArrayAsString $BadPMs)" -ForegroundColor Red
    }

    Write-Host "Checking for 'PMDisk' done." -ForegroundColor Green
    Write-Host $separator




# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================
