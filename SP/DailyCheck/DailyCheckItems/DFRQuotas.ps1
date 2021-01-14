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
$HtmlBody = "<TABLE Class='DFRQUOTAS' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================
    
    Write-Host $separator

    $QuotaLowerBound = $Xml.DailyCheck.DFRQuotas.LowerThreshold
    $QuotaUpperBound = $Xml.DailyCheck.DFRQuotas.UpperThreshold

    #DFR Zone7 Quota Checking
    $HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='center'>DFR Quota usage status</td></TR>"
    $HtmlBody += "<TR style=background-color:$($Commoncolor['LightGray']);font-weight:bold><TD colspan='2'>Path</TD><TD>Used(%)</TD><TD>Size(GB)</TD><TD colspan='2'>Limited(GB)</TD></TR>"

    $Zone7DFRServerName = get-Centeralvm -role dfr -Network 11| select -First 1 -ExpandProperty Name
    
    $Zone7DFRSession = New-PSSession -ComputerName $Zone7DFRServerName 
    $zone7info = Invoke-Command -Session $Zone7DFRSession -ScriptBlock{
       $Zone7Quota = Get-FSRMQuota H:\DFSRoots\SHA_Zone7
       "{0:N0}" -f ($zone7Quota.Size/1GB)
       "{0:N0}" -f ($Zone7Quota.Usage/1GB)

    }
    $Zone7QuotaSize = $zone7info[0]
    $Zone7QuotaUsage = $zone7info[1]
    $Zone7P = $Zone7QuotaUsage/$Zone7QuotaSize

    Remove-PSSession -Session $Zone7DFRSession -ErrorAction SilentlyContinue
    
    If ($Zone7P -ge $QuotaUpperBound) {
        $HtmlBody += "<TR style=background-color:#EA0000><TD colspan='2'>SHA_Zone7</TD><TD>$("{0:P0}" -f $Zone7P)</TD><TD>$Zone7QuotaUsage</TD><TD colspan='2'>$Zone7QuotaSize</TD></TR>"
    }
    ElseIf ($zone7P -ge $QuotaLowerBound) {
        $HtmlBody += "<TR style=background-color:Yellow><TD colspan='2'>SHA_Zone7</TD><TD>$("{0:P0}" -f $Zone7P)</TD><TD>$Zone7QuotaUsage</TD><TD colspan='2'>$Zone7QuotaSize</TD></TR>"
    }
    Else {
        $HtmlBody += "<TR><TD colspan='2'>SHA_Zone7</TD><TD>$("{0:P0}" -f $Zone7P)</TD><TD>$Zone7QuotaUsage</TD><TD colspan='2'>$Zone7QuotaSize</TD></TR>"
    }

    #DFR Zone6 Quota Checking
    $Zone6DFRServerName = (get-Centeralvm -role dfr -Network 9| select name)[0].Name
    $Zone6DFRSession = New-PSSession -ComputerName $Zone6DFRServerName 
    $zone6QuotaSize = Invoke-Command -Session $Zone6DFRSession -ScriptBlock{

        $RootFolder = "H:\DFSRoots"
        $SubFolder = Get-ChildItem $RootFolder
        foreach($folder in $SubFolder)
        {
            if($folder.Name -notin "PAVC","Symbols")
            {
            $s = Get-FsrmQuota $folder.FullName 
            "{0:N0}" -f ($s.size/1GB)
            }
         }
    }
    $zone6QuotaUsage = @()
    $zone6QuotaUsage = Invoke-Command -Session $Zone6DFRSession -ScriptBlock{

        $RootFolder = "H:\DFSRoots"
        $SubFolder = Get-ChildItem $RootFolder
        foreach($folder in $SubFolder)
        {
            if($folder.Name -notin "PAVC","Symbols")
            {
            $s = Get-FsrmQuota $folder.FullName 
            "{0:N0}" -f ($s.Usage/1GB)
            }
         }
    }
    $zone6FolderName = @()
    $zone6FolderName = Invoke-Command -Session $Zone6DFRSession -ScriptBlock{

        $RootFolder = "H:\DFSRoots"
        $SubFolder = Get-ChildItem $RootFolder
        foreach($folder in $SubFolder)
        {
            if($folder.Name -notin "PAVC","Symbols")
            {
                $folder.Name
            }
         }
    }
    for([int]$i = 0; $i -lt $zone6FolderName.Count; $i++) {
        $Zone6Path = $zone6FolderName[$i]
        $S = $zone6QuotaSize[$i]
        $U = $zone6QuotaUsage[$i]
        $V = $zone6QuotaUsage[$i]/$zone6QuotaSize[$i]
        $P = "{0:P0}" -f ($V)

        If ($V -ge $QuotaUpperBound) {
            $HtmlBody += "<TR style=background-color:#EA0000><TD colspan='2'>$Zone6Path</TD><TD>$P</TD><TD>$U</TD><TD colspan='2'>$S</TD></TR>"
        }
        ElseIf ($V -ge $QuotaLowerBound) {
            $HtmlBody += "<TR style=background-color:Yellow><TD colspan='2'>$Zone6Path</TD><TD>$P</TD><TD>$U</TD><TD colspan='2'>$S</TD></TR>"
        }
        Else {
            $HtmlBody += "<TR><TD colspan='2'>$Zone6Path</TD><TD>$P</TD><TD>$U</TD><TD colspan='2'>$S</TD></TR>"
        }
    }
    Remove-PSSession -Session $Zone6DFRSession -ErrorAction SilentlyContinue

    Write-Host "Checking for 'DFRQuotas' done." -ForegroundColor Green
    Write-Host $separator


# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
##===============================================================================================================