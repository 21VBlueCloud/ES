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
$HtmlBody = "<TABLE Class='CheckItemTemplate' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

# Functions

# Main code
#---------------------------------------------------------------------------------------------------------------

$TableHeader = "<HEADER>"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD colspan='6' align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"



#---------------------------------------------------------------------------------------------------------------
Write-Host $separator
# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================