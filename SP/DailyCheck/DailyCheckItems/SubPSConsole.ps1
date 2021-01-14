## This script is used for invoke sub scripts in new PS console. The sub scripts are usually used to
## collect data without HTML outputs. It will invoke new PS console for each sub script.

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
$Sseparator = "-" * $Host.UI.RawUI.WindowSize.Width

# Initialize HTML table
$HtmlBody = "<TABLE Class='SUBPSCONSOLE' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================

# Functions
function ReadCommand {
    Param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument] $Xml
    )

    # Read the xml object and return command objects.
    $Command = @($Xml.subPSConsole.Command)

    return $Command

}

function MakeExpression {
    Param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement] $Command,
        [string] $Path

    )

    # convert comamnd objects as string array and return.
    if($Path) {
        $expression = Join-Path $Path $Command.Name
    }
    else {
        $expression = $Command.Name
    }
    $arguments = ""
    foreach($arg in $Command.Argument) {
        $arguments = $arguments, $arg -join " "
    }

    $expression += $arguments

    return $expression

}

# Main code
#---------------------------------------------------------------------------------------------------------------

$TableHeader = "SubPSConsole"
$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px><TD colspan='6' align='center' style=color:#FAF4FF>"`
            + $TableHeader`
            + "</TD></TR>"

Write-Host $separator
Write-Host "$TableHeader is running"

#$subconfig = "D:\UserData\oe-songwende\MyScripts\DevTools2\config-subPSConsole.xml"
$subconfig = $Xml.DailyCheck.SubPSConsole.Config

$subconfigXml = LoadXML -Path $subconfig

$expression = @($subconfigXml.subPSConsole.Expression)

$LogPath = $subconfigXml.subPSConsole.LogPath
if(!(Test-Path $LogPath)) {
    # Create output path
    Write-Host "Log path [$LogPath] does not exist! Create it." -ForegroundColor Yellow
    New-Path $LogPath -Type Directory
}

$ErrorPath = Join-Path $LogPath "Errors"
if(!(Test-Path $ErrorPath)) {
    Write-Host "Error path [$ErrorPath] does not exist! Create it." -ForegroundColor Yellow
    New-Path $ErrorPath -Type Directory
}



$OutputPath = Join-Path $LogPath "Outputs"
if(!(Test-Path $OutputPath)) {
    Write-Host "Output path [$OutputPath] does not exist! Create it." -ForegroundColor Yellow
    New-Path $OutputPath -Type Directory
}



$command = ReadCommand -Xml $subconfigXml
$DateStr = Get-Date -Format "yyyyMMddhhmmss"

$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
$expressionPath = Join-Path $scriptPath "SubConsoleScripts"


$processObj = @()

foreach($cmd in $command) {
    
    Write-Host $Sseparator
    Write-Host "Sub console: $($cmd.Name)" -ForegroundColor Green
    $FileName = $env:USERNAME, $cmd.Name, $DateStr -join "_"
    $errorFile = ($FileName,"Error" -join "_") + ".txt"
    $errorFile = Join-Path $ErrorPath $errorFile
    Write-Host "Errors send into [$errorFile]" -ForegroundColor Yellow

    <#
    $outputFile = ($FileName,"Output" -join "_") + ".txt"
    $outputFile = Join-Path $OutputPath $outputFile
    Write-Host "Outputs send into [$outputFile]" -ForegroundColor Yellow
    #>
    $expression = MakeExpression -Command $cmd -Path $expressionPath

    Write-Host "Start to invoke [$expression]" 
    
    ## Just redirect errors into file. The standard outputs will show in console.
    $process = Start-Process -FilePath PowerShell -ArgumentList "-File $expression" -RedirectStandardError $errorFile -WindowStyle Minimized -PassThru

    # to create custome object to store process and errorfile.
    $property = @{PID=$process.Id;ErrorFile=$errorFile}
    $processObj += new-object -TypeName PSObject -Property $property

    Start-Sleep 1

}

# Startup a new PowerShell console to monitor sub-consoles.
$JanitorScript = $Xml.DailyCheck.SubPSConsole.JanitorScript -replace '%%username%%', $env:USERNAME
$Duration = $Xml.DailyCheck.SubPSConsole.Duration
if(!(Test-Path $JanitorScript)) {
    Write-Host "Janitor script [$JanitorScript] does not exsit!" -ForegroundColor Red
    # notify daily check admin.
    Send-Email -To $Xml.Common.DailyCheckAdmin -mailsubject "Janitor script can not be found!" -mailbody "Janitor script [$JanitorScript] does not exsit!"
}
else {
    
    #Export process objects as csv for janitor reading.
    $tempFile = Join-Path $env:TEMP "processObj_$DataStr.csv"
    $processObj | Export-Csv $tempFile

    $argument = "-Path $tempFile"
    if($Duration) {
        $argument += " -Duration $Duration"
    }
    $JanitorScript = $JanitorScript, $argument -join " "
    $errorFile = Join-Path $ErrorPath "JanitorErrors.txt"
    # $outputFile = Join-Path $OutputPath "JanitorOutput.txt"
    Start-Process -FilePath PowerShell -ArgumentList "-NoExit -File $JanitorScript" -RedirectStandardError $errorFile
    
}



#---------------------------------------------------------------------------------------------------------------
Write-Host $Sseparator
Write-Host "$TableHeader is done."
Write-Host $separator
# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================