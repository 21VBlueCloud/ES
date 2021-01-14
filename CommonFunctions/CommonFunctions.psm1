<#
#####################################################################################
#
#File: Common_functions.psm1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date      Version   Alias       Reason for change
##  --------  -------   --------    ---------------------------------------
##  11/14/2016  1.0       Wind       First version.
##                                   Function:
##                                    Format-HtmlTable
##                                    Send-Email
##                                    Format-TenantUri
##                                    Get-SPDFQDN
##                                    
##  11/23/2016  1.1       Wind      Add function: Watch-Jobs   
##
##  12/6/2016   1.1.1     Wind      Watch-Jobs: Adding blocked job treatment.
##
##  12/17/2016  1.1.2     Wind      Move Format-TenantUri out to tenant helper.
##
##  12/19/2016  1.1.3     Wind      Move 'Generate-RandomPassword' in from OCEprovisioning.    
##
##  12/27/2016  1.1.4     Wind      Add function: Start-SleepProgress
##
##  12/28/2016  1.1.5     Wind      Add parameter TitleColor for Format-HtmlTable.
##
##  1/2/2016    1.1.6     Wind      1. Delete mandatory restriction of parameter "Contents" in function Format-HtmlTable.  
##                                  2. Add function: Convert-ArrayAsString
##
##  2/9/2017    1.1.7     Wind      Add function: Get-FileName     
##
##  4/26/2017   1.1.8     Wind      Add function: Convert-StringToHTML    
##                                    Update Format-HtmlTable: convert string to html table.  
##
##  4/26/2017   1.2.0     Wind      Add Function:
##                                    Get-EncryptedString
##                                    Get-DecryptedString
##
##  5/3/2017    1.2.1     Wind      Add parameter 'Status' for Watch-Jobs
##
##  5/14/2017   1.3.0     Wind        Move function 'Get-EncryptedString', 'Get-DecryptedString' and 'Search-DomainUser' out
##                                    to module 'MyUtilities'.
##  
##  5/24/2017   1.3.1     Wind      Update function Send-Email: Change SMTP to EXO.
##                                  Move New-RandomPassword to MyUtilities.
##
##  6/6/2017    1.3.2     Wind      Function Start-SleepProgress: Bug fix for avoid time percent greater than 100.
##
##  6/14/2017   1.3.3     Wind      Function Send-Email: Add parameter 'Attachments'
##
##  6/15/2017   1.3.4     Wind      Function Get-SPDFQDN: add parameter 'first', 'random' and delete 'unique'.
##
##  7/10/2017   1.3.5     Wind      Function Format-HtmlTable: add mandatory for Parameter 'Contents' to avoid null value.
##
##  9/25/2017   1.3.6     Wind      update function: Convert-ArrayAsString
##                                    1. Add parameter 'Quote' with option 'Double', 'Single' and 'None'.
##                                    2. Avoild return error if input is null.
##
##  11/14/2017  1.5.0     Wind     Add function: New-Path
##                                 Move 'Get-EncryptedString', 'Get-DecryptedString' in.
##
##  6/13/2018   1.5.1     Wind     Add function: LoadXML
##  
##  7/16/2018   1.5.2     Wind     Bug fix:
##                                   LoadXML: use provider path in XML load function to avoid UNC path loading error.
##
##  10/22/2018  1.5.3     Wind     function get-SPDFQDN change:
##                                   Now SPD farm have not been distinguished by PR and DR. Delete paramter "PirmaryFarm".
##
##  02/06/2020  1.5.4     Wind     function get-SPDFQDN change:
##                                   1. Change dsfqdn getting logical because new DS farm has new DsFqdn.
##                                   2. Add parameter FarmId.
#####################################################################################
#>

$CommonColor = @{
    Red = "#FF0000"
    Cyan = "#00FFFF"
    Blue = "#0000FF"
    purple = "#800080"
    Yellow = "#FFFF00"
    Lime = "#00FF00"
    Magenta = "#FF00FF"
    White = "#FFFFFF"
    Silver = "#C0C0C0"
    Gray = "#808080"
    Black = "#000000"
    Orange = "#FFA500"
    Brown = "#A52A2A"
    Maroom = "#800000"
    Green = "#008000"
    Olive = "#808000"
    Pink = "#FFC0CB"
    DarkBlue = "#00008B"
    DarkCyan = "#008B8B"
    DarkGray = "#A9A9A9"
    DarkGreen = "#006400"
    DarkMagenta = "#8B008B"
    DarkOrange = "#FF8C00"
    DarkRed = "#8B0000"
    DeepPink = "#FF1493"
    LightBlue = "#ADD8E6"
    LightCyan = "#E0FFFF"
    LightGray = "#D3D3D3"
    LightGreen = "#90EE90"
    LightYellow = "#FFFFE0"
    MediumBlue = "#0000CD"
    MediumPurple = "#9370DB"
    YellowGreen = "#9ACD32"
}

Function Format-StringToHTMLTable {
<#
.SYNOPSIS
Add string array into HTML table.

.DESCRIPTION

.PARAMETER InputString
String array whcih you want to format as HTML table.

.PARAMETER Header
HTML table header which is the field name. It will not add header if omit.

.OUTPUTS
Return a HTML string.

.EXAMPLE
PS C:\> Format-StringToHTMLTable $StringArray 
#>
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String[]] $InputString,
        [String] $Header
    )

    if (!$Header) { $Header = "CustomField"; $withoutHeader = $true }
    
    $objects = @()
    foreach ($instr in $InputString) {
        $obj = New-Object -TypeName PSObject -Property @{$Header=$instr}
        $objects += $obj
    }
    $result = $objects | ConvertTo-Html -Fragment 
    # convertTo-Html will convert only one note property name as "*". It's bug. So replace * as field name.
    if ($withoutHeader) {
        $result = $result[0,1+3..($result.count-1)]
    }
    else { $result = $result -replace "\*",$Header }

    return $result
}

Function Format-HtmlTable {
<#
.SYNOPSIS
Convert objects as a HTML table.

.DESCRIPTION
It can add title before the table and set table border and cell padding and spacing.

.PARAMETER Contents
The objects which will convert to html table.

.PARAMETER Title
Table title added before the table.

.PARAMETER TitleColor
Title font color. Default is 'Black'.

.PARAMETER TitleSize
Title font size. Range is between 8 and 128. Default is 32.

.PARAMETER TableBorder
Table border. Default is 1.

.PARAMETER Cellpadding
Table cell padding. Default is 0.

.PARAMETER Cellspacing
Table cell spacing. Default is 0.

.OUTPUTS
String array.

.EXAMPLE
PS C:\> Format-HtmlTable -Contents $GFE -Title "GFE list"
Format GFE list as a HTML table with title "GFE list".
#>    
    Param (
        # Update: 1.3.5: To avoid content as null.
        [Parameter(Mandatory=$true)]
        [Object[]] $Contents,
        [String] $Title,
        [String] $TitleColor = $CommonColor["Black"],
        [ValidateRange(8,128)]
        [Int] $TitleSize = 32,
        [Int] $TableBorder=1,
        [Int] $Cellpadding=0,
        [Int] $Cellspacing=0
    )
    $TableHeader = "<table border=$TableBorder cellpadding=$Cellpadding cellspacing=$Cellspacing"
    $PreContent = "<H1 style=color:$Titlecolor;font-size:$TitleSize>$Title</H1>"

    # Convert to HTML
    # Update: 1.1.8: convert string to html table.
    if ($Contents[0] -is [string]) {
        $HTMLContent = Format-StringToHTMLTable $Contents
        $HTMLContent = $PreContent + $HTMLContent
    }
    else {
        $HTMLContent = $Contents | ConvertTo-Html -Fragment -PreContent $PreContent
    }
    #Add table border
    $HTMLContent = $HTMLContent -replace "<table", $TableHeader
    # Strong table header
    $HTMLContent = $HTMLContent -replace "<th>", "<th><strong>"
    $HTMLContent = $HTMLContent -replace "</th>", "</strong></th>"
    
    Return $HTMLContent
}

Function Send-Email {
<#
.SYNOPSIS
Send E-mail.

.DESCRIPTION
This function invoke Send-MailMessage to send E-mail.

.PARAMETER To
Recipients. Format is "song.wende@oe.21vianet.com" or "Wind SONG <song.wende@oe.21vianet.com>".

.PARAMETER mailbody
Mail contents. Accept string array.

.PARAMETER mailsubject
Mail subject.

.PARAMETER From
Mail sender. Default is "spo_automail@21vianet.com".

.PARAMETER SmtpServer
Default is "shasmtp-int.core.chinacloudapi.cn".

.PARAMETER BodyAsHtml
Content will treat as html format.

.PARAMETER Attachments
Mail attachments.

.OUTPUTS
None.
#>
    param
    (
        [Parameter(Mandatory=$true)]
        [string[]] $To,
        [Parameter(Mandatory=$true)]
        [string[]] $mailbody,
        [Parameter(Mandatory=$true)]
        [string] $mailsubject,
        [String] $From = "spo_automail@21vianet.com",
        [String[]] $Attachments,
        [String] $SmtpServer,
        [System.Management.Automation.PSCredential] $Credential,
        [Switch] $BodyAsHtml,
        [String] $XmlConfig = "\\chn\tools\DailyCheck\config-dailycheck.xml"
    )

    $Body = $mailbody | Out-String 

    # update: 1.3.1
    $Xml = New-Object -TypeName System.Xml.XmlDocument
    $Xml.PreserveWhitespace = $false
    $Xml.Load($XmlConfig)
    $Config = $xml.DailyCheck.Common

    if (!$SmtpServer) {
        $SmtpServer = $Config.SmtpServer
    }

    $UseSsl = $config.SmtpUseSsl -as [bool]

    if (!$Credential) {
        $Passphrase = $config.Passphrase
        $encryptedpassword = $config.SmtpPassword
        $password = Get-DecryptedString -EncryptedString $encryptedpassword -Sender $Passphrase
        $key = ConvertTo-SecureString -String $password -AsPlainText -Force
        $username = $config.SmtpUserName
        $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username,$key
    }

    $messageParameters = @{
        Subject = $mailsubject
        Body = $body  
        From = $From
        To = $To
        SmtpServer = $SmtpServer
        UseSsl = $UseSsl
        Credential = $Credential
    }

    if ($Attachments) {
        $messageParameters.Add("Attachments", $Attachments)

    }

    $command = "Send-MailMessage @messageParameters"

    If ($BodyAsHtml) {
        $command = $command,"-BodyAsHtml" -join " "
    }

    try {
    Invoke-Expression $command
    }
    catch [System.IO.FileNotFoundException] {
        Write-Host "Cannot find attachments: " -ForegroundColor Red -NoNewline
        Write-Host (Convert-ArrayAsString $Attachments)
        break
    }
}

Function Get-SPDFQDN {
<#
.SYNOPSIS
This function is used for getting DS FQDN or hosts FQDN of SPD.

.DESCRIPTION
Get SPD FQDN has a bit hard. This function offer a convenient method to get it.

.PARAMETER DsFQDNonly
Just return domain service FQDN.

.PARAMETER Random
Get only one SPD FQDN in random.

.PARAMETER Frist
Get number of SPD FQDN from beginning of the SPD array.

.PARAMETER FarmId
SPD farm ID.
#>   


    Param (
        [Parameter(Mandatory=$true,ParameterSetName="DsOnly")]
        [Switch] $DsFQDNonly,
        [Parameter(Mandatory=$true,ParameterSetName="ForHost")]
        [Int] $First,
        [Int] $FarmId,
        [Parameter(ParameterSetName="Random")]
        [Switch] $Random
    )

    if($FarmId) {
        $SPDs = Get-CenteralVM -Role SPD -State Running -Farm $FarmId
    }
    else {
        $SPDs = Get-CenteralVM -Role SPD -State Running
    }

    $Farms = $SPDs | group FarmId | %{Get-CenteralFarm $_.Name}

    if($PSCmdlet.ParameterSetName -eq "DsOnly") {
        return $Farms | group DsFqdn | select -ExpandProperty Name
    }

    
    if ($First) {
        $SPDs = $SPDs | select -First $First
    }

    if ($Random.IsPresent) {
        $SPDs = $SPDs | Get-Random
    }

    $SPDName = @()
    foreach($spd in $SPDs) {
        $farm = $Farms | ? FarmId -eq $spd.FarmId
        $dsfqdn = $farm.DsFqdn
        $SPDName += $spd.Name, $dsfqdn -join '.'
    }

    return $SPDName
}

Function Watch-Jobs {
<#
.SYNOPSIS
Watch PowerShell jobs to finish and return results.

.DESCRIPTION
This function is used to monitor PS backgroud job running.
It will treate the jobs in stats "Finished", "failed" and "Blocked".
"Finished" jobs will be receive returned results and delete jobs.
"Failed" jobs will be receicve jobs and show failed computer name and delete jobs.
Just show warning but don't delete for "Blocked".

.PARAMETER Id
PowerShell job ID.

.PARAMETER Activity
Display in progress bar.

.PARAMETER TimeOut
The limited time for PowerShell job running. Default is 120 seconds.
If time end up, all unfinished jobs will be terminated.

.PARAMETER NoWait
This function will return result immediately once any job finished or failed.

.EXAMPLE
Watch-Jobs -Activity "Watching jobs" -Status "Executing" -TimeOut 3600
Watch all background jobs in 1 hour.

.EXAMPLE
Watch-Jobs -NoWait
Watch all background jobs in 120 seconds. It will end up if one of jobs finished.

.EXAMPLE
Watch-Jobs -Id 2
Just watch job 2 in 120 seconds.
#>
    Param (
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Int[]] $Id,
        [String] $Activity = "Getting PowerShell Job(s) info",
        [String] $Status = "Querying",
        [Int] $TimeOut = 120,
        [Switch] $NoWait
    )

    $results = @()
    $FailedComputers = $()

    If (!$Id) { $command = "Get-Job" }
    Else { $command = "Get-Job -Id $Id -ErrorAction Ignore" }

    while ($jobs = Invoke-Expression $command) {

        $RunningJobs = $jobs | ? State -EQ "Running"
        $CompletedJobs = $jobs | ? State -EQ "Completed"
        $FailedJobs = $jobs | ? State -EQ "Failed"
        $blockedJobs = $jobs | ? State -EQ "Blocked"

        if ($RunningJobs.Count -eq $jobs.Count) {
            $timeout --
            Write-Progress -Activity $Activity -Status "$Status" -CurrentOperation "$($jobs | ? State -EQ "Running" | % {$_.Name})"
            Start-Sleep 1 
            Write-Progress -Activity $Activity -Status "$Status ." -CurrentOperation "$($jobs | ? State -EQ "Running" | % {$_.Name})"
            Start-Sleep 1 
            Write-Progress -Activity $Activity -Status "$Status .." -CurrentOperation "$($jobs | ? State -EQ "Running" | % {$_.Name})"
            Start-Sleep 1 
            Write-Progress -Activity $Activity -Status "$Status ..." -CurrentOperation "$($jobs | ? State -EQ "Running" | % {$_.Name})"
            Start-Sleep 1
        }

        if ($CompletedJobs) {
            $results += Receive-Job $CompletedJobs
            Remove-Job $CompletedJobs
        }

        if ($FailedJobs) {
            $results += Receive-Job $FailedJobs -ErrorAction Ignore
            If ($FailedJobs.ChildJobs) {
                $ChildJobs = $FailedJobs.ChildJobs
                $FailedComputers = $ChildJobs | ? State -EQ "Failed" | Select-Object -ExpandProperty Location
            }
            Remove-Job $FailedJobs
        }

        # Blocked job will not be treated.
        If ($blockedJobs.Count -eq $jobs.Count) {
            Write-Warning "There are some blocked jobs that will not be clean!"
            $blockedJobs
            Break
        }

        # Should be end up if 'NoWait' is present.
        if ($NoWait.IsPresent -and ($CompletedJobs -or $FailedJobs)) { break }

        If ($timeout -lt 1) {
            Write-Progress -Activity "Jobs running time out." -Status "Terminating ..."
            get-job | Remove-Job -Force
            Write-Host "Jobs running time out. Terminated!" -ForegroundColor Red
        }
    }

    If ($FailedComputers) {
        Write-Host "Job can't be run for below machines:"
        $FailedComputers | %{Write-Host $_}
    }

    Return $results
}

Function Start-BatchJobs {
<#
.SYNOPSIS
Kick off local PS jobs in batch mode.

#>

    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [ScriptBlock] $ScriptBlock,
        [Parameter(ParameterSetName="Computer")]
        [String[]] $ComputerName,
        [Parameter(ParameterSetName="Argument")]
        [Object[]] $ArgumentList,
        [AuthenticationMechanism] $Authentication,
        [PSCredential] $Credential,
        [PSObject] $InputObject,
        [String] $Name,
        [Version] $PSVersion,
        [Int] $ActiveThreads = 20,
        [Int] $TimeOut=120 #seconds
    )

    $command = "Start-Job -ScriptBlock `$ScriptBlock"
    #compose command statement
    # if ($ArgumentList) { $command += " -ArgumentList `$ArgumentList" }
    if ($Authentication) { $command += " -Authentication `$Authentication" }
    if ($Credential) { $command += " -Credential `$Credential" }
    if ($InputObject) { $command += " -InputObject `$InputObject" }
    if ($Name) { $command += " -Name `$Name" }
    if ($PSVersion) { $command += " -PSVersion `$PSVersion" }
    
    $Result = @()

    if ($PSCmdlet.ParameterSetName -eq "Computer") {
        if ($ComputerName.Count -gt 1) {
            foreach ($computer in $ComputerName) {
                $expression = $command, "-ArgumentList `$computer"
                Invoke-Expression $expression | Out-Null
                
            }
        }
        else {

        }
    }


}

Function Start-SleepProgress {

<#
.SYNOPSIS
To display progress bar for waiting.

.DESCRIPTION
Start wating with progress bar.

.PARAMETER Seconds
The waiting seconds.

.PARAMETER Activity
Activity shows in progress bar. Default is "Sleeping progress".

.PARAMETER Status
Status shows in progress bar. Default is "Waiting".

#>

    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [Int] $Seconds,
        [String] $Activity = "Sleeping progress",
        [String] $Status = "Waiting"
    )
   
    $EndTime = (Get-Date).AddSeconds($Seconds)

    Do {
        
        $CurTime = Get-Date
        $Remaining = ($EndTime - $CurTime).TotalSeconds
        $Escape = $Seconds - $Remaining
        # Bug fix: 1.3.2
        if ($Escape -gt $Seconds) { $Escape = $Seconds }

        $Percent = $Escape / $Seconds * 100 -as [Int]
        Write-Progress -Activity $Activity -Status $Status -SecondsRemaining $Remaining -PercentComplete $Percent

    } While ($Remaining -gt 0)
    Write-Progress -Activity $Activity -Completed
}

Function Convert-ArrayAsString {
<#
.SYNOPSIS
Convert array string to a string.

.DESCRIPTION
Array output every element in new line.
This function convert array as a string with double quota such as "abc","def" as default.
You can change the quote as single or none with parameter 'Quote'.

.Parameter Array
The input array. It's must can be convert to string type.

.Parameter Quote
The quote mark type include "Single","Double" and "None".
Default is "Double".

.Parameter Separator
The separate mark between two objects. Default is comma.

.Example
Convert-ArrayAsString 1,2,3,4,5
output: "1","2","3","4","5"
Convert integer array as string with double quote.

#>
    Param (
        [Parameter(Position=1)]
        [String[]] $array,
        [String] $Separator=',',
        [ValidateSet("Single","Double","None")]
        [String] $Quote="Double"
    )
    switch ($Quote) {
        "Double" {$QuoteMark = "`""}

        "Single" {$QuoteMark = "'"}

        "None" {$QuoteMark = ""}
    }

    if ($array) {
    $String = $QuoteMark

    if ($array.Count -gt 1) {
        $String += $array -join "$QuoteMark$Separator$QuoteMark"
    }
    else {
        $String += $array -as [String]
    }

    $String += $QuoteMark
    }

    return $String
}

Function Write-Log {
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String] $Path,
        [Parameter(Mandatory=$true,Position=2)]
        [DateTime] $Time,
        [ValidateSet("OK","Warning","Error","Mandatory")]
        [String] $Status = "OK",
        [Parameter(Mandatory=$true,Position=3)]
        [String] $Message,
        [Switch] $Override
    )

    $Logfile = Get-Item $Path

    $Log = $Time.ToString()
    $Log = $Log, $Status -join ", "
    $Log = $Log, $Message -join ", "

    If (!$Logfile -or $Override.IsPresent) {
        $Title = "DateTime,Status,message"
        $Title | Out-File $Logfile
    }

    $Log | Out-File $Logfile -Append
}

Function Get-FileName {
<#
.SYNOPSIS
Generate a serial file names.

.DESCRIPTION
This function is used for providing a serial names for batch. Generate full path is allowed if you run it with parameter 'Path'.

.PARAMETER Path
File path added into file name.

.PARAMETER Name
File name without extesion.

.PARAMETER FileType
Thie file name extesion. Default is "log".

.PARAMETER StartNumber
The first index number in file name. It must be less than 'EndNumber'.

.PARAMETER EndNumber
The last index number in file name. It must be greater than 'StartNumber'.

.PARAMETER Amount
The amount of file names. Default is 1.

.PARAMETER NumberFormat
How many digits represent number. Default is "XX".

.PARAMETER WithDateTime
File name with date time formated as "yyyyMMddhhmmss".

.EXAMPLE
PS C:\>
#>
    Param (
        [String] $Path,
        [Parameter(Mandatory=$true,Position=0)]
        [String] $Name,
        [String] $FileType = "log",
        [Parameter(Mandatory=$true,ParameterSetName="Range",Position=1)]
        [Int] $StartNumber,
        [Parameter(Mandatory=$true,ParameterSetName="Range",Position=2)]
        [ValidateScript({$_ -gt $StartNumber})]
        [Int] $EndNumber,
        [Parameter(ParameterSetName="Batch")]
        [Int] $Amount = 1,
        [Validateset("X","XX","XXX")]
        [String] $NumberFormat = "XX",
        [Switch] $WithDateTime
    )

    # Convert amount to numbers.
    If ($PSCmdlet.ParameterSetName -eq "Batch") {
        $Start = 1
        $End = $Amount
    }

    # Concatenate path and file name.
    If ($Path) {
        $Name = Join-Path $Path $Name
    }

    Switch ($NumberFormat) {
        "X" {$format = "D"}
        "XX" {$format = "D2"}
        "XXX" {$format = "D3"}
    }

    If ($WithDateTime.IsPresent) {
        $date = Get-Date -Format "yyyyMMddhhmmss"
        $Name = $Name, $date -join "_"
    }
    
    $NameArray = @()
    For ($i=$Start;$i -le $End;$i++) {
        $SerialNumber = $i.ToString($format)
        $n = $Name, $SerialNumber -join "_"
        $n = $n, $FileType -join "."
        $NameArray += $n
    }

    Return $NameArray
}

Function Get-MaxLengthString {

<#
.SYNOPSIS
Get the string which is the longest in array.

.DESCRIPTION
Return string which is the longest in array.

.PARAMETER InputObject
String array.

.EXAMPLE
PS C:\> $array = "You", "are", "my", "love"
PS C:\> Get-StringMaxLength $array
Return "love".
#>

    Param (
        [String[]] $InputObject
    )
    $maxString = ""

    foreach ($object in $InputObject) {
        If ($object.Length -gt $maxString.Length) {
            $maxString = $object
        }
    }

    Return $maxString
}

Function Get-StringMaxLength {

<#
.SYNOPSIS
Get the length of the longest string from array.

.DESCRIPTION
To return the longest string length from string array.

.PARAMETER InputObject
String array.

.EXAMPLE
PS C:\> $array = "You", "are", "my", "love"
PS C:\> Get-StringMaxLength $array
Return 4 which is the length of "love".
#>

    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String[]] $InputObject
    )

    $maxString = Get-MaxLengthString -InputObject $InputObject

    Return $maxString.Length
}

function New-Path {
<#
.SYNOPSIS
Create file or directory path in recurse.

.DESCRIPTION
New-Item just create single file or directory. This function will create target items along with the path.
Such as, the directory 'test' and file 'temp.txt' are not exist in the path "c:\test\temp.txt".
You can create them at once by "New-Path c:\test\temp.txt".

.PARAMETER Path
The full path which you want to create.

.PARAMETER PassThru
Return created items if this parameter is present.

.PARAMETER Type
The final item type in the path.

.EXAMPLE
New-Path c:\test\abc\temp.txt
Create directories 'test', 'abc' and file 'temp.txt' if they are not exist.

.EXAMPLE
New-Path c:\test\abc\temp.txt -Type Directory
create directories 'test', 'abc' and 'temp.txt'.

.EXAMPLE
New-Path c:\test\abc\temp.txt -PassThru
Create directories 'test', 'abc' and file 'temp.txt' and return them.

#>
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [String] $Path,
        [ValidateSet("File","Directory")]
        [String] $Type = "File",
        [Switch] $PassThru
    )

    $stack = New-Object -TypeName System.Collections.Stack
    $result = @()

    while ($Path[-2] -ne ":") {
        if (!(Test-Path $Path)) {
            $stack.Push($Path)
            $Path = Split-Path $Path -Parent
        }
        else { break }
    }

    $t = "Directory"
    while ($stack.Count) {
        $p = $stack.Pop()
        if ($stack.Count -eq 0) { $t = $Type }
        try {
            $result = New-Item -ItemType $t -Path $p -ErrorAction Stop
        }
        catch {
            Write-Host "[function New-Path] Cannot create path [$p]" -ForegroundColor Red
            Write-Host $_
            break
        }
    }

    if ($PassThru.IsPresent) { return $result }

}

Function Get-EncryptedString {
<#
.SYNOPSIS
Convert plan text to encrypted string.

.DESCRIPTION
This function is used to encrypt plan string.
 
.PARAMETER String
The plan text you want to encrypt.

.EXAMPLE 
PS C:\> Get-EncryptedString Passw0rd!
It will return encrypt key for text 'Passw0rd!'.
#>
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String] $String
    )

    # Get current user as passphrase.
    $user = $env:USERNAME
    $ADUser = Get-ADUser $user
    $passphrase = $ADUser.UserPrincipalName

    $key = Encrypt-String -String $String -Passphrase $passphrase

    Return $key
}

Function Get-DecryptedString {
<#
.SYNOPSIS
Decrypt secret string.

.DESCRIPTION
This function is used to decrypt the encryped string as plan text.
You must provide 'sender' who encrypted it. This function will use sender's UPN as passpharse to decrypt secret.
'Sender' accept fuzzy query but SAM account is preferred.

.PARAMETER EncryptedString
Encrypted string which you want to decrypt.

.PARAMETER Sender
The person who encrypted the string. We use the sender's UPN to encrypt string. SAM account name is preferred.
This parameter accept fuzzy search too. You just provide the sender name, it will return AD user list for you to select which AD account as passpharse.

.EXAMPLE
Gallatin PS D:\userdata\oe-songwende> Get-DecryptedString $key -Sender oe-songwende
'oe-songwende' is SAM account name. 

.EXAMPLE
Gallatin PS D:\userdata\oe-songwende> Get-DecryptedString $key -Sender Wende

Number Name               SamAccountName     DisplayName        UserPrincipalName
------ ----               --------------     -----------        -----------------
     1 Wende Song         oe-songwende       Wende Song         oe-songwende@CHN.SPONETWORK.COM
     2 Wende SONG - Admin oe-songwende_admin Wende SONG - Admin oe-songwende_admin@CHN.SPONETWORK.COM


Select user number in above list which you want to use: 1

Password!123
Gallatin PS D:\userdata\oe-songwende>
#>
    Param (
        [Parameter(Mandatory=$true,Position=1)]
        [String] $EncryptedString,
        [Parameter(Mandatory=$true)]
        [String] $Sender
    )

    # Find sender out from AD.
    $ADuser = @(Search-DomainUser -Name $Sender)
    $ADuser = $ADuser | Select-Object -Property *Name

    # Pick up 1 account.
    If ($ADuser.count -gt 1) {
        $i = 1
        foreach ($u in $ADuser) {
            $u | Add-Member -NotePropertyName Number -NotePropertyValue $i
            $i++
        }
        $ADuser | Format-Table -AutoSize -Property Number,Name,SamAccountName,DisplayName,UserPrincipalName | Out-Host
        $num = (Read-Host -Prompt "Select user number in above list which you want to use") -as [Int]
        if (!$num) {
            Write-Host "Cannot accept non number input!" -ForegroundColor Red
            Break
        }

        $ADuser = $ADuser | ? Number -EQ $num
    }

    If (!$ADuser) {
        Write-Host "Cannot find sender!" -ForegroundColor Red
        Break
    }

    $passphrase = $ADuser.UserPrincipalName

    $password = Decrypt-String -Encrypted $EncryptedString -Passphrase $passphrase

    Return $password
}

Function LoadXML {
<#
.SYNOPSIS
To read XML file and return XMLDocument object.

#>
    [Cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true,Position=0)]
        [String] $Path
    )

    $FullPath = Resolve-Path $Path -ErrorAction Ignore
    if(!$FullPath) {
        Write-Host "The path `"$Path`" does not exist!" -ForegroundColor Red
        break
    }

    $xml = New-Object System.Xml.XmlDocument
    try {
        $xml.Load($FullPath.ProviderPath)
    }
    catch {
        Write-Host $_ -ForegroundColor Red
        break
    }
    
    return $xml

}