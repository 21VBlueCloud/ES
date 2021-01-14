<#
######################################################################################
#
#File: FlightFunctions.ps1
#Author: chen.xiaoyue@oe.21vianet.com (Maggie Chen)
#Version: 3.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##  9/9/2017   1.0       Xiaoyue      First mayjor version
##                                    
##  9/14/2017  2.0       Xiaoyue      Add start-job to compare Flights and CenteralFlights 
##                                    at same time to short the executing time
##  9/26/2017  3.0       Xiaoyue      Add filter unctions
#
######################################################################################
#>

Function Export-AllFlights {

<#
.SYNOPSIS
This function for export the Flights and CenteralFlights into the folder named by date 
under \\chn\tools\Flights.

.DESCRIPTION 
If the type is not be specified, it will save both Flights and CenteralFlights.

.PARAMETER Types
"OnlyFlights","OnlyCenteralFlights" or "All"

.EXAMPLE
Export-AllFlights

.EXAMPLE
Export-AllFlights -Types "All"

.EXAMPLE
Export-AllFlights -Types "OnlyCenteralFlights"

#>
    
    Param ( 
        [ValidateSet("OnlyFlights","OnlyCenteralFlights","All")]
        [String[]] $Types="All"
    )
    
    Write-Debug "Flights are retrieving ..."

    $OnlyFlights = $OnlyCenteralFlights= @()

    If ($Types -eq "All") {
        $Types = @("OnlyFlights","OnlyCenteralFlights")
    }

    $AllFlights = @{}

    Switch ($Types) {

        "OnlyFlights" {
            $OnlyFlights = Get-Flight -WarningAction Ignore
            $AllFlights.Add("OnlyFlights",$OnlyFlights)
        }

        "OnlyCenteralFlights" {
            $OnlyCenteralFlights = Get-CenteralFlight
            $AllFlights.Add("OnlyCenteralFlights",$OnlyCenteralFlights)
        }
    }

    $ExportPath = "\\CHN.SPONETWORK.COM\Tools\Flights"
    If (!(Test-Path $ExportPath)) {

        Write-Debug "Can not find $ExportPath" -ForegroundColor Red
        Return $AllFlights
    }

    $Date = Get-Date
    $StrDate = Get-Date -Date $Date -Format "yyyyMMdd"
    $StrDateTime = Get-Date -Date $Date -Format "yyyyMMddhhmmss"
    $StoreFolder = Join-Path $ExportPath $StrDate

    If (!(Test-Path $StoreFolder)) {
        try { New-Item -Path $StoreFolder -ItemType Directory | Out-Null }
        catch {
            Write-Debug "Can not create folder $StoreFolder!"
            Return $AllFlights
        }
    }

    foreach ($Item in @($AllFlights.Keys)) {
        $xml = $Item + "_" + $StrDateTime + ".xml"
        $Path = Join-Path $StoreFolder $xml
        $AllFlights[$Item] | Export-Clixml -Path $Path
        If (Test-Path $Path) {
            Write-Host "Export: '$Path'" -ForegroundColor Green
        }
    }

    Write-Debug "Flights data retrieving done."
    Return $AllFlights
}

Function Get-FlightVersions {

<#
.SYNOPSIS
Function for returning all of versions of the CHN Flight stored.

.DESCRIPTION

If you want to get Flights info from another path, please use parameter 'StorePath'.
Default path is "\\CHN.SPONETWORK.COM\Tools\Flights".

.PARAMETER Latest
Value of "Latest" define the number of latest selected versions. The version was ordered descendingly by default  

.PARAMETER StorePath
Define the place where it get flight versions information from.

.EXAMPLE
Get-FlightVersions -Latest 8
Get the latest 8 versions

#>
    
    Param (
       [Parameter(Mandatory=$true)]
       [ValidateScript({$_ -gt 0})][Int] $Latest,
       [String]$StorePath="\\CHN.SPONETWORK.COM\Tools\Flights"
    )

    Write-Debug "Getting  Flights' version information ..."

    $XmlFiles = Get-ChildItem -Path $StorePath -file -Recurse
    $FilesName = $XmlFiles.BaseName
    $Versions = $FilesName -split '_' | ? {$_ -Match "^\d+"} | Select-Object -Unique | Sort-Object
    
    $ObjVer = @()
    $Order = 0
    
    foreach ( $v in $Versions) {

        $Obj = New-Object -TypeName PsObject
        $Order ++
        $Obj | Add-Member -NotePropertyName Order -NotePropertyValue $Order
        $Obj | Add-Member -NotePropertyName Version -NotePropertyValue $v
        $thisVersionFileNames = $FilesName | ?{ $_ -match $v } 
        $types = $thisVersionFileNames -split '_' | ? {$_ -notmatch "\d+"} 
        $Obj | Add-Member -NotePropertyName Types -NotePropertyValue $types 
        $ObjVer += $Obj
        
    }

    $ObjVer = $ObjVer | Select-Object -Last $Latest

    Remove-Variable Order

    Return $ObjVer
}

Function Get-sepcificFlightFile{

<#
.SYNOPSIS
Choose sepcified Flight version and return the Flight Files's full path.

.DESCRIPTION
Choose sepcified Flight version and return the Flight Files's full path.

.PARAMETER Versions
Define the optional version and make user choose it from. 

.PARAMETER Prompting
Prompt a text to guide the user input the order of the version to choose it.

.PARAMETER StorePath
Define the place where it get flight files from.

.EXAMPLE
Get-sepcificFlightFile -versions $versions -prompting $prompting

#>
    
    Param (
        [CmdletBinding()]
        [Parameter(Mandatory=$true)]
        [Object]$versions,
        [string]$Prompting,
        [string]$StorePath="\\CHN.SPONETWORK.COM\Tools\Flights"
    )

    Out-Host -InputObject $versions
    $Order = Read-Host $prompting
    $object = $versions|?{$_.Order -eq $Order}
    $result=@()

    If (!$object) {
      Write-Host "CANNOT get the data you choose" -ForegroundColor Red
    }

    $StrToday=Get-Date -Format yyyyMMdd
    If ($object.Version -match $StrToday){

      write-host "This version was exported today. " -ForegroundColor Yellow
      $confirmFlag = Read-Host "Are you sure to compare Flights state with it?(Y/N)" 

      if($confirmFlag -eq 'Y'){
         $version = $object.Version
      }

      elseif($confirmFlag -eq 'N'){
         Get-sepcificFlightFile -versions $versions -prompting $prompting
      }

      else{
        write-host 'Please input "Y" or "N" to re-choose version' -ForegroundColor Red
        Get-sepcificFlightFile -versions $versions -prompting $prompting
      }
    }

    Else{
      $version=$object.Version
    }

    Write-Debug "Getting Latest Flight Version done..." 
    $result = Get-ChildItem -Path $StorePath -file -Recurse |?{$_.Name -match $version}
    return $result
}

Function Import-FlightXmlData {

<#
.SYNOPSIS
Convert the XML file into objects

.DESCRIPTION
Convert the XML file into objects

.PARAMETER Types
"All","OnlyFlights" or "OnlyCenteralFlights".It define which kind of Flights it return.

.PARAMETER XmlFiles
The Data of XML file

.EXAMPLE
Get-sepcificFlightFile -versions $versions -prompting $prompting

#>

    Param (
        [ValidateSet("OnlyFlights","OnlyCenteralFlights","All")]
        [String[]] $Types="All",
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo[]] $XmlFiles
    )

    Write-Debug "Starting import reference Flights information ..."

    If ($Types -eq "All") {
        $Types = @("OnlyFlights","OnlyCenteralFlights")
    }

    $Flights =@{}

    Switch ($Types) {

        "OnlyFlights" {
            $OnlyFlights = $XmlFiles | ? BaseName -Match "^OnlyFlights" | Import-Clixml
            $Flights.Add($_,$OnlyFlights)
        }

        "OnlyCenteralFlights" {
            $OnlyCenteralFlights = $XmlFiles | ? BaseName -Match "^OnlyCenteralFlights" | Import-Clixml
            $Flights.Add($_,$OnlyCenteralFlights)
        }
    }

    Write-Debug "Import done."

    Return $Flights
}

Function Compare-FlightsMain {
<#
.SYNOPSIS
The main function of flights comparing

.DESCRIPTION
The main function of flights comparing

.PARAMETER Types
Define the type to compare. If it's "All", it means the both Flights and CenteralFlights are compared.

.PARAMETER CompareProperty
It will compare its properties in detail if it's enabled.

.PARAMETER MailTo
The email address to receive the result.

.EXAMPLE
Compare-FlightsMain  -MailTo chen.xiaoyue@oe.21vianet.com -CompareProperty
Compare all flights and their properties in detail then send the result to sepecified person.

.EXAMPLE
Compare-FlightsMain  -MailTo chen.xiaoyue@oe.21vianet.com
Only find the new or deleted Flights and CenteralFlights.

.EXAMPLE
Compare-FlightsMain  -Types "OnlyFlights" -CompareProperty
Just compare Flights and its properties in detail.

#>
    Param (

        [ValidateSet("OnlyFlights","OnlyCenteralFlights","All")]
        [String[]] $Types = "All",
        [Switch]$CompareProperty,
        [String[]] $MailTo
    )

    $Results=@{}
    [String]$MailBody

    #region initialize data
    If ($Types -eq "All") {
       $Types = @("OnlyFlights","OnlyCenteralFlights")
    }

    Write-Host "Start to compare " -NoNewline
    Write-Host "`"$($Types -join '","')`"" -ForegroundColor Green
    
    $versions = Get-FlightVersions -Latest 8
    $StrToday = Get-Date -Format yyyyMMdd
    If(!($versions -match $StrToday)){

       Export-AllFlights;
       $versions = Get-FlightVersions -Latest 8
    }

    $ReferXmls = Get-sepcificFlightFile -versions $versions -prompting "Please input ('Order' number) above to select the Reference objects"
    Write-Host "Importing reference flights..."
    $ReferObjects=Import-FlightXmlData -XmlFiles $ReferXmls
    
    $DiffXmls = Get-sepcificFlightFile -versions $versions -prompting "Now select the Difference objects"
    Write-Host "Getting current flights..."
    $DiffObjects=Import-FlightXmlData -XmlFiles $DiffXmls
    #endregion

    #region ValidateXML
    If (!$ReferObjects) {
        write-host "Cannot find reference record" -ForegroundColor Red
        Break}

    If (!$DiffObjects) {
        write-host "Cannot find Difference record" -ForegroundColor Red
        Break}
    #endregion

    #region ValidateType
    # Need to identify loaded Reference xml files include Types.
    $FileNames = $ReferXmls.Name
    $Flag = $true
    $AbsentTypes = @()
    foreach ($type in $Types) {
        $Matched = $FileNames -match $type
        If (!$Matched) {
            $Flag = $false
            $AbsentTypes += $type
        }
    }

    If (!$Flag) {
        Write-Host "Types " -NoNewline
        Write-Host "`"$($AbsentTypes -join '","')`" " -ForegroundColor Red -NoNewline
        Write-Host "cannot be found in reference objects$($FileNames -join ',')"
        Break
    }

    # Need to identify loaded Difference xml files include Types.
    $FileNames = $ReferXmls.Name
    $Flag = $true
    $AbsentTypes = @()
    foreach ($type in $Types) {
        $Matched = $FileNames -match $type
        If (!$Matched) {
            $Flag = $false
            $AbsentTypes += $type
        }
    }

    If (!$Flag) {
        Write-Host "Types " -NoNewline
        Write-Host "`"$($AbsentTypes -join '","')`" " -ForegroundColor Red -NoNewline
        Write-Host "cannot be found in Difference objects$($FileNames -join ',')"
        Break
    }
    #endregion

    #region ScriptBlock
    $ComFlt={
            
            $FlightRst = Compare-Flights $args[0] $args[1] -CompareProperty:$args[2]

            $FlightRst=$FlightRst|?{$_.record -notmatch 'Same'}
            $Body1 = Format-HtmlTable -Contents $FlightRst -Title "Flights Changes" -ColorItems|out-string
            return $Body1
    }

    $ComGFlt={
            
            $GFlightRst = Compare-CenteralFlights $args[0] $args[1] -CompareProperty:$args[2]

            $GFlightRst=$GFlightRst|?{$_.record -notmatch 'Same'}
            $Body2 = Format-HtmlTable -Contents $GFlightRst -Title "CenteralFlights Changes" -ColorItems|out-string
            return $Body2
    }
    #endregion

    #Get-Job|Remove-Job -ErrorAction SilentlyContinue

    If($Types -match 'OnlyFlights') {
       
       Write-Host "Comparing Flights..."
       $job1 = Start-Job -Name 'CompareFlights' -InitializationScript $Imported_functions -ScriptBlock $ComFlt -ArgumentList $ReferObjects['OnlyFlights'],$DiffObjects['OnlyFlights'],$CompareProperty
    }

    If($Types -match 'OnlyCenteralFlights') {
    
       Write-Host "Comparing CenteralFlights..."
       $job2 = Start-Job -Name 'CompareCenteralFlights' -InitializationScript $Imported_functions -ScriptBlock $ComGFlt -ArgumentList $ReferObjects['OnlyCenteralFlights'],$DiffObjects['OnlyCenteralFlights'],$CompareProperty
    }
    
    $alljobs = $job1,$job2|Wait-Job -ErrorAction SilentlyContinue
    $result =$alljobs|Receive-Job
    $alljobs|Remove-Job

    $MailBody = $result|Out-String

    If (!($MailBody -match "New"  -or $MailBody -match "Deleted")) {

        Write-Warning -Message "No new/deleted Flight/CenteralFlight in version $ReferenceXmls"
        $MailBody = "<h1 Style=`"color:Green`">No new or deleted Flight!</h1>"
    }
    Else {
        Write-Host "Comparing completed!" -ForegroundColor Green
    }

    $date=Get-Date -Format MMMddHHmm
    $MailBody|Out-File -FilePath \\chn\tools\Flights\$date.html

    If ($MailTo) {

        Write-Host "Sending mail to $MailTo" -ForegroundColor Green

        $Subject="Compare Flights $ReferXmls with $DiffXmls"
        $SmtpServer="shasmtp-int.core.chinacloudapi.cn"
        $From = "FlightCompare@21vianet.com"

        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $MailTo -Subject $Subject -Body $MailBody -BodyAsHtml
    }

    Else {
        Write-Host "Didn't send result(\\chn\tools\Flights\$date.html) as mail! If you want, please re-run this script with parameter 'MailTo'." -ForegroundColor Yellow
    }    

    Return $Results
}

Function Initialize-FlightFilter{
<#
.SYNOPSIS
Convert the Flight filter string user input into Powershell Select string.
Such as:Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')
Convert into: $_.Name -like '*Client*' -and ($_.team -eq 'WEX' -or $_.team -like '*snap*')

.DESCRIPTION
Covert the flight filter string

.PARAMETER FilterString
Receive the filter string user input

.EXAMPLE
$Filterstring= "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')"
Initialize-FlightFilter -filterString $Filterstring
#>    
    Param (

       [String] $FilterString
       )

    $subStrs = ($filterString -split '-and') -split '-or'

    Foreach ($subStr in $subStrs)
    {
      $conditStr = $subStr.Trim().TrimStart("(").TrimEnd(")")
      $conditPara = ($conditStr -split " ")[0]

      Switch($conditPara){
      "ID"{
         
         $replaceStr = "`$_.Flight"+$conditStr

      }

      "Name"{

         $replaceStr = "`$_."+$conditStr
      
      }

      "State"{

         $replaceStr = "`$_.Flight"+$conditStr

      }

      "Team"{
        
        $replaceStr = "`$_."+$conditStr
      }

      "Alias"{

        $conditVal = ($conditStr -split " ")[2].Trim("'")
        $conditOper = ($conditStr -split " ")[1].Trim("-")
        
        if($conditOper -eq 'eq'){
        
           $replaceStr = "(`$_.CreatedBy -eq '$conditVal' -or `$_.UpdatedBy -eq '$conditVal' -or `$_.NotificationAliases -eq '$conditVal')"
           
           }
        else{
           
           $replaceStr = "(`$_.CreatedBy -like '$conditVal' -or `$_.UpdatedBy -like '$conditVal' -or `$_.NotificationAliases -like '$conditVal')"
           
           }
      
      }
      
      default {"The condition parameter is incorrect"}
      
      }

    #$filterString = $filterString -replace $conditStr,$replaceStr
    $filterString = $filterString.Replace($conditStr,$replaceStr)

    }

    Return $filterString

}

Function Initialize-CenteralFlightFilter{
<#
.SYNOPSIS
Convert the CenteralFlight filter string which user input into Powershell Select string.
Such as:Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')
Convert into: $_.Name -like '*Client*' 
Because there is no 'Alias' and 'Team' property for CenteralFlight 

.DESCRIPTION
Covert the flight filter string

.PARAMETER FilterString
Receive the filter string user input

.EXAMPLE
$Filterstring= "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')"
Initialize-CenteralFlightFilter -filterString $Filterstring
#>
    
    Param (

       [String] $filterString
       )

    $subStrs = ($filterString -split '-and') -split '-or'

    Foreach ($subStr in $subStrs)
    {
      $conditStr = $subStr.Trim().TrimStart("(").TrimEnd(")")
      $conditSub = ($conditStr -split " ")[0]

      Switch($conditSub){
      "ID"{
         
         $replaceStr = "`$_."+$conditStr
         $filterString = $filterString.Replace($conditStr,$replaceStr)
      }

      "Name"{

         $replaceStr = "`$_."+$conditStr
         $filterString = $filterString.Replace($conditStr,$replaceStr)
      }

      "State"{
         
         $val = ($conditStr -split " ")[2].Trim("'")
         if ($val -eq 'Active' -or $val -eq 'Created' -or $val -eq 'Enabled' -or $val -eq 'true'){

            $replaceStr = "`$_.ConfigContent -match 'Enabled = ture'"

            }
         else{

            $replaceStr = "`$_.ConfigContent -match 'Enabled = false'"

            }
          
         $filterString = $filterString.Replace($conditStr,$replaceStr)
      }

      "Team"{
        
         $replaceStr = "EMPTY"
         $filterString = $filterString.Replace($subStr,$replaceStr)
      }

      "Alias"{

         $replaceStr = "EMPTY"
         $filterString = $filterString.Replace($subStr,$replaceStr)
      
      }
      
      default {"The condition parameter is incorrect"}
      
    }
  }

  $filterString = $filterString.Replace("-andEMPTY","").Replace("-orEMPTY","")
  Return $filterString

}

Function Select-Flights {

<#
.SYNOPSIS
Filter the Flights and CenteralFlights by the fliter string

.DESCRIPTION
Filter the Flights and CenteralFlights by the fliter string

.PARAMETER Types
Its value is "OnlyFlights","OnlyCenteralFlights","All". Default value is "All". Specifiy the type if you want to query only Flights or CenteralFlights

.PARAMETER Filter
This is query string. Its format like AD filter. It support parameter ID,Name,State,Team,Alias
Get-CenteralFlight didn't have Team or alias property. It will skip conditions related when querying CenteralFlight

.PARAMETER Includehistory
This is a switch. It will only query today's files if it's switch off. It will query all history files when it's enabled.

.PARAMETER Mail
Who will receive query result 

.PARAMETER StorePath
Where the source files and results store. The default path is "\\CHN.SPONETWORK.COM\Tools\Flights"

.EXAMPLE
Select-Flights -filter "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*') -and alias -like '*ma*'" -includehistory -MailTo chen.xiaoyue@oe.21vianet.com

.EXAMPLE
Select-Flights -filter "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')" -includehistory -MailTo chen.xiaoyue@oe.21vianet.com

.EXAMPLE
Select-Flights -filter "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')"  -MailTo chen.xiaoyue@oe.21vianet.com  

.EXAMPLE
Select-Flights -Types OnlyFlights -filter "Name -like '*Client*' -and (team -eq 'WEX' -or team -like '*snap*')" -includehistory

.EXAMPLE
Select-Flights -Types OnlyFlights -filter "Id -eq '22'" -includehistory -MailTo chen.xiaoyue@oe.21vianet.com 
#>
    
    Param (

       [ValidateSet("OnlyFlights","OnlyCenteralFlights","All")]
       [String[]] $Types = "All",
       [String]$Filter,
       [Switch]$Includehistory,
       [String[]] $MailTo,
       [String]$StorePath="\\CHN.SPONETWORK.COM\Tools\Flights"
    )
    
    $result=@()
    $BodyPart=@()
    $BodyPart = "<p>"

    If ($Types -eq "All") {
       $Types = @("OnlyFlights","OnlyCenteralFlights")
    }

    Write-Host "Importing reference flight files..."

    $XmlFiles = Get-ChildItem -Path $StorePath -file -Recurse
    $FilesName = $XmlFiles.BaseName
    
    $StrToday = Get-Date -Format yyyyMMdd
    If(!($FilesName -match $StrToday)){

       Export-AllFlights;}

    #region select Flight files
    Switch($Types){

       "OnlyFlights"{

            $FlightXmlFiles = Get-ChildItem -Path $StorePath -file -Recurse|?{$_.Name -match 'OnlyFlights_*'}

            If(!$includehistory){
    
               $FlightXmlFiles = $FlightXmlFiles[$FlightXmlFiles.Count -1]
            }
            Foreach ($Fl in $FlightXmlFiles){
            
            $result += "######################################################################################################################"
            $result += "##############################################"+$Fl.BaseName+"##############################################"
            $result += "######################################################################################################################"
            
            $selectstr1 = Initialize-FlightFilter -filterString $filter
            $FilterString1 = "`$Fl|Import-Clixml|?{$selectstr1}|out-string"
            $FilterRes1 = Invoke-Expression $FilterString1
            $result += $FilterRes1
            
            }
            $BodyPart += $FlightXmlFiles.Name -join "</p><p>"
            $BodyPart += "</p>"
       }

       "OnlyCenteralFlights"{

            $CenteralFlightXmlFiles = Get-ChildItem -Path $StorePath -file -Recurse|?{$_.Name -match 'OnlyCenteralFlights_*'}

            If(!$includehistory){
            
               $CenteralFlightXmlFiles = $CenteralFlightXmlFiles[$CenteralFlightXmlFiles.Count -1]
            
            }
            Foreach ($GFl in $CenteralFlightXmlFiles){
            
            $result += "######################################################################################################################"
            $result += "#############################################"+$GFl.BaseName+"###########################################"
            $result += "######################################################################################################################"
            
            $selectstr2 = Initialize-CenteralFlightFilter -filterString $filter
            $FilterString2 = "`$GFl|Import-Clixml|?{$selectstr2}|out-string"
            $FilterRes2 = Invoke-Expression $FilterString2
            $result+=$FilterRes2

            }
            $BodyPart += $CenteralFlightXmlFiles.Name -join "</p><p>"
       }
    }

    $BodyPart += "</p>"

    $date = Get-Date -Format MMMddHHmm
    $fullpath = Join-Path $StorePath "FilterResult_$date.txt"
    $result> $fullpath

    #endregion

    #region deal with result
    If ($MailTo) {

        Write-Host "Sending mail to $MailTo" -ForegroundColor Green

        If ($includehistory){

            $Subject = "Filter "
            $Subject += $Types -join ','
            $Subject += " files including history record"

            $MailBody = "<h2>Filter Flight from:</h2>"+"$BodyPart"
            $MailBody += "<h2>Please find the result attached</h2>"
            
            }

        Else {
            
            $Subject = "Only Filter today's " 
            $Subject += " files $FlightXmlFiles.Name,$CenteralFlightXmlFiles.Name"
            $MailBody = "<h2>Only searched today's files. Add parameter '-Includehistory' if you want to query history files </h2>"
            $MailBody += "<h2>Please find the result attached</h2>"
            
            }

        $SmtpServer = "shasmtp-int.core.chinacloudapi.cn"
        $From = "FlightFilter@21vianet.com"
        Send-MailMessage -SmtpServer $SmtpServer -From $From -To $MailTo -Subject $Subject -Body $MailBody -BodyAsHtml -Attachments $fullpath
    }

    Else {

        Write-Host "Didn't send result($fullpath) as mail! If you want, please re-run this script with parameter 'MailTo'." -ForegroundColor Yellow
    }
    #endregion
}