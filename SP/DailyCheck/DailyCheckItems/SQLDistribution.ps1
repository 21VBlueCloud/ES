#####################################################################################
#
#File: SQLDistribution.ps1
#Author: JianKun XU
#Version: 1.0
#
##  Revision History:
##  Date        Version    Alias       Reason for change
##  --------    -------   --------    ---------------------------------------
##  5/8/2018    1.0        JK          First version.
##     
##  10/25/2018  1.1        Wind        Update SQL Farm getting method.
##
##  10/25/2018  1.2        Wind        Add SQL mirror check for PR farm.
##
##  11/1/2018   1.3        Wind        Provide the second method to evaluate SQL distribution
##                                     in case of Get-CenteralSQLDBInfo failed.
##
##  11/22/2018  1.4        Wind        Bug fix for 1.3:
##                                       It will execute 2 plans if the last SQL return 0 DB.
##                                       Set plan flag to choice plan.
##
##  1/21/2019   1.5        Wind        Bug fix: Add DataCenter module first to avoid Get-CenteralSqlDBInfo failure.
##
##  11/20/2019  1.6        Jun         Bug fix: Add Multipe PR logic and format the output
#####################################################################################

<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER

.EXAMPLE
PS C:\>
#>

Param (
    [Parameter(Mandatory=$true)]
    [String] $xmlConfig,
    [String[]] $Module,
    [String[]] $Helper    
)

#Pre-loading
##==============================================================================================================
# Import module
if ($Module) {
    Import-Module $Module
}

## Update 1.5: Add DataCenter module first to avoid Get-CenteralSqlDBInfo failure.
Import-Module "C:\Centeral\CenteralManager\dcapi\Microsoft.Sharepoint.Datacenter.dll"


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
$HtmlBody = "<TABLE Class='SQLDISTRIBUTION' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"

##===============================================================================================================

Write-Host $separator
Write-Host "Start to check SQL distribution." -ForegroundColor Green

[Int]$DefaultWarning = $Xml.DailyCheck.SQLDistribution.Threshold.DefaultWarning
[Int]$DefaultAlert = $Xml.DailyCheck.SQLDistribution.Threshold.DefaultAlert

$HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='center'>SQLDistribution</td></TR>"

## Update 1.1: Update SQL farm getting method
$ContentFarm = Get-CenteralFarm -Role Content


$PRContentFarm = $ContentFarm | ? RecoveryFarmId -NE 0 | ? State -EQ 'Active' # filter Active


#$SQLFarm = Get-CenteralFarm -Identity $PRContentFarm.SqlFarmId

$PRContentFarm | % { 

    Write-Host "Farm: $($_.FarmID) SQLFarm:$($_.SQlFarmID)" -ForegroundColor yellow

    $HtmlBody += "<TR style=background-color:#66CCFF;font-weight:bold;font-size:16px;color:#FAF4FF><td colspan='6' align='left'>SQLFarm: $($_.SqlFarmId)</td></TR>"
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightGray']);font-weight:bold><TD colspan='3'>ServerName</TD><TD colspan='3'>PrincipalDB</TD></TR>"

    $SqlVMs = Get-CenteralVM -Farm $($_.SqlFarmId) -Role SQL | ? PMachineid -ne -1 | Select-Object -ExpandProperty Name
    ## Update 1.4: Add Plan flag for choice plan.
    $PlanFlag = 1

    # $sqlvms = Get-CenteralVM -Farm 642 -Role SQL | ? PMachineid -ne -1 | Select-Object -ExpandProperty Name

    foreach ($sqlvm in $SqlVMs)
    {
        $psqldistribution = Get-CenteralSqlDbInfo -SqlVm $sqlvm | ? MirrorRole -eq 'PRINCIPAL'

        ## Update 1.3: To check the command Get-CenteralSqlDbInfo failed.
        if(!$?) {
            $PlanFlag = 2
            Write-Host "Warning: We cannot get principal DB from command Get-CenteralSqlDbInfo!" -ForegroundColor Yellow
            break
        }

        $servername = $sqlvm
        $PrincipalDB = $psqldistribution.count
        if ($PrincipalDB -gt $DefaultAlert) { 
            $color = "Red" 
            Write-Host "$servername got Alerting!" -ForegroundColor Red
        }
        
        elseif ($PrincipalDB -gt $DefaultWarning -and $PrincipalDB -lt $DefaultAlert) { 
                $color = "Yellow" 
                Write-Host "$servername got warning!" -ForegroundColor Yellow
        }

        else {$color = "White"}
        $HtmlBody += "<TR><TD colspan='3'>$servername</TD><TD style=background-color:$color colspan='3'>$PrincipalDB</TD></TR>"
    }

    ## Update 1.3: The second plan is using Get-CenteralDB
    if($PlanFlag -eq 2) {
    
        Write-Host "Trying the second plan..." -ForegroundColor Green

        $PrincipalDB = Get-CenteralDB -Farm $PRContentFarm | ? CurrentLoad -NE 0

        if($PrincipalDB) {

            $GroupCount = $PrincipalDB | group SQLHostName -NoElement
        
            foreach ($group in $GroupCount) {

                if ($group.Count -gt $DefaultAlert) { 
                    $color = $CommonColor['Red'] 
                    Write-Host "$servername got Alerting!" -ForegroundColor Red
                }
                elseif ($group.Count-gt $DefaultWarning -and $group.Count -lt $DefaultAlert) { 
                        $color = $CommonColor['Yellow'] 
                        Write-Host "$servername got warning!" -ForegroundColor Yellow
                }

                else { $color = $CommonColor['White'] }

                $HtmlBody += "<TR><TD colspan='3'>$($group.Name)</TD><TD style=background-color:$color colspan='3'>$($group.Count)</TD></TR>"
            }
        }
        else {
            # Show alert if no principal DB
            $alert = "We cannot find any principal DB in farm $($PRContentFarm.SqlFarmId)!"

            $HtmlBody += "<TR><TD Style=color:$($CommonColor['Red']) colspan='6'>" + $alert + "</TD></TR>"
            Write-Host $alert -ForegroundColor Red
        }
    }

        ## Update 1.2: Add DB mirror check
    # Usage DB is not important. Ignore if it don't have mirror.
    $DBs = Get-CenteralDB -Farm $($_.FarmId) -State Running | ? Role -NE Usage
    $NonMirrorDBs = $DBs | ? MirrorId -EQ 0

    if ($NonMirrorDBs) {
        $DBNames = Convert-ArrayAsString $NonMirrorDBs.Name
        Write-Host "$DBNames don't have Mirror!" -ForegroundColor Red
        $HtmlBody += ("<TR><TD Style=color:{1} colspan='6'>Database {0} don't have Mirror DB!</TD></TR>" -f $DBNames, $CommonColor['Red'])
    }
}



$HtmlBody += "</table>"

Write-Host "Check 'SQLDISTRIBUTION' done!" -ForegroundColor Green
Write-Host $separator

return $HtmlBody
