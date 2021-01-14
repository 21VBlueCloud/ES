<#
#####################################################################################
#
#File: Tenants.ps1
#Author: Wende SONG (Wind)
#Version: 1.0
#
##  Revision History:
##  Date       Version    Alias       Reason for change
##  --------   -------   --------    ---------------------------------------
##
##  9/19/2016   1.0       Wind        V6 edition.
##
##  6/14/2017   1.2       Wind        Send tenantcounter.csv to Emma if added record.
##
##  11/11/2018  1.3       Wind        Delete parameter "PrimaryFarm" from get-SPDFQDN.
##
##  5/21/2019   1.4       Bruce       Add Real Customers Pending Tenants count where 
##                                    SpsiteSubscription is not Null 
##
##  6/12/2019   1.5       Wind        Change query method that query tenants from SPD
##                                    directly.
##
##  11/6/2019   1.6       Wind        Add foreach looping to query Tenant counts for 
##                                    multiple legacy netowrk.
#####################################################################################
#>

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
$HtmlBody = "<TABLE Class='TENANTS' border='1' cellpadding='0'cellspacing='0' style='Width:900px'>"
##===============================================================================================================
    
    Write-Host $separator

    [Int] $LowerCorrupted = $Xml.DailyCheck.Tenants.Corrupted.LowerThreshold
    [Int] $UpperCorrupted = $Xml.DailyCheck.Tenants.Corrupted.UpperThreshold
    [Int] $LowerPending = $Xml.DailyCheck.Tenants.Pending.LowerThreshold
    [Int] $UpperPending = $Xml.DailyCheck.Tenants.Pending.UpperThreshold
    
    # 1.4. Referring to Customers Pending Tenant lower and upper threshold
    [Int] $LowerCustomersPending = $Xml.DailyCheck.Tenants.CustomersPending.LowerThreshold
    [Int] $UpperCustomersPending = $Xml.DailyCheck.Tenants.CustomersPending.UpperThreshold

    $CurrentDate = get-date
    $FirstDateOfThisMonth = Get-Date -Year $CurrentDate.Year -Month ($CurrentDate.Month) -Date 1
    $FirstDateStr = $FirstDateOfThisMonth.ToUniversalTime().ToString("yyyyMMddHHmmss.0Z")
    $Today = $CurrentDate.ToUniversalTime().ToString("yyyyMMddHHmmss.0Z")

    $SPDFQDN = Get-SPDFQDN -First 1
    $searchbase = "OU=Tenants,OU=MSOnline,DC=SPODS00300022,DC=MsoCHN,DC=msft,DC=net"

    ## update 1.6: Add foreach looping to query Tenant counts for multiple legacy netowrk.
    $SPONetwork = @(Get-CenteralNetwork | ? name -Match "LegacyNetwork" | ? RecoveryNetworkId -NE 0)

    # To check network have SPD VMs to avoid tenant count query failed.
    $NetworkId = @()
    foreach ($network in $SPONetwork) {
        $HaveSPD = Get-CenteralVM -Network $network -Role SPD -State Running | Measure-Object | Select-Object -ExpandProperty Count
        if ($HaveSPD) { $NetworkId += $network.NetworkId }
    }
    # $NetworkId = $SPONetwork.NetworkId


    Import-Module D:\DebugHotSync\OCE\POP\Onboarding\WorkItemHelper.ps1 | Out-Null

    <#
    # $PendingTenantCount=(Get-SXDSPendingTenantsCount -NetworkId $NetworkId).Count
    $PendingTenants = Get-SXDSPendingTenants -NetworkId $NetworkId
    $PendingTenantCount = $PendingTenants.Count

    # 1.4. Adding Customers Pending Tenants count
    
    $CustomersPendingTenantCount = 0
    foreach ($sin in $PendingTenants) {
        if ((Get-CenteralTenant -SpSiteSubscription $sin.CompanyId) -ne $null) {
            $CustomersPendingTenantCount ++
        }
    }
    $CorruptTenantCount=(Get-SXDSCorruptedTenants -NetworkId $NetworkId -LastUpdatedSince $FirstDateOfThisMonth | measure).Count
    #>

    ## Update 1.5: Query corrupted tenants from SPD directly
    Write-Host "Getting corrupted tenants ..." -ForegroundColor Green
    
    $CorruptedTenantSearcher = "MsOnline-ObjectID -like '*' -and " + `
                                "objectClass -eq 'OrganizationalUnit' -and " + `
                                "whenChanged -ge '$FirstDateStr' -and " + `
                                "whenChanged -le '$Today' -and " + `
                                "SPO-IsDataCorrupted -eq 'TRUE'"

    $CorruptedTenants = @(Get-ADOrganizationalUnit -SearchBase $searchbase -SearchScope Subtree `
                          -Server $SPDFQDN -Filter $CorruptedTenantSearcher -Properties SPO-IsSynthetic, DisplayName)

    # filter out synthetic tenants
    $CorruptedTenants = @($CorruptedTenants | ? SPO-IsSynthetic -ne $true)

    # filter out Heart beat
    $CorruptedTenants = @($CorruptedTenants | ? DisplayName -NotMatch "SPOProvHeartbeat")

    $CorruptTenantCount = $CorruptedTenants.Count
    Write-Host "Corrupted tenant count: $CorruptTenantCount"
    Write-Host "Done" -ForegroundColor Green

    ## Update 1.5: Query pending tenants from SPD directly
    Write-Host "Getting pending tenants ..." -ForegroundColor Green

    $PendingTenantSearcher = "MsOnline-ObjectID -like '*' -and " + `
                                "objectClass -eq 'OrganizationalUnit' -and " + `
                                "whenChanged -ge '$FirstDateStr' -and " + `
                                "whenChanged -le '$Today' -and " + `
                                "(msOnline-ProvisioningStamp -like '*' -or msOnline-PublishingStamp -like '*') -and " + `
                                "msOnline-PendingDeletion -notlike '*'"

    $PendingTenants = @(Get-ADOrganizationalUnit -SearchBase $searchbase -SearchScope Subtree `
                          -Server $SPDFQDN -Filter $PendingTenantSearcher -Properties SPO-IsSynthetic, DisplayName)

    # filter out synthetic tenants
    $PendingTenants = @($PendingTenants | ? SPO-IsSynthetic -ne $true)

    # Filter out displayname contains 'test'.
    $PendingTenants = @($PendingTenants | ? Displayname -NotMatch 'test')
    $CustomersPendingTenantCount = $PendingTenants.Count
    Write-Host "Pending tenant count:$CustomersPendingTenantCount"
    Write-Host "Done" -ForegroundColor Green



    $HtmlBody += "<TR style=background-color:#0066CC;font-weight:bold;font-size:17px>`
                    <TD colspan='6' align='center' style=color:#FAF4FF>Tenant Count</TD></TR>"

    if($CorruptTenantCount -lt $LowerCorrupted) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>CorruptedTenant in Network $NetworkId (from $($FirstDateOfThisMonth.ToString("MM/dd/yyyy"))) :  $CorruptTenantCount </TD></TR>"
    }
    ElseIf ($CorruptTenantCount -ge $LowerCorrupted -and $CorruptTenantCount -lt $UpperCorrupted) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Yellow']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>CorruptedTenant in Network $NetworkId (from $($FirstDateOfThisMonth.ToString("MM/dd/yyyy"))) :  $CorruptTenantCount </TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);font-weight:bold;font-size:17px><TD colspan='6' align='left' style=color:$($CommonColor['White'])>CorruptedTenant in Network $NetworkId (from $FirstDateOfThisMonth) :  $CorruptTenantCount </TD></TR>"
    }

    <#
    if($PendingTenantCount -lt $LowerPending) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>PendingTenant in Network $NetworkId :  $PendingTenantCount </TD></TR>"
    }
    elseif ($PendingTenantCount -ge $LowerPending -and $PendingTenantCount -lt $UpperPending) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Yellow']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>PendingTenant in Network $NetworkId :  $PendingTenantCount </TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);font-weight:bold;font-size:17px><TD colspan='6' align='left' style=color:$($CommonColor['White'])>PendingTenant in Network $NetworkId :  $PendingTenantCount </TD></TR>"
    }
    #>

    # 1.4. Adding the count result into HtmlBody
    if($CustomersPendingTenantCount -lt $LowerCustomersPending) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>CustomersPendingTenant in Network $NetworkId (from $($FirstDateOfThisMonth.ToString("MM/dd/yyyy"))):  $CustomersPendingTenantCount </TD></TR>"
    }
    Elseif ($CustomersPendingTenantCount -ge $LowerCustomersPending -and $CustomersPendingTenantCount -lt $UpperCustomersPending) {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Yellow']);font-weight:bold;font-size:17px><TD colspan='6' align='left'>CustomersPendingTenant in Network $NetworkId (from $($FirstDateOfThisMonth.ToString("MM/dd/yyyy"))):  $CustomersPendingTenantCount </TD></TR>"
    }
    else {
        $HtmlBody += "<TR style=background-color:$($CommonColor['Red']);font-weight:bold;font-size:17px><TD colspan='6' align='left' style=color:$($CommonColor['White'])>CustomersPendingTenant in Network $NetworkId  (from $($FirstDateOfThisMonth.ToString("MM/dd/yyyy"))):  $CustomersPendingTenantCount </TD></TR>"
    }

    #===================================================================================================================================
    Write-Host "Checking Sign up & GA tenants count ..." -ForegroundColor Green

    $GADate = (Get-Date -Date "04-15-2014").ToUniversalTime().ToString("yyyyMMddHHmmss.0Z")
    $SignUpDate = Get-Date -Date "8/7/2013 7:00 PM"
    
    $searchFilter = "MsOnline-ObjectID -like '*' -and " + `
                                "objectClass -eq 'OrganizationalUnit' -and " + `
                                "whenCreated -ge '$GADate' -and " + `
                                "whenCreated -le '$Today' -and " + `
                                "(msOnline-ProvisioningStamp -notlike '*' -or msOnline-PublishingStamp -notlike '*') -and " + `
                                "msOnline-PendingDeletion -notlike '*'"

    $GAResults = @(Get-ADOrganizationalUnit -SearchBase $searchbase -SearchScope OneLevel `
                          -Server $SPDFQDN -Filter $searchFilter)

    If ($GAResults -ne $null) { 
        $GACount = $GAResults.Count 
    }

    ## Udpate 1.6: Add foreach looping to query Tenant counts for multiple legacy netowrk.
    $SignUpCount=0;
    foreach($i in $NetworkId)
    {
        $SignUpResults = Get-SXDSUpdatedTenantsCount -NetWorkId $i -LastUpdatedSince $SignUpDate
        $SignUpCount +=$SignUpResults.Count
    }


    # Add tenants counter.
    $Properites = @{DateTime=$CurrentDate;GATenants=$GACount;Tenants=$SignUpCount}
    $Record = New-Object -TypeName PSObject -Property $Properites
    $CounterPath = $Xml.DailyCheck.Tenants.CounterPath
    # If 'TenantCounter.csv' is not exsiting, create it.
    If (!(test-path $CounterPath)) {
        Write-Host "Creating Tenant counts." -ForegroundColor Cyan
        New-Item -Path $CounterPath -ItemType File
        $Record | Export-Csv -Path $CounterPath
    }

    # Record once every day.
    $HistoryRecord = @(Import-Csv -Path $CounterPath)
    $AddRecord = $true
    $ShortDate = $CurrentDate.ToShortDateString()
    foreach ($hr in $HistoryRecord) {
        $HShortDate = (get-date -Date $hr.DateTime).ToShortDateString()
        If ($HShortDate -eq $ShortDate) { 
            $AddRecord = $false 
            Write-Host "Will not record tenant counts due to it done today!" -ForegroundColor Yellow
            Break
        }
    }

    If ($AddRecord) {
        Write-Host "Export tenant counts!" 
        $Record | Export-Csv -Path $CounterPath -Append
        # update: 1.2
        if ($?) {
            $To = $Xml.DailyCheck.Tenants.SendAttachments.To
            $Attachments = $Xml.DailyCheck.Tenants.SendAttachments.Attachments
            try {
                Send-Email -To $To -mailbody "PFA" -mailsubject "Tenant Counter on $($CurrentDate.ToString("yyyyMMddhhmmss"))" `
                            -Attachments $Attachments
            }
            catch {
                Write-Host "Cannot send mail!" -ForegroundColor Red
                Write-Host $_
            }
        }
    }
    

    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px><td colspan='6' align='left'>$GACount GA Tenants Provisioned since 4/15/2014</td></TR>"
    $HtmlBody += "<TR style=background-color:$($CommonColor['LightBlue']);font-weight:bold;font-size:16px><td colspan='6' align='left'>$SignUpCount Tenants Provisioned since 8/7/2013</td></TR>"
    Write-Host "Done" -ForegroundColor Green

    


    Write-Host "Checking for 'Tenants' done." -ForegroundColor Green
    Write-Host $separator

# Post process
##===============================================================================================================
$HtmlBody += "</table>"

return $HtmlBody
#$HtmlBody | Out-File .\test.html
#Start .\test.html
##===============================================================================================================