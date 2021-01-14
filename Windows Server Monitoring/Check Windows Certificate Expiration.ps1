
#   此脚本用于域环境中，监控证书的使用和是否即将过期。
#   在设定的时间运行，可以触发邮件并把所得结果邮件发送给指定收件人。 邮件过期信息也会根据时间显示颜色不同。
#.Description
#    This script will generate 2 reports to folder "'D:\Tools\Scripts\certs\", 1 is getting certs for all DC servers OXE or POE
#    2 is for all the certs which are going to expire in 90 days.
# 


#Get personal certs
Function Get-Cert( $computer=$env:computername)
{

    $ro=[System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"

    $lm=[System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"

    $store=new-object System.Security.Cryptography.X509Certificates.X509Store("\\$computer\my",$lm)

    $store.Open($ro)

    $store.Certificates

}

#Check script location & name
if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
    $index = $Script:scriptname.LastIndOXEf("_")
    $scriptname = $scriptname.substring(0,$index)
}
else
{
    $Script:basePath = "."
}

. D:\Operation\Tools\Library\LogHelper.ps1

#Load-ManagementShell

$StartTime = Get-Date  

$Logdate = Get-Date -Format "yyyy-MM-dd"
Write-POELog -Message "Generating report for $($Logdate)"

#Create report folder to save all report files
$Folder = $Script:basePath + "\Reports"
if(!(test-path $folder))
{
    New-Item -Path $Folder -ItemType directory
}

$Allcerts = @()
$Computers = @()
$IssueServers =@()
$Computers = Get-ADComputer -Filter *|Sort Name

$i = 0
# Get all certs on all servers
foreach($Computer in $Computers)
{ 
    Write-Progress -Activity "Retrievw Certs" -Status "($i/$($computers.count)) Complete,now processing $($computer.name):" -PercentComplete ($i/$($computers).count*100)
    try
    {
        $Certs = @()
        $Certs =Get-Cert $Computer.Name
        Write-POELog -Message "$($Computer.Name) Certs count $($certs.count)"
        $Certs|Add-Member -MemberType NoteProperty  -Name 'ServerName' -Value $Computer.Name
	}
    catch
    {
      #"Cannot remote to $($Computer.Name)"
       $IssueServers += @($Computer.Name)
    }
    
    $Allcerts += $Certs
    $i++
}

# Define keywords and color for the prepartion mail
$Keywords_Color = @{}
$keywords_color.add('WoSign','#BBFFFF')
$keywords_color.add('CNNIC','#FFE4B5')
$keywords_color.add('DigiCert','#BCEE68')


#Export all certs to a csv file

Write-POELog  "Total certs count is  $($Allcerts.count) certs"
$AllcertsProp =$Allcerts|select @{Name="Expiration_date";Expression={$_.notafter.tostring().split('')[0]}},`
                 @{Name="Issued_To";Expression={$sub = $_.subject.split("=")[1];$sub.split(',')[0]}},`
                 @{Name="Issued_By";Expression={$issuer =$_.issuer.split("=")[1];$issuer.split(',')[0]}},`
                 @{Name = "Server_Name";;Expression={$_.ServerName}},Thumbprint
                 
$AllcertsProp|export-csv $Folder\"AllCerts_$($Logdate)".csv 

#Check Active non-Wosign
$AllActiveCerts=@()
$AllActiveCerts = $Allcerts|?{$_.NotAfter -ge (get-date) -and $_.Issuer -notlike "*Wosign*"}
$AllActivecertsProp =$AllActiveCerts|select @{Name="Expiration_date";Expression={$_.notafter.tostring().split('')[0]}},`
                 @{Name="Issued_To";Expression={$sub = $_.subject.split("=")[1];$sub.split(',')[0]}},`
                 @{Name="Issued_By";Expression={$issuer =$_.issuer.split("=")[1];$issuer.split(',')[0]}},`
                 @{Name = "Server_Name";;Expression={$_.ServerName}},Thumbprint
$AllActivecertsProp|export-csv $Folder\"AllActiveCerts_$($Logdate)".csv 

# Check near expiration non-wosign certs, if not null, generate reports, and create near expiration certs file,  send mail for further investigation

$NearExpirationCerts = $Allcerts|?{$_.NotAfter -le (get-date).AddDays(90) -and $_.NotAfter -ge (get-date) -and $_.Issuer -notlike "*Wosign*"}

#Prepare  Expiration notification report
if($NearExpirationCerts -ne $null)
{
    $NeCertsProp = $NearExpirationCerts|select @{Name="Expiration_date";Expression={$_.notafter.tostring().split('')[0]}},`
                 @{Name="Issued_To";Expression={$sub = $_.subject.split("=")[1];$sub.split(',')[0]}},`
                 @{Name="Issued_By";Expression={$issuer =$_.issuer.split("=")[1];$issuer.split(',')[0]}},`
                 @{Name = "Server_Name";;Expression={$_.ServerName}},Thumbprint
               
    $NeCertsProp|export-csv  $Folder\"NearExpirationCerts_$($Logdate)".csv  

    #Put Certs are going to expire Report in the mail body

    $NecertsGroup = $NeCertsProp|group Expiration_Date,Issued_by,Issued_To -NoElement|sort Issued_by|`
                     select @{Name='Total Count';Expression={$_.Count}},`
                     @{Name='Expiration Date'; Expression={$_.Name.Split(',')[0]}},`
                     @{Name='Issued_By'; Expression={$_.Name.split(',')[1]}},`
                     @{Name='Issued_To'; Expression={$_.Name.split(',')[2]}}
    $title = "Below Certs are going to expire in later 90 days, please check" 
    $emailbody = Format-HtmlTable  -contents $NecertsGroup -Title $title -keywords_color $keywords_color                                              
   
  }
else
{
    #Prepare  all certs notification mailCount -ne 1}  
    $title = "No Certs are going to expire in later 90 days @^_^"
    $AllcertsGroup = $AllActivecertsProp|group Expiration_Date,Issued_by,Issued_To -NoElement|?{$_.Count -gt 2}|`
                     select @{Name='Total_Count';Expression={$_.Count}},`
                     @{Name='Expiration_Date'; Expression={$_.Name.Split(',')[0]}},`
                     @{Name='Issued_By'; Expression={$_.Name.split(',')[1]}},`
                     @{Name='Issued_To'; Expression={$_.Name.split(',')[2]}} |Sort Issued_by,Total_Count
    $title = "No Certs are going to expire in later 90 days, All certs status Summary is as below" 
    $emailbody = Format-HtmlTable  -contents $AllcertsGroup -Title $title -keywords_color $keywords_color  

}  

#Send report to team

if($NearExpirationCerts.count -eq 0)
{
    Write-POELog  -Message "No certs are going to expire,send mail to team for reference"    
    
    Send-mail -Body $emailbody -Subject "POE-Certs monitor report"  -Attachments $Folder\"allCerts_$($Logdate)".csv  -option 1    
              
   
}
else
{
    Write-POELog  -Message "certs in the mail  are going to expire ,send mail to team for further investigation"
    
    Send-Mail  -Body $emailbody -Subject "POE-Certs Monitoring report" -Attachments $Folder\"NearExpirationCerts_$($Logdate)".csv, $Folder\"allCerts_$($Logdate)".csv   -option 1 
              
}
   

$endtime = get-date
$timespan = $endtime -$StartTime
if($timespan.Hour -le 0)
{
    if($timespan.Minutes -gt 1)
    {
        Write-POELog  -Message "total execution time is $($timespan.minutes) minutes & $($timespan.seconds) seconds"
    }
    else
    {
        Write-POELog  -Message "total execution time is $($timespan.seconds) seconds"
    }
}
else
{
    Write-POELog -Level Error -Message "Report Generation is longer than 1 hour, please check"
}
Write-POELog -Message "$($IssueServers.Count) servers are not able to connect, please check the list $IssueServers "
Write-Host "total execution time is $($timespan.minutes) minutes & $($timespan.seconds) seconds"
#Read-Host -Prompt “Press Enter to exit”






