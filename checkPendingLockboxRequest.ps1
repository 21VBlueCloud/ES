 ## 该脚本可以监控系统中，队列中的请求信息， 并通过指定的邮箱发送提醒。
 ##  获取的请求信息，会通过一个function放入邮件内容中的table里，展示效果清晰易懂。
 



 Function Send-Mail 
{
    Param (
           
            $SmtpServer="mail.21vianet.com",
            [String] $from = "test@test.com",
            [String[]] $to = "test@test.com;",
           
            [Parameter(Mandatory=$true)]
            [String] $Subject,
            [Parameter(Mandatory=$true)]
            $Body,
            $Attachments,
            $BCC,
            $CC,
            $option
    	    )

    If ($Body.GetType().IsArray) { $MailBody = -join $Body }
    Else { $MailBody = $Body }
   
     try
    { 
        switch($option)
        {
            0{Write-host "Send mail with CC List";Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -cc $cc -ErrorAction Stop}
            1{Write-host "Send mail with CC List & Attachments"; Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -Attachments $Attachments -cc $cc -ErrorAction Stop}
            2{Write-host "Send mail with CC List & BCC List & Attachments"; Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -Attachments $Attachments -cc $cc -Bcc $bcc -ErrorAction Stop}
            3{Write-host "Send mail with CC List & BCC List "; Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -cc $cc -bcc $bcc -ErrorAction Stop}
            4{Write-host "Send mail with BCC List & Attachments";Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -Attachments $Attachments -bcc $bcc -ErrorAction Stop}
            5{Write-host "Send mail with  BCC List";Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -bcc $bcc -ErrorAction Stop}
            6{Write-host "Send mail with Attachments";Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -Attachments $Attachments -ErrorAction Stop}
            default{Send-mailmessage -Body $MailBody -smtpServer $smtpserver -to $to -from $from -Subject $Subject -BodyAsHtml -ErrorAction Stop}                
        }
    }
    catch
    {
         Write-host $_.exception.message -level Error
         Continue
    }
   
}

 Function Format-HtmlTable 
{
    
    Param (
    $Contents,
    [String]$Title = $null,
    [Int] $TableBorder=1,
    [Int] $Cellpadding=0,
    [Int] $Cellspacing=0,
    [INT]$width = 1000,
    [hashtable]$keywords_color
    )
    $TableHeader = "<table style='Width:$width' +'px' border=$TableBorder cellpadding=$Cellpadding cellspacing=$Cellspacing" 

    # Add title
    if($Title -ne $null -and $Title -ne '' -and $Title.Count -ne 0)
    {
        $HTMLContent = $Contents | ConvertTo-Html -Fragment -PreContent "<H2>$Title</H2>"
    }
    else
    {
         $HTMLContent = $Contents | ConvertTo-Html -Fragment
    }
    #Add table border
    $HTMLContent = $HTMLContent -replace "<table", $TableHeader
    # Strong table header
    $HTMLContent = $HTMLContent -replace "<th>", "<th><strong>"
    $HTMLContent = $HTMLContent -replace "</th>", "</strong></th>"

    # Color Items
    foreach($item in $keywords_color.Keys)
    {
        for ($i=0;$i -lt $HTMLContent.count;$i++) 
        {
            If ($HTMLContent[$i].IndexOf("$item") -ge 0) 
            {
                $InsertString = " bgcolor= " + $keywords_color[$item]
                $FindStr = "<tr"
                $Position = $HTMLContent[$i].LastIndexOf($FindStr) + $FindStr.Length
                $HTMLContent[$i] = $HTMLContent[$i].Insert($Position,$InsertString)
            }
            
        }
    
    }

    Return $HTMLContent
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

function CheckPendingLockboxRequest
{
    $date = Get-Date

    $allLockboxRequest = Get-LockboxRequestForAudit -StartTimeUtc ((get-date).AddHours(-5)) -ErrorAction SilentlyContinue
    Write-Host "Get all lockbox requests"
    $PendingLockboxRequest = $allLockboxRequest|?{$_.ApprovalStatus.Value -eq 'Pending'}|select Id,CreateTime,Requestor,ApprovalStatus,Role,Reason,OperationName,Workload

    Write-Host "Get pending requests"
    $ExpiredLockboxRequests = $allLockboxRequest|?{$_.ApprovalStatus.Value -eq 'Expired' }|select Id,CreateTime,Requestor,ApprovalStatus,Role,Reason,OperationName,Workload
    Write-Host "Get Expired requests"

    if($PendingLockboxRequest)
    {
       $emailbody = Format-HtmlTable -Contents $PendingLockboxRequest -Title "Pending Lockbox Request. Please take action (Createtime is UTC time)" -TableBorder 1 -keywords_color $keywords_color
       Send-mail -Body $emailbody -Subject "Gallatin_DMS-Pending_LockboxRequest-$date"  -to test@test.com
      #Send-mail -Body $emailbody -Subject "Gallatin_DMS-Pending_LockboxRequest-$date"  -to test@test.com
       Write-Host "Send out Pending request mail"   
    }
    if($ExpiredLockboxRequests)
    {
       $emailbody = Format-HtmlTable -Contents $ExpiredLockboxRequests -Title "Expired Lockbox Request, Please take action (Createtime is UTC time)" -TableBorder 1 -keywords_color $keywords_color
       Send-mail -Body $emailbody -Subject "Gallatin_DMS-Expired_LockboxRequest-$date"  -to test@test.com
      #Send-mail -Body $emailbody -Subject "Gallatin_DMS-Expired_LockboxRequest-$date"  -to test@test.com
       Write-Host "Send out Expired request mail"
    }

    }


 

#Endtime is 12 hours later compare with the time when the script starts.
$EndTime = (get-date).AddHours(5)

    While((Get-Date) -le $EndTime)
{ 
    #check pending lockbox reuqest in DMS.
   checkpendinglockboxrequest  
   write-host "Now the progress will sleep for 15 minutes."
   # this function provided by Wind, which will set the progress sleep for 15 mins, then continue.
   write-host "The next round of check will execute in 15 minutes later"
   sleep 900

}
