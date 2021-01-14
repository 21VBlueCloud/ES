<#
#####################################################################################

##此脚本可以通过使用中国区由世纪互联运营的Office 365 邮箱托管服务， 来监控邮件的投递情况。
##  此脚本通过实现自动发送邮件以及查询邮件的投递情况，可以及时的发现邮箱投递过程中遇到的问题，并抓取的信息邮件发送给管理员。
##  
##  
##  
##  
##                              

#####################################################################################
#>
 begin
 {
    
    function Test-MailDeliver
    {
        <#
        .SYNOPSIS
        This function is used to check whether send-mailmessage cmdlet works 
        It will try to send a mail with same subject 3 times every 2 mins if send-mailmessage failed at first place. If send-mailmessage
        failed after 3 times retry, send mail to ops for further investigation with catched error message.
        #>
    param(
            [string]$subject,
            [string]$sender,
            [string]$recipient
         )

    $StopEmailLoop=$false
    [int]$RetryCount = 1
    #Update send mailmessage to make it try 3 times, if faile then send mail to team for further check
    Do
    {
        Try
        {          
            $startDate = Get-Date
            Write-POELog "Start to send test mail at $startDate."
            Send-mailmessage -Body Test -smtpServer smtp.partner.outlook.cn -to $recipient -from $sender  -subject $subject -credential $credential -UseSSL  -ErrorAction Stop
            $StopEmailLoop =$true
            $IsMailout  = $true
        }
        Catch
        {
            If ($RetryCount -gt 3)
            {					
                $errors = $_.Exception.Message
                $body = "Mail Subject $Subject fail to send out by cmdlet send-mailmessage, below is the error mesage"
                $body += $errors
                Send-Mail -Body $body -subject "OXE Mail Deliver Function Failed on send-mailmessage step, Please check!!!" -option 8
			    $StopEmailLoop = $true 
                $IsMailout = $false                  
		    }
		    Else 
            {			    
                Write-Host "Cannot send email $subject Trying again in 2 mins later "
                Write-POELog $_.Exception.Message -Level Error
			    Start-Sleep -Seconds 120
			    $RetryCount ++
		    }
        }                
    }
    While($StopEmailLoop -eq $false)

    return $IsMailout
}


    function Is-MTStatusDelivered
    {
    
    
    param(
        [string]$sender,
        [string]$subject,
        [string]$recipient,
        [string]$Organization 
    )


    $isMTDelivered = $false
    $retryCount = 1
    $status = $null
    
    do
    {
        $SendTrace = @(Get-MessageTrace –SenderAddress $sender  -Organization $Organization  –RecipientAddress $recipient |
                        ?{$_.Subject -eq $subject} | sort Received | select -Last 1)
        $status = $SendTrace.Status
        if($status -ne 'Delivered')
        {
            Write-POELog "Current sender's mail message is $status, we'll check it 1 mins later"
            Start-Sleep 60
        }
        else
        {
            $isMTDelivered  = $true
        }
        $retryCount ++
    }
    while($status -ne 'Delivered' -and $retryCount -ne 6)


    return $isMTDelivered
    }

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
    . D:\Operation\Tools\Library\CommonHelper.ps1
     
    $MaximumFunctionCount  =8192
    Load-ManagementShell

     
    $Logdate = Get-Date -Format "yyyy-MM-dd"
    Write-POELog -Message "Generating report for $($Logdate)"
    

    #Create Report folder to store CSV files
    $Folder = $Script:basePath + "\Reports"
    if(!(test-path $folder))
    {
        New-Item -Path $Folder -ItemType directory
    }
    $sender = "MailCheck@Stest.partner.onmschina.cn"
    $global:credential = New-Object Management.Automation.PSCredential $sender,('Abcd--1234' | ConvertTo-SecureString -AsPlainText -Force)
    #$sender = $credential.UserName
    #Better use another tenant’s admin account as  the recipient so we can get the receive trace
    $recipient= “MorningCheck@OXEtest.partner.onmschina.cn"
    $i = 1


}
process
{

    while((!($startDate.Hour -eq 23 -and $startDate.Minute -gt 55)))
    {
        ### Send mail failed for this time frame. We'll restart a new run test
        $subject = "[POE]Test Mail Delivery Function $LogDate <$i>"
        $IsMailOut = Test-MailDeliver -subject $subject -sender $sender -recipient $recipient
        if($IsMailOut -eq $false)
        {
            Write-POELog "Mail deliver failed at send-mailmessage step. No need to check message trace. "
            Write-POELog  "Notification mail sent out to ops, we'll wait 5 mins to start a new round test to make sure if it's a temp issue"
            Start-Sleep 300
        }
        else
        {
            Write-POELog "Testing mail $subject send successfully by cmdlet, check message trace for double confirm"
            $IsDelivered = Is-MTStatusDelivered -sender $sender -subject $subject -recipient $recipient -Organization  $sender.split(“@”)[1]

            if($IsDelivered -ne $True)
            {
                
                Wirte-Log "Check recipient message trace to check if the mail was received"
                $IsReceived = Is-MTStatusDelivered -sender $sender -subject $subject -recipient $recipient -Organization  $recipient.split(“@”)[1]
                
                if($IsReceived -ne $true)
                {
                    Write-Host "Get the latest trace"
                    $SendTrace = @(Get-MessageTrace –SenderAddress $sender  -Organization $sender.split(“@”)[1]  –RecipientAddress $recipient |
                        ?{$_.Subject -eq $subject} | sort Received | select -Last 1)
                    $RecipientTrace = @(Get-MessageTrace –SenderAddress $sender  -Organization $recipient.split(“@”)[1]  –RecipientAddress $recipient |
                        ?{$_.Subject -eq $subject} | sort Received | select -Last 1)


                    Write-POELog "Get the message trace detail for Both sender & Recipient"

                    $SendDetail  = $SendTrace | Get-MessageTraceDetail -Organization  $sender.split(“@”)[1] 
                    $RecipientDetail =  $RecipientTrace | Get-MessageTraceDetail -Organization  $recipient.split(“@”)[1]


                    #Below part is the message trace & trace detail & tracking log information for the ops


                    $Status = $SendTrace.Status
                    $MessageID = $($SendTrace.MessageId)
                    $MessageID = $MessageID -replace "<","&lt;"
                    $MessageID = $MessageID -replace ">","&gt;"
                    $MessageTraceId = $SendTrace.MessageTraceId.GUID

                    Write-POELog "Get the MB servers for the log filter"
                
                    #$Mtlog = Get-MessageTrackingLog -Server $MBServers -MessageId $MessageID


                    $trace = $SendTrace | Select Fromip,ToIP,received 
                    $sDetail = $SendDetail| Select Date,Event,Detail
                    $rDetail = $RecipientDetail | Select Date,Event,Action,Detail
                    #$TracedString = $traceProperties|Out-String
                    #Write-POELog $TraceString  -Level Warn
                    $MBServers = @([regex]::matches($($RecipientDetail.Detail),'\w{3}PR01MB\d{3,4}').value)

                    $body = "<Table><tr><td style=background-color:$($CommonColor['Red'])>Unhealthy Status : $Status</td></tr>" 
                    $body += "<tr><td>Sender Address : $sender </td></tr>" 
                    $body += "<tr><td>Recipient Address: $recipient </td></tr>" 
                    $body += "<tr><td>Subject : $subject </td><tr>" 
                    $body += "<tr><td>MessageID : $MessageID </td></tr>" 
                    $body += "<tr> <td>MessageTraceID: $MessageTraceId</td></tr>"
               

                    if($MBServers -ne $null)
                    {
                        $body += "<tr><td style=background-color:$($CommonColor['yellow']) font-weight:bold fontsize:30>Run below cmdlet in DMS capacity session 
                        `to get message tracking log detail</td></tr>"
                        $body += "<tr><td>Get-MessageTrackingLog -MessageId $MessageID -Server $($MBServers[0])}</td></tr>"
                    }


                    $body += "</Table>"
                    #add trace & trace detail in the mail 
                    $body +=  Format-HtmlTable $trace -Title "Message Trace"
                    $body +=  Format-HtmlTable $sDetail -Title "Send Message Trace Detail"
                    $body +=  Format-HtmlTable $rDetail -Title "Recipient Message Trace Detail"
                    Send-Mail -option 8 -Body $body -Subject "Urgent : OXE Mail Delivery Function Failed, please check !!!" 
            }
                else
                {
                    Write-POELog "Mail $subject was successfully delivered to recipient $recipient"
                }
            }
            else
            {
                Write-POELog "Mail deliver for $subject is success, we'll wait 5 mins to start a new round test"
                Start-Sleep 300
            }
      
        }    
  
       $i ++            
    }
}
end
{
    $summary = "$Logdate $scriptname executed $i times,and sent out $i mails"
    Write-POELog $summary
    Send-mail -Body $summary -Subject "[POE]Check_MailDelievery_Tool_Running_Status-$Logdate"
}
