
# 该脚本可以提供账户过期提示功能。
#将该脚本置入计划任务，脚本会定期运行并检测指定的用户的账户信息情况，并及时发送提醒给收件人。

#Get all IM user list from local file.
$IMUsers = get-content D:\Operation\Tools\Notification\POE_Send-notification_UserList.txt

ForEach ($User in $IMUsers) {

$IMUsersSplited = $User.Split(",")
$UserName = $IMUsersSplited[0]

#get users' properties from AD server

#Get the last changing time of password
$pwdLastSet = Get-ADUser $UserName -Properties * | Select-Object -ExpandProperty PasswordLastSet

#The cycle of the expired password
$pwdlastday=($pwdLastSet).AddDays(70)

#Get the current date
$now = Get-Date

#Identify wheter this is a never expires user
$neverexpire= Get-ADUser $UserName -Properties * | Select-Object -ExpandProperty PasswordNeverExpires

#Indentify how long the password will expire next time? 
$expire_days=($pwdlastday-$now).Days

#if match the expire condition, then send an email to the user

if($expire_days -lt 7 -and $neverexpire -like "false") {

$MailSender = "OXE-POE-TeamBox@TEST.com"

$MailRecevier = $IMUsersSplited[1]

$POETeam = "O365_POE@TEST.com"

$MailSubject = "POE account password expiration notification!"

$Name = $UserName.Split('_')[0]

$MailBody = "Hi $Name,

Your POE debug account password is going to expire in 7 days or already expired. Please follow the instructions below and change your password as soon as possible.

How to reset my POE account password?

>	Login to POE Lockbox.
>	Click 'CTRL + ALT + END'.
>	Click 'Change a Password' then Proceed with standard Windows change password steps.


If you can't change the password by yourself, please contact O365_POE@TEST.com."

#Send email message

send-mailmessage -smtpServer TEST.CHINACLOUDAPI.CN -from $MailSender -to $MailRecevier -Cc $POETeam -subject $MailSubject -body $MailBody

}
}

#Send POE team a notification email for checking POE debug account of IM folks
$today = Get-Date
send-mailmessage -smtpServer "TEST.CHINACLOUDAPI.CN" -from "OXE-POE-TeamBox@TEST.com" -to "O365_POE@TEST.com" -subject "The daily work of checking IM debug accounts have completed on $today" -body "Success! ^_^"
