[string]$MailBody
$EOPRACK=Get-CentralAdminRack

$MailBody="<TABLE border='1' cellpadding='0'cellspacing='0' style='Width:800px'>"
$MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='center'>XXX</TD></TR>"



Foreach($RACK in $EOPRACK)
{
If($RACK.Name -match "BJB")
  {
   $PMS=Get-CentralAdminMachine|?{$_.rack -match $RACK.Name}

   $MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='center'>BJB_Rack:$($RACK.Name)</TD></TR>"
   $i=$PMS.count-1
   $n=0
   While($n -le $i)
      {
       $MailBody+="<TR style=font-weight:bold><TD>$($PMS[$n].Name)</TD>"
       $n++
       if($n -le $i)
        {
        $MailBody+="<TD>$($PMS[$n].Name)</TD></TR>"
        }
        Else
        {
        $MailBody+="<TD></TD></TR>"
        }
        $n++

      }

   }
Else
   {
      $PMS=Get-CentralAdminMachine *SH*|?{$_.rack -match $RACK.Name}

   $MailBody+="<TR style=font-weight:bold;font-size:17px><TD colspan='2' align='center'>SHA_Rack:$($RACK.Name)</TD></TR>"
   $i=$PMS.count-1
   $n=0
   While($n -le $i)
      {
       $MailBody+="<TR style=font-weight:bold><TD>$($PMS[$n].Name)</TD>"
       $n++
       
       if($n -le $i)
        {
        $MailBody+="<TD>$($PMS[$n].Name)</TD></TR>"
        }
        Else
        {
        $MailBody+="<TD></TD></TR>"
        }
       $n++

      }
   
   }

}

$MailBody+="</Table>"

    $credentials = new-object Management.Automation.PSCredential <Account>, (<psw> | ConvertTo-SecureString -AsPlainText -Force)
    Send-mailmessage -Body $MailBody -BodyAsHtml -smtpServer <SMTP> -to "Recipient" -from "Sender" -subject "XXX" -credential $credentials -UseSsl
