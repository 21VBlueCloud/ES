
#	 
#此脚本可以通过查询 scedueled task 备份出的XML信息，来实现 一键式部署计划任务。
#在以往过程中，由于服务器各种各样的问题，会导致计划任务失效，本脚本可以一键部署计划任务，省时省力。

if ($MyInvocation.MyCommand.Path -ne $null)
{
    $Script:basePath = Split-Path $MyInvocation.MyCommand.Path
    $Script:scriptname = Split-Path $MyInvocation.MyCommand.Path -Leaf
}
else
{
    $Script:basePath = "."
}

. D:\Operation\Tools\Library\LogHelper.ps1

$Xmls = Get-ChildItem -Filter *.xml -Path D:\Operation\tools\ -Recurse

CD   D:\Operation\tools\
foreach($xml in $xmls)
{
    [string]$taskname = $($xml.Name.split('.')[0])
    [string]$xmlFullName = $($xml.FullName)
    try
    {
        Write-POELog "Checking running task $taskname"
        schtasks /query  /tn "$($taskname)"
        
    }
    catch
    {
        Write-POELog -Message $_.Exception.Message -Level Error
        $errormsg = $_.Exception.Message
    }
    
    if($schedule -eq $null)
    {
        schtasks.exe /Create /TN $taskname /XML $xmlFullName /RU "TESTdomian\testUser" /RP ("testPWD."|ConvertTo-SecureString -AsPlainText -Force)
    }
    else
    {
        Write-POELog "$taskname is presented on the server"
    }
}






