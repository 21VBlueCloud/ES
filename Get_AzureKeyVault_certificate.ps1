## 该脚本可以轻易的获取到 Azure tenant里的Azure Key Vault数量。便于在一个tenant里有多个或者数量巨大的订阅中，获取Azure Key Vault信息的实例。


#Login-AzureRmAccount -EnvironmentName AzureChinaCloud
$allsub = Get-AzureRmSubscription #|select -First 2

$allsub.id 
$kvresults = @()
#$result1 =@()
foreach($sub in $allsub){	
    Set-AzureRmContext -SubscriptionId $sub.id
    $keyvaults = (Get-AzureRmKeyVault).vaultname 
    $keyvaults
    
    foreach($kv in $keyvaults)
    {
        #Set-AzureRmKeyVaultAccessPolicy -VaultName $kv -ObjectId  5f65a83a-6bed-4e76-a956-fb6762ffe498 -PermissionsToCertificates get,list 
        $count = @(Get-AzureKeyVaultCertificate -VaultName $kv).count
        $count
        $kvresult = New-Object PSObject
        $kvresult | Add-Member -MemberType NoteProperty -Name 'kVname' -Value $kv
        $kvresult | Add-Member -MemberType NoteProperty -Name 'Count' -Value $count
        $kvresult | Add-Member -MemberType NoteProperty -Name 'SubName' -Value $sub.Name
        $kvresult | Add-Member -MemberType NoteProperty -Name 'SubID' -Value $sub.ID
        $kvresults += $kvresult

    }


}
