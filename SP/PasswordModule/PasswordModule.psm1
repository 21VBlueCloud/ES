function New-SecretCredential {
<#
.Synopsis
Create new secret credential and store into local.

.Description
This function will create a new local credential and store it at "$home\Mypassword\<ComputerName>".
The parameter 'Alias' is unique identity to distinguish credential in local.
The parameter 'UserName' is your credential's username, such as 'chn\oe-songwende_ada'.
The parameter 'Password' is the credential's password. Console will prompt to enter it, if you ommit
this parameter.

.Example
New-SecretCredential -Alias ada -UserName chn\oe-songwende_ada
This command will prompt you enter a secure string and create a local credential.

.Example
New-SecretCredential ado chn\oe-songwende_ado -Password $ado_cred
This command will create a local credential.

#>

    param(
        [Parameter(Mandatory=$true, Position=0)]
        [String]$Alias,
        [Parameter(Mandatory=$true, Position=1)]
        [String]$UserName,
        [System.Security.SecureString]$Password
    )

    if(!$Password) {
        $Password = Read-Host -Prompt "Please enter password" -AsSecureString
    }

    # check current Json file and make sure there is no same alias exists.
    $ExistObj = FindCredential -Alias $Alias
    if($ExistObj) {
        Write-Error "Same alias exists!"
        $LASTEXITCODE = 10
        break
    }

    # add new credential into Json.
    $Objects = @(GetJsonPSObject)
    $Objects += MakePSObject -Alias $Alias -UserName $UserName -Password $Password
    $Objects | ConvertTo-Json | Out-File $Global:SecretPath

    # Show new list
    Show-SecretCredential

}

function Set-SecretCredential {
<#
.Synopsis
Update encrypted password or with domain password both.

.Description
Update encrypted password in local or update domain password too.

The parameter 'Alias' is a identifer to select which local credential to update.
The parameter 'Password' is your new password to udpate. If you ommit it, console will prompt you to enter.
If parameter 'ADUpdate' is on, this command will use current local password as old password and new password you entered 
to update domain password and then update local.

.Example
Set-SecretCredential -Alias ada
Update local crednetial for ada account only after you enter your new password.

.Example
Set-SecretCredential -Alias ada -ADUpdate
Update ada account password in AD and local both after you enter your new password.
#>

    param(
        [Parameter(Mandatory=$true)]
        [String]$Alias,
        [System.Security.SecureString]$Password,
        [Switch]$ADUpdate
    )

    if(!$Password) {
        $Password = Read-Host -Prompt "Your new password" -AsSecureString
    }

    # find out current local credential.
    $Objects = GetJsonPSObject
    $object = $Objects | Where-Object Alias -eq $Alias

    if($object) {

        $UserName = $object.UserName -split '\\' | Select-Object -Last 1

        # Update domain password first to make sure local time stamp greater than 
        # last password set time in AD.
        if($ADUpdate.IsPresent) {

            # update domain password
            $OldPassword = $object.EncryptedPwd | ConvertTo-SecureString


            Set-ADAccountPassword -Identity $UserName -OldPassword $OldPassword `
                -NewPassword $Password -ErrorAction Stop

            if($?) {
                Write-Host "Domain password has been updated!" -ForegroundColor Green
            }
            else {
                break
            }

        }

        # We should make sure the password has been updated in AD before we change in local
        # to avoild wrong password in local.
        $ADUser = Get-ADUser $UserName -Properties PasswordLastSet
        if($ADUser.PasswordLastSet -lt $object.DateTime) {
            Write-Host "Warning: Your domain password has not been updated!" -ForegroundColor Yellow
            $proceed = Read-Host -Prompt "Are you sure to update local password without domain?(yes/no)"
            if($proceed -ne 'yes') {
                Write-Error "User aborted!"
                break
            }
        }
        
        # time stamp shift 1 minute to domain password set time.
        $DateTime = Get-Date | ForEach-Object -MemberName AddMinutes -ArgumentList 1
        $object.DateTime = $DateTime
        $object.EncryptedPwd = $Password | ConvertFrom-SecureString

        $Objects | ConvertTo-Json | Out-File -FilePath $Global:SecretPath 
        if($?) {
            Write-Host "Loca credential has been updated!" -ForegroundColor Green
        }
        else {
            Write-Host "Local credential update failed!" -ForegroundColor Red
        }

    }

}

function Test-SecretCredential {
<#
.Synopsis
Test the encrypted password is valid or not. It also can test domain account expiration.

.Description
It will show you that local password, AD password are expired or not and show how many days will AD
password expire.

Parameter 'Passthru' will return test result as object.

.Example
Test-SecretCredential -Alias ada
Just test ada account expiration for local and domain both.
#>

    param(
        [Parameter(Mandatory=$true)]
        [String]$Alias,
        [Switch]$Passthru
    )

    $object = FindCredential -Alias $Alias

    if($object) {
        # to verify the date is before expiration.
        $DateTime = Get-Date -Date $object.DateTime 
        $UserName = $object.UserName -split '\\' | Select-Object -Last 1
        $ADUser = Get-ADUser $UserName -Properties PasswordLastSet -ErrorAction Ignore
        $LocalPasswordExpired = $false
        $ADPasswordExpired = $false

        if($ADUser) {

            $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy
            $now = Get-Date
            $ExpireDate = $ADUser.PasswordLastSet + $PasswordPolicy.MaxPasswordAge
            if($ADUser.PasswordLastSet -gt $DateTime) {
                $LocalPasswordExpired = $true
            }

            if($ExpireDate -lt $now) {
                $ADPasswordExpired = $true
            }

            $ExpireDays = $ExpireDate - $now

        }
        else {
            Write-Error "AD user '$($object.UserName)' does not exist!"
            break
        }
    }
    else {
        Write-Error "No valid credential for alias '$Alias'."
        break
    }

    if($Passthru.IsPresent) {
        return New-Object -TypeName PSObject -Property @{
            LocalPasswordExpired = $LocalPasswordExpired
            ADPasswordExpired = $ADPasswordExpired
            ADExpireDays = $ExpireDays.Days
        }
    }
    else {
        Write-Host "Local password expired: $LocalPasswordExpired" 
        Write-Host "AD password expired: $ADPasswordExpired" 
        Write-Host "AD password expire days: $ExpireDays"
    }
}

function Remove-SecretCredential {
<#
.Synopsis
Remove lcoal credential only.

.Description
Just remove local credential. Domain account will not touch anything.
#>

    param(
        [Parameter(Mandatory=$true)]
        [String]$Alias
    )

    $Object = FindCredential -Alias $Alias

    if($Object) {

        $Object | Format-Table -AutoSize Alias, UserName, DateTime
        $confirm = Read-Host -Prompt "Are you sure delete it?(yes/no)"

        if($confirm -eq 'yes') {
            $Objects = GetJsonPSObject | Where-Object Alias -ne $Alias
            $Objects | ConvertTo-Json | Out-File -FilePath $Global:SecretPath -ErrorAction Stop
            Write-Host "Deleted!" -ForegroundColor Green
        }
        else {
            Write-Error "User aborted!"
        }

    }
    else {
        Write-Error "Cannot find local credential!"
    }

}

function Get-SecretCredential {
<#
.Synopsis
Get the local credential and return a PS credential.

.Description
Just retrive local credential only. But it will check expiration in local and AD both and show warning.

.Example
$ada_cred = Get-SecretCredential -Alias ada
Retrive ada credential.
#>

    param(
        [Parameter(Mandatory=$true)]
        [String]$Alias
    )

    $TestResult = Test-SecretCredential -Alias $Alias -Passthru

    if($TestResult) {

        if($TestResult.LocalPasswordExpired) {
            Write-Host "Your local password has been expired!" -ForegroundColor Red
            Write-Host "Please use 'Set-SecretCredential -Alias $Alias' to update." -ForegroundColor Yellow
            break
        }

        if($TestResult.ADPasswordExpired) {
            Write-Host "Your domain password has been expired!" -ForegroundColor Red
            Write-Host "Please use 'Set-SecretCredential -Alias $Alias -WithDomain' to update." -ForegroundColor Yellow
            break
        }

        if($TestResult.ADExpireDays -lt 7) {
            Write-Host "Warning: Domain password will expire in $($TestResult.ADExpireDays) days." -ForegroundColor Yellow
            Write-Host "Run command 'Set-SecretCredential -Alias <Alias> -ADUpdate' to update password in local and AD both."
        }

        $Object = FindCredential -Alias $Alias

        if($Object) {
            $password = $Object.EncryptedPwd | ConvertTo-SecureString
            $credential = New-Object -TypeName pscredential -ArgumentList $Object.UserName, $password

        }

    }

    return $credential

}

function Show-SecretCredential {
<#
.Synopsis
List all local credential information.
#>
    $Objects = GetJsonPSObject

    $Objects | Format-Table -AutoSize -Property Alias, UserName, DateTime

}

# nested functions

function GetJsonPSObject {
    # read json file and convert content to PSObject
    if(Test-Path $Global:SecretPath) {
        $str = Get-Content $Global:SecretPath
        $TimeZone = Get-TimeZone
        $JsonObjs = $str | ConvertFrom-Json
        $JsonObjs | ForEach-Object {$_.DateTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($_.DateTime, $TimeZone)}

        return $JsonObjs
    }
}

function FindCredential {
    # find credential object from PS Objects.
    param(
        [parameter(ParameterSetName='username')]
        [String]$UserName,
        [parameter(ParameterSetName='alias')]
        [String]$Alias
    )

    $Objects = GetJsonPSObject

    if($PSCmdlet.ParameterSetName -eq "username") {
        $obj = $Objects | Where-Object UserName -eq $UserName
    }

    if($PSCmdlet.ParameterSetName -eq "alias") {
        $obj = $Objects | Where-Object Alias -eq $Alias
    }

    return $obj
}

function MakePSObject {

    param(
        [String]$Alias,
        [String]$UserName,
        [System.Security.SecureString]$Password
    )

    $Encryptedpwd = $Password | ConvertFrom-SecureString
    $DateTime = Get-Date

    $p = @{
        Alias = $Alias
        UserName = $UserName
        EncryptedPwd = $Encryptedpwd
        DateTime = $DateTime
    }
    
    return New-Object -TypeName PSObject -Property $p

}