<#
Jisc SBC baseline setup for testing.
#>

#region:variables

Write-Host "Please enter the test numbers"
$number = ""
$i = 0
$TestNumbers = @()

do {
    $i++
    Read-host "Number $i: (enter nothing to finish)" -OutVariable number
    if ($number -ne "") {
        $TestNumbers += $number
    }
} while ($number -ne "")

$skuId = "0e142028-345e-45da-8d92-8bfd4093bbb9"
$Location = "GB"
$symbols = '!@#$%^&*'.ToCharArray()
$characterList = 'a'..'z' + 'A'..'Z' + '0'..'9' + $symbols


do {
    $password = -join (0..14 | % { $characterList | Get-Random })
    [int]$hasLowerChar = $password -cmatch '[a-z]'
    [int]$hasUpperChar = $password -cmatch '[A-Z]'
    [int]$hasDigit = $password -match '[0-9]'
    [int]$hasSymbol = $password.IndexOfAny($symbols) -ne -1

}
until (($hasLowerChar + $hasUpperChar + $hasDigit + $hasSymbol) -ge 3)


$CustFQDN = Read-Host "Customer SBC FQDN (fhc.sbc.jisc.ac.uk): "

#Provsioning accounts to be created first
$provisioningAccounts = @(
"voice.test1@$CustFQDN"
"voice.test2@$CustFQDN"
)

#endregion

#Not using mg as not request to add enterprise app completed
Connect-MicrosoftTeams
Connect-MgGraph -Scopes "Directory.ReadWrite.All"

#endregion

#region:1. Create Provisioning Accounts - these are licensed accounts with the SBC domain as the UPN

#Setup and test service with test accounts this has not been tested - please delete this text when it has!!!!!!!!!!
#Create the new user account with license

$PasswordProfile = @{
    ForceChangePasswordNextSignIn = $false
    Password = $password
}

foreach ($User in $provisioningAccounts){
    Get-MgUser -UserId $User -OutVariable $usertest -ErrorAction SilentlyContinue
    if (!$usertest) {
        New-MgUser -OutVariable account -AccountEnabled:$true -Department "SBC Provisionig Account" -UserPrincipalName $User -PasswordProfile $PasswordProfile -UsageLocation $Location
        Start-Sleep -s 3
        Set-MgUserLicense -UserId $account.Id -AddLicenses @(@{ SkuId = $skuId } ) -RemoveLicenses @()
    } else {
        Write-host "Found user: $User" -ForegroundColor Green
    }
}

#endregion


#region:2. Configure Teams with basic config

Set-CsOnlinePSTNUsage -Identity Global -Usage @{add="UnrestrictedPstnUsage"}
try {
    New-CsOnlineVoiceRoute -Identity "UnrestrictedPstnUsage" -NumberPattern ".*" -OnlinePstnGatewayList $CustFQDN -OnlinePstnUsages "UnrestrictedPstnUsage" -erroraction stop
}
catch {
    $Error[0]
    Write-Host "If this command has failed with error: Cannot find specified gateway $CustFQDN then please ensure `
that a user with enterprise voice enabled has their primary SIP address aligned with this domain. If this has `
already been completed then you may have to wait for the service to be enabled by MS" -ForegroundColor yellow
}

New-CsOnlineVoiceRoute -Identity "UnrestrictedPstnUsage" -NumberPattern ".*" -OnlinePstnGatewayList $CustFQDN -OnlinePstnUsages "UnrestrictedPstnUsage"
New-CsOnlineVoiceRoutingPolicy "No Restrictions" -OnlinePstnUsages "UnrestrictedPstnUsage"

#endregion

#region:3. Finish configuring our provisioning accounts

foreach ($User in $provisioningAccounts) {
   
    #Enable users for enterprise voice

    if ($(Get-CsOnlineVoiceUser -Identity $user).enterprisevoiceenabled -ne "True"){
        Write-Host "Enabling user: $user for ent voice" -ForegroundColor yellow
        Set-CsPhoneNumberAssignment -Identity $user -EnterpriseVoiceEnabled $true
    }
    else{
        Write-Host "User: $user already has ent voice enabled" -ForegroundColor Green
    }

    Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName "No Restrictions"

}

#assign the 2 testing numbers to the accounts

Set-CsPhoneNumberAssignment -Identity $provisioningAccounts[0] -PhoneNumber $TestNumbers[0] -PhoneNumberType DirectRouting
Set-CsPhoneNumberAssignment -Identity $provisioningAccounts[1] -PhoneNumber $TestNumbers[1] -PhoneNumberType DirectRouting

#endregion


#region:4. Remove numbers

#Remove test numbers from the provisioning accounts after testing for assignment to other users

foreach ($number in $TestNumbers) {
    $targetID = get-csonlineuser $(Get-CsPhoneNumberAssignment -NumberType directrouting -TelephoneNumber $number).AssignedPstnTargetId
    if ($targetID) { Remove-CsPhoneNumberAssignment -Identity $($targetID.UserPrincipalName) -RemoveAll }
}

#endregion