<#
Jisc SBC baseline setup for testing.
#>

#region:variables

$TestNumbers = @(
"+442037636566"
"+442037636567"
)

$baselicense = 
$licensetoassign = "MCOEV_FACULTY" #Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits
$Location = "GB"
$Password = "OrangeTurtleWhiskers"

$CustFQDN = "wmc.sbc.jisc.ac.uk"

#Provsioning accounts to be created first
$provisioningAccounts = @(
"voice.test1@$CustFQDN"
"voice.test2@$CustFQDN"
)

#endregion

#region:check and install required modules if required
Write-Host "Checking that prerequisite modules are installed - please wait."
$Module = Get-Module -ListAvailable 
if (! $($Module | Where-Object -property name -like "microsoftteams*")) {
  Write-Host "MS Teams module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name MicrosoftTeams -force -Scope CurrentUser
  }
  else {
    Write-Host "MS Teams module is required. Please install module using Install-Module MicrosoftTeams cmdlet."
    Pause
    Exit
  }
}

#Not using mg as not request to add enterprise app completed
Connect-MicrosoftTeams
Connect-AzureAD

$TenantLicenses = Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits
$SKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq $licensetoassign).skuid

$TenantLicenses = Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits
$BaseSKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq $baselicense).skuid

#endregion

#region:1. Create Provisioning Accounts - these are licensed accounts with the SBC domain as the UPN

#Setup and test service with test accounts this has not been tested - please delete this text when it has!!!!!!!!!!
#Create the new user account with license

$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.EnforceChangePasswordPolicy = $false
$PasswordProfile.ForceChangePasswordNextLogin = $false
$PasswordProfile.Password = $Password

foreach ($User in $provisioningAccounts){
    $usertest = get-azureaduser -ObjectId $User
    if (! $usertest){
        $account = New-AzureAdUser -UserPrincipalName $user -Department "SBC Provisionig Account" -AccountEnabled $true -DisplayName $user -PasswordProfile $PasswordProfile -MailNickName $($user.Split('@')[0]) -PasswordPolicies "DisablePasswordExpiration, DisableStrongPassword"
        #Add a license to the new account
        $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License.SkuId = $BaseSKUID
        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License

        $License2 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License2.SkuId = $SKUID
        $LicensesToAssign.AddLicenses = $License2

        Set-AzureADUser -ObjectID $($Account.objectid) -UsageLocation $Location
        Set-AzureADUserLicense -ObjectId $($Account.objectid) -AssignedLicenses $LicensesToAssign
    }
    else{
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

foreach ($User in $provisioningAccounts){
   
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

foreach ($number in $TestNumbers){
    $targetID = get-csonlineuser $(Get-CsPhoneNumberAssignment -NumberType directrouting -TelephoneNumber $number).AssignedPstnTargetId
    if ($targetID){Remove-CsPhoneNumberAssignment -Identity $($targetID.UserPrincipalName) -RemoveAll}
}

#endregion