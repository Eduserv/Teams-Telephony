<#
Jisc teams configuration currently includes the following elements.

Caller ID Policies
Users
CAP
Rooms
Call Queues

#>

#Requires -RunAsAdministrator

#region:License friendly names - need to complete this list
#https://docs.microsoft.com/en-us/microsoftteams/teams-add-on-licensing/assign-teams-add-on-licenses
$FriendlyLicenseNames = New-Object psobject -property @{
ENTERPRISEPREMIUM = "Enterprise E5 (with Audio Conferencing)"
MCOEV = "Microsoft Phone System Plan"
SPE_E5 = "Microsoft 365 E5"
}

$FriendlyLicenseNamesTable = @()
foreach ($Record in $($FriendlyLicenseNames.psobject.properties)){
    $name = $($Record.Name)
    $value = $($Record.Value)
  
    $FriendlyLicenseNamesTable += New-Object psobject -property @{
        Name = $name
        Value = $value
    }
}
#endregion

#region check and install required modules if required
Write-Host "Checking that prerequisite modules are installed - please wait."
$Module = Get-Module -ListAvailable 
if (! $($Module | Where-Object -property name -like "microsoftteams*")) {
  Write-Host "MS Teams module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name MicrosoftTeams -force
  }
  else {
    Write-Host "MS Teams module is required. Please install module using Install-Module MicrosoftTeams cmdlet."
    Pause
    Exit
  }
}

if (! $($Module | Where-Object -property name -eq ImportExcel)) {
  Write-Host "ImportExcel module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name ImportExcel -force
  }
  else {
    Write-Host "ImportExcel module is required. Please install module using Install-Module ImportExcel cmdlet."
    Pause
    Exit
  }
}
if (! $($Module | Where-Object -property name -eq AzureAD)) {
  Write-Host "AzureAD module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name AzureAD -force
  }
  else {
    Write-Host "AzureAD module is required. Please install module using Install-Module AzureAD cmdlet."
    Pause
    Exit
  }
}
#endregion

#region: Set Variables for the script

$licensetoassign = "MCOEV_FACULTY"
$TenantLicenses = Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits,prepaidunits.enabled
$SKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq $licensetoassign).skuid

$CAPGroupName = "SecGrp Common Area Phone Resource Accounts"
$CAPPassword = "GreenWaspHorseEars"
$CAPSKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq "MCOEV_FACULTY").skuid # this needs to be changed to the common area phone license sku
$Location = "GB"

#endregion 

#region test excel master spreadsheet and connect to teams
$excelFile = Read-Host "Please provide path to the source Excel workbook"

$excelfile = "C:\Users\Simon.Dix\OneDrive - Jisc\Desktop\Sparsholt.xlsx"

$exist = Test-Path $excelFile
If (!$exist) {
  Write-Host "Excel file does not exist at that path - please re-run the script and enter the correct path"
  exit
}

Connect-MicrosoftTeams
Connect-AzureAD

#endregion

#region: caller ID Policies
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Caller ID Policies)"
if ($workSheet -eq ""){$workSheet = "Caller ID Policies"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

foreach ($policy in $source){
    New-CsCallingLineIdentity -Identity $policy.'Policy Name' `
    -Description $policy.Description `
    -EnableUserOverride $policy.'User Override' `
    -ServiceNumber $policy.ServiceNumber `
    -CallingIDSubstitute $policy.CallingIDSubstitute `
    -BlockIncomingPstnCallerID $policy.BlockIncomingID `
    -ResourceAccount $policy.ResourceAccount `
    -CompanyName $policy.CompanyName
}

#endregion

#region: users

#bring in the data for users
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Telephony Enabled Users)"
if ($workSheet -eq ""){$workSheet = "Telephony Enabled Users"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

Foreach ($user in $source){
    #Find the user and check licenses
    $upn = $user.'User Name (UPN)'
    $ADUser = Get-AzureADUser -ObjectId $upn | select assignedlicenses
    if ($skuID -notin $($aduser.AssignedLicenses.skuid)){
        Write-Host "Assigning Licence: $licensetoassign to $user" -ForegroundColor yellow
        try{
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $SKUID -ErrorAction stop
        }
        catch{
            Write-Host "Error assigning license to user $upn - please add code here to ID the issue - maybe no spare lic's"
            continue
        }
            
    }
    else{Write-Host "Found Licence: $licensetoassign assignd to $upn" -ForegroundColor Green}

    #Enable users for enterprise voice

    if ($(Get-CsOnlineVoiceUser -Identity $upn).enterprisevoiceenabled -ne "True"){
        Write-Host "Enabling user: $upn for ent voice" -ForegroundColor yellow
        Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true
    }
    else{
        Write-Host "User: $upn already has ent voice enabled" -ForegroundColor Green
    }

    Write-Host "Completing number assignment for user: $upn" -ForegroundColor Green

    if ($user.'Direct Dial' -ne ""){
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $($user.'Direct Dial') -PhoneNumberType DirectRouting
    }
    if ($user.'External Number' -and $user.'Internal Extension'){
        $PhoneNumber = $($user.'External Number' + ";ext=" + $user.'Internal Extension') 
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $PhoneNumber -PhoneNumberType DirectRouting
    }

    #assign call routing policy
    Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $($user.'Call Routing Policy')

    #External Number presentation - By default, when a Teams user makes a call to a PSTN phone, the phone number of the Teams user is visible. 
    #reference: https://docs.microsoft.com/en-us/microsoftteams/caller-id-policies
    #Check what caller ID policy should be assigned (if not the global default is already assigned to user)

    if($user.'Caller ID Policy'){
        #Check that a caller ID policy exists for the configuration
        if ($(Get-CsCallingLineIdentity -Identity $($user.'Caller ID Policy'))){
            #Assign the policy
            Grant-CsCallingLineIdentity -PolicyName $($user.'Caller ID Policy') -Identity $upn
        }
    }
}

#endregion

#region: common area phone resource accounts
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Common Area Phones)"
if ($workSheet -eq ""){$workSheet = "Common Area Phones"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

#https://docs.microsoft.com/en-us/microsoftteams/set-up-common-area-phones
#http://blog.schertz.name/2020/04/common-area-phones-in-microsoft-teams/
#Remote sign in: https://docs.microsoft.com/en-us/microsoftteams/devices/remote-provision-remote-login

#Check and create the default CAP security group for CAP accounts

$GroupCheck = Get-AzureADGroup -$CAPGroupName
if (!$GroupCheck)
{ 
    $groupCheck = New-azureadgroup -DisplayName $CAPGroupName -SecurityEnabled $true -MailEnabled $false -MailNickName "NotSet" -Description "This group contains resource accounts for Teams Common Access Phones"
}

#Check for CA Phone Users and Create with license and password that doesn't expire also add to to group

$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.EnforceChangePasswordPolicy = $false
$PasswordProfile.ForceChangePasswordNextLogin = $false
$PasswordProfile.Password = $CAPPassword

foreach ($CAP in $source){
    $usertest = get-azureaduser -ObjectId $cap.'User (UPN)'
    if (! $usertest){
        $account = New-AzureAdUser -UserPrincipalName $cap.'User (UPN)' -Department "Common Area Phone Accounts" -AccountEnabled $true -DisplayName $cap.DisplayName -PasswordProfile $PasswordProfile -MailNickName $($cap.'User (UPN)'.Split('@')[0]) -PasswordPolicies "DisablePasswordExpiration, DisableStrongPassword"
        #Add new account to the CAP secuity group
        Add-AzureADGroupMember -ObjectId $($groupcheck.objectid) -RefObjectId $($Account.objectid)
        #Add a license to the new account
        $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License.SkuId = $CAPSKUID
        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License
        Set-AzureADUser -ObjectID $($Account.objectid) -UsageLocation $Location
        Set-AzureADUserLicense -ObjectId $($Account.objectid) -AssignedLicenses $LicensesToAssign
    }
    else{
        Write-host "Found user: $($cap.'User (UPN)')" -ForegroundColor Green
    }
}

#Assign these CAP users enterprise voice and voice routing policy and IPPhone Policy

foreach ($CAP in $source){
    $upn = $cap.'User (UPN)'
    if ($(Get-CsOnlineVoiceUser -Identity $upn).enterprisevoiceenabled -ne "True"){
        Write-Host "Enabling user: $upn for ent voice" -ForegroundColor yellow
        Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true
    }
    else{
        Write-Host "User: $upn already has ent voice enabled" -ForegroundColor Green
    }

    Write-Host "Completing number assignment for user: $upn" -ForegroundColor Green

    if ($cap.'Direct Dial' -ne ""){
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $($user.'Direct Dial') -PhoneNumberType DirectRouting
    }
    if ($cap.'External Number' -and $cap.'Internal Extension'){
        $PhoneNumber = $($cap.'External Number' + ";ext=" + $cap.'Internal Extension') 
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $PhoneNumber -PhoneNumberType DirectRouting
    }

    #assign call routing policy
    Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $($user.'Call Routing Policy')

    #assign caller IP Policy
    if($cap.'Caller ID Policy'){
        #Check that a caller ID policy exists for the configuration
        if ($(Get-CsCallingLineIdentity -Identity $($user.'Caller ID Policy'))){
            #Assign the policy
            Grant-CsCallingLineIdentity -PolicyName $($user.'Caller ID Policy') -Identity $upn
        }
    }

    #assign IP Phone Policy
    if($cap.IPPhonePolicy){
        Grant-CsTeamsIpPhonePolicy -Identity $upn -PolicyName $($cap.IPPhonePolicy)
    } 
}

#Add phones to Teams.

#endregion

#region: teams rooms - Not developed
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Teams Rooms)"
if ($workSheet -eq ""){$workSheet = "Teams Rooms"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

#endregion

#region: call queues - Not developed..
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Call Queues)"
if ($workSheet -eq ""){$workSheet = "Call Queues"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet
#https://docs.microsoft.com/en-us/microsoftteams/create-a-phone-system-call-queue-via-cmdlets


#skuid = PHONESYSTEM_VIRTUALUSER
#obtain resource account licenses: https://docs.microsoft.com/en-us/microsoftteams/manage-resource-accounts#obtain-microsoft-teams-phone-resource-account-licenses


#endregion
