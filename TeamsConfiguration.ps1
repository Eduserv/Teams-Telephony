<#
Note that this script should only be used after the 'JiscStandardElements' have been implemented.

Jisc teams configuration currently includes the following elements:

Caller ID Policies
Users
CAP
Call Queues
AutoAttendant - Resource accounts and numbers only
Rooms

#>

#region: Functions

function Assign-Number {
    [CmdletBinding()]
	    param(
		    [Parameter(Mandatory)]
		    [string]$Number,


            [Parameter(Mandatory)]
		    [string]$TargetUPN
	    )

    Try{        
        Set-CsPhoneNumberAssignment -Identity $TargetUPN -PhoneNumber $Number -PhoneNumberType DirectRouting -LocationId $null -erroraction stop
    }
    Catch{
        Write-Host "Failed to associate number $Number to user: $TargetUPN" -ForegroundColor red
        $CurrentAssignment = get-csonlineuser $(Get-CsPhoneNumberAssignment -NumberType directrouting -TelephoneNumber $Number).AssignedPstnTargetId
        if (!$CurrentAssignment){
            Write-Host "Number currently not associated - try again later (if you've only just licensed the user simply run the code again in a few minutes)" -ForegroundColor red 
            return
        }
        Write-Host "Number currently associated with: $($currentAssignment.UserPrincipalName)" -ForegroundColor red 
        #option here to move number
        $confirmation = Read-host "Would you like to transfer the number? (Y/N)"
        if ($confirmation -eq 'y') {
            Write-Host "Move number selected" -ForegroundColor yellow
            Remove-CsPhoneNumberAssignment -Identity $($currentAssignment.UserPrincipalName) -RemoveAll
            while ($CurrentAssignment){
                $CurrentAssignment = get-csonlineuser $(Get-CsPhoneNumberAssignment -NumberType directrouting -TelephoneNumber $Number).AssignedPstnTargetId -erroraction SilentlyContinue
                Start-Sleep -Seconds 10
                }
            Set-CsPhoneNumberAssignment -Identity $TargetUPN -PhoneNumber $Number -PhoneNumberType DirectRouting
            Write-Host "Number move complete" -ForegroundColor green
        }
    }
}

Function Add-License {
    [CmdletBinding()]
	    param(
		    [Parameter(Mandatory)]
		    [string]$SkuID,

            #[Parameter(Mandatory)]
		    #[string]$Group,

            [Parameter(Mandatory)]
		    [string]$TargetUPN,

            [Parameter(Mandatory)]
		    [string]$Location

            #add param here for group name and expand function to simply add to group
	    )

        #ensure spare licenses are avaliable
        $LicenseDetails = $(Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits,@{N="Total";E={$_.prepaidunits.enabled}}) | ? -Property Skuid -eq $skuid
        if ($($LicenseDetails.ConsumedUnits) -lt $($LicenseDetails.Total)){
             Write-Host "Spare licenses available - attempting to add" -ForegroundColor green 
        }
        else{
            Write-Host "No spare licenses available" -ForegroundColor red 
            return
        }

        #Add user to AD group or add license direct
        #AD group add here please

        #$confirmation = Read-host "Would you like to try and add a license (Licenses are ideally added using AAD groups)? (Y/N)"
        $confirmation = "y" #remove this if the above line is uncommented
        if ($confirmation -eq 'y') {
 
            $ADUserObject = Get-AzureADUser -ObjectId $TargetUPN | select assignedlicenses,objectid
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuId = $SKUID
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License
            Set-AzureADUser -ObjectID $($ADUserObject.ObjectId) -UsageLocation $Location
            Set-AzureADUserLicense -ObjectId $($ADUserObject.ObjectId) -AssignedLicenses $LicensesToAssign
            Write-Host "Licence applied for: $TargetUPN" -ForegroundColor yellow
        } 
}

#endregion

#region: License friendly names - need to complete this list but dont really use it yet!!
#https://docs.microsoft.com/en-us/microsoftteams/teams-add-on-licensing/assign-teams-add-on-licenses
$FriendlyLicenseNames = New-Object psobject -property @{
ENTERPRISEPREMIUM = "Enterprise E5 (with Audio Conferencing)"
MCOEV = "Microsoft Phone System Plan"
SPE_E5 = "Microsoft 365 E5"
MCOCAP_FACULTY = "Common Area Phone for faculty"
PHONESYSTEM_VIRTUALUSER_FACULTY = "Microsoft Teams Phone Resource Account for faculty"
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

#region: Check and install required modules if required
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

if (! $($Module | Where-Object -property name -eq "ImportExcel")) {
  Write-Host "ImportExcel module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name ImportExcel -force -Scope CurrentUser
  }
  else {
    Write-Host "ImportExcel module is required. Please install module using Install-Module ImportExcel cmdlet."
    Pause
    Exit
  }
}

if (! $($Module | Where-Object -property name -eq "AzureAD")) {
  Write-Host "AzureAD module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name AzureAD -force -Scope CurrentUser
  }
  else {
    Write-Host "AzureAD module is required. Please install module using Install-Module AzureAD cmdlet."
    Pause
    Exit
  }
}

if (! $($Module | Where-Object -property name -eq "ExchangeOnlineManagement")) {
  Write-Host "ExchangeOnlineManagement module is not available" -ForegroundColor yellow
  $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
  if ($Confirm -match "[yY]") {
    Install-Module -Name ExchangeOnlineManagement -force -Scope CurrentUser
  }
  else {
    Write-Host "ExchangeOnlineManagement module is required. Please install module using Install-Module ExchangeOnlineManagement cmdlet."
    Pause
    Exit
  }
}


#endregion

#region: Connect

Connect-MicrosoftTeams #$(get-cstenant).displayname
Connect-AzureAD # $(get-azureadtenantdetail).displayname
Connect-ExchangeOnline

#endregion

#region: Set Variables for the script

# Users
$licensetoassign = "MCOEV_FACULTY"
$TenantLicenses = Get-AzureADSubscribedSku | Select -Property Sku*,ConsumedUnits,@{N="Total";E={$_.prepaidunits.enabled}}
$SKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq $licensetoassign).skuid

# Common Area Phones
$CAPGroupName = "SecGrp Common Area Phone Resource Accounts"
$CAPPassword = "GreenWaspHorseEars1!" #Current this password is used from CAP (Common Area Phones) and TRA (Teams Resource Accounts)
$CAPSKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq "MCOCAP_FACULTY").skuid # this needs to be changed to the common area phone license sku
$Location = "GB"

# Teams Resource Accounts
$TRAGroupName = "SecGrp Teams Resource Accounts"
$TRASKUID = $($TenantLicenses | ? -Property SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER_FACULTY").skuid #skuid = PHONESYSTEM_VIRTUALUSER

#endregion 

#region: Test excel master spreadsheet

$excelfile = "C:\Users\Simon.Dix\OneDrive - Jisc\Desktop\Sparsholt teams config.xlsx"
#$excelfile = "C:\Users\Simon.Dix\OneDrive - Jisc\Desktop\WMC teams configuration.xlsx"
$excelFileManual = Read-Host "Please provide path to the source Excel workbook or click enter to accept $excelfile"

If ($excelFileManual) {$excelfile = $excelFileManual}

$exist = Test-Path $excelFile
If (!$exist) {
  Write-Host "Excel file does not exist at that path - please re-run the script and enter the correct path"
  exit
}


#endregion

#region: Caller ID Policies
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Caller ID Policies)"
if ($workSheet -eq ""){$workSheet = "Caller ID Policies"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

foreach ($policy in $source){
    if ($policy.'Policy Name' -eq $null){continue}
    
    $CallIDPol = Get-CsCallingLineIdentity -Identity $policy.'Policy Name' -erroraction SilentlyContinue

    If ($CallIDPol -eq $null){
        Write-Host "Could not find CallerID Policy: $($policy.'Policy Name')" -ForegroundColor yellow
        Write-Host "Creating CallerID Policy: $($policy.'Policy Name')" -ForegroundColor green
        
        $allowoverride = [system.convert]::ToBoolean($($policy.'User Override'))
        $blockIncoming = [system.convert]::ToBoolean($($policy.'BlockIncomingID'))

        if ($policy.CallingIDSubstitute -eq "Resource"){
            $ObjId = (Get-CsOnlineApplicationInstance -Identity $($policy.ResourceAccount)).ObjectId
            if ($ObjId -ne $null){
                New-CsCallingLineIdentity -Identity $($policy.'Policy Name') `
                -Description $policy.Description `
                -EnableUserOverride $allowoverride `
                -CallingIDSubstitute $policy.CallingIDSubstitute `
                -BlockIncomingPstnCallerID $policy.BlockIncomingID `
                -ResourceAccount $objid
            }
            else{
                Write-Host "Cannot find resource account: $($policy.ResourceAccount) - Please check that the resource account exists - you may need to run other builds and return to this." -ForegroundColor red          
            }
        }

        if ($policy.CallingIDSubstitute -eq "Service"){
            New-CsCallingLineIdentity -Identity $($policy.'Policy Name') `
            -Description $policy.Description `
            -EnableUserOverride $allowoverride `
            -ServiceNumber $policy.ServiceNumber `
            -CallingIDSubstitute $policy.CallingIDSubstitute `
            -BlockIncomingPstnCallerID $policy.BlockIncomingID
        }
    }
    else{
        Write-Host "Found existing CallerID Policy: $($policy.'Policy Name')" -ForegroundColor green
    }
}

#endregion

#region: Users

#bring in the data for users
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Telephony Enabled Users)"
if ($workSheet -eq ""){$workSheet = "Telephony Enabled Users"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

$UserProcessingError = @()
$count = 0
$UserInput = ""

Foreach ($user in $source){
    #Progress
    $upn = $user.'User Name (UPN)'
    $count ++
    Write-Progress -Activity "Processed user count: $count of $($source.count) Currently Processing: $upn" -PercentComplete $($Count / $($source.count) * 100)
    Write-host "________________" -ForegroundColor DarkYellow
    
    #Find the user and check licenses
    $ADUser = ""
    $ADUser = Get-AzureADUser -ObjectId $upn | select assignedlicenses,objectid
    if(!$ADUser){
         Write-Host "Can't find User: $upn" -ForegroundColor yellow
         $UserProcessingError += $upn
         continue
    }
    else{
        Write-Host "Found existing User: $upn" -ForegroundColor green
    }
  # } #you can uncomment here and run lines 143 to 158 if you just want to check the users upn's
    
    #This code makes sure we allow for organisations who are licensing via AAD groups (which they ideally will be)
    if ($skuID -notin $($aduser.AssignedLicenses.skuid)){
        Write-Host "Licence not found for: $upn" -ForegroundColor yellow
        try{
            Add-License -SkuID $SKUID -TargetUPN $upn -Location $Location -ErrorAction stop
        }
        catch{
            Write-Host "Error assigning license to user $UPN - continuing to next user" -ForegroundColor red
            continue
        }
    }
    else{Write-Host "Found Licence: $licensetoassign assignd to $upn" -ForegroundColor Green}

    #Enable users for enterprise voice
    $CSUser = Get-CsOnlineVoiceUser -Identity $upn -erroraction SilentlyContinue

    if ($CSUser.enterprisevoiceenabled -ne "True"){
        Write-Host "Enabling user: $upn for ent voice" -ForegroundColor yellow
        try{
            Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true -erroraction Stop
        }
        catch{
            Write-Host "Error Enabling user: $upn for ent voice - skipping user (if they have just been licensed there can be a long delay)" -ForegroundColor red
            if ($UserInput.Length -lt 2){
                $UserInput = read-host "Wait (W) or Continue (C) to re-run script later (WA or CA will set wait or continue for all further users)?"
            }
            if (($UserInput -eq "w") -or ($UserInput -eq "WA")){
                while ($CSUser.enterprisevoiceenabled -ne "True"){
                    start-sleep -s 5    
                    Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true -erroraction SilentlyContinue
                    $CSUser = Get-CsOnlineVoiceUser -Identity $upn -erroraction SilentlyContinue
                }
            }
            else{
                Write-Host "Error Enabling user: $upn for ent voice - skipping user" -foregroundcolor yellow
                $UserProcessingError += $upn
                continue
            }
            Write-Host "User: $upn enabled for ent voice after wait" -ForegroundColor Green
        }
    }
    else{
        Write-Host "User: $upn already has ent voice enabled" -ForegroundColor Green
    }

    #Check users phone number

    Write-Host "Checking number assignment for user: $upn" -ForegroundColor Green

    if ($($user.'Direct Dial').length -gt 5){
        if ($($user.'Direct Dial').Replace("+","") -ne $($CSUser.Number)){
            Assign-Number -TargetUPN $upn -Number $($user.'Direct Dial')
        }
        else{
            Write-Host "Number assignment correct for user: $upn" -ForegroundColor Green
        }
    }
    if ($($user.'External Number').Length -gt 5 -and $($user.'Internal Extension').length -ge 1){
        $PhoneNumber = $($user.'External Number' + ";ext=" + $user.'Internal Extension') 
        Assign-Number -TargetUPN $upn -Number $PhoneNumber
    }
    
    <#Check users teams upgrade policy
    $csuserteamspol = $(Get-CsOnlineUser -Identity $upn).TeamsUpgradePolicy

    if ($csuserteamspol -eq $null){
        Write-Host "User: $upn is configured to use the GLOBAL 'Teams upgrade settings'" -ForegroundColor green
    }
    if ($($csuserteamspol.name) -eq "UpgradeToTeams"){
        Write-Host "User: $upn is configured DIRECTLY to use Teams only" -ForegroundColor green
    }
    if (($($csuserteamspol.name) -ne "UpgradeToTeams") -and ($csuserteamspol -ne $null)){
        Write-Host "User: $upn has unknown Teams upgrade settings please check" -ForegroundColor red
        #we may wish to inject code here to resolve this state
    }
    #>

    #assign call routing policy - add a check here rather than just change eachtime
    if($user.'Call Routing Policy'){
        Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $($user.'Call Routing Policy')
    }

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

#endregion users

#region: common area phone resource accounts

<#
https://docs.microsoft.com/en-us/microsoftteams/set-up-common-area-phones
http://blog.schertz.name/2020/04/common-area-phones-in-microsoft-teams/
Remote sign in: https://docs.microsoft.com/en-us/microsoftteams/devices/remote-provision-remote-login
#>

$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Common Area Phones)"
if ($workSheet -eq ""){$workSheet = "Common Area Phones"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

    #region Create CSV file to upload into Teams
    #Adding phones to Teams is a manual process with verification required on the phones. 
    #Also the phone will require Intune configuration for Android devices (device administrator enrollments enabled) 

    $CSVPhoneExport = @()
                                foreach ($Phone in $source){
    $MACID = $phone.'MAC Address'
    $Location = $phone.Location
  
    $CSVPhoneExport += New-Object psobject -property @{
        "MAC Id"  = $MACID
        "Location" = $Location
    }
    }
    $CSVPhoneExport | Export-Csv -Path "C:\Temp\Phones.csv" -force -NoTypeInformation
    #endregion Create CSV file to upload into Teams

    #region Check and create the default CAP security group for CAP accounts
    $GroupCheck = Get-AzureADGroup -SearchString $CAPGroupName
    if (!$GroupCheck){ 
    write-host "Can not find the specified AAD group for the Common Area Phone accounts" -ForegroundColor Yellow
    write-host "Creating: $CAPGroupName" -ForegroundColor Yellow
    $groupCheck = New-azureadgroup -DisplayName $CAPGroupName -SecurityEnabled $true -MailEnabled $false -MailNickName "NotSet" -Description "This group contains resource accounts for Teams Common Access Phones"
    }
    else{
        write-host "Found CAP group: $CAPGroupName" -ForegroundColor Green
    }

    #endregion Check and create the default CAP security group for CAP accounts

    #region Check/Create CA Phone User Acounts

    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.EnforceChangePasswordPolicy = $false
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $PasswordProfile.Password = $CAPPassword

    $caperrors = @()

    $count = 0

    foreach ($CAP in $source){
        if($($cap.'User (UPN)').length -eq 0){continue} #skip blank records from bottom of import if we get any
        #Progress
   
        $count ++
        Write-Progress -Activity "Processed count: $count of $($source.count) Currently Processing: $($cap.'User (UPN)')" -PercentComplete $($Count / $($source.count) * 100)

        try{
            get-azureaduser -ObjectId $cap.'User (UPN)' -ErrorAction Stop | Out-Null
            Write-host "Found user: $($cap.'User (UPN)')" -ForegroundColor Green
        }
        Catch{
            Write-host "Can not find CAP user: $($cap.'User (UPN)')" -ForegroundColor Yellow
            Write-host "Creating CAP user: $($cap.'User (UPN)')" -ForegroundColor Yellow
            try{
                $mailnickname = $($cap.'User (UPN)'.Split('@')[0]) -replace '[\W]',''
                $account = New-AzureAdUser -UserPrincipalName $cap.'User (UPN)' -Department "Common Area Phone Accounts" -AccountEnabled $true -DisplayName $cap.DisplayName -PasswordProfile $PasswordProfile -MailNickName $mailnickname -PasswordPolicies "DisablePasswordExpiration, DisableStrongPassword" -ErrorAction Stop
                #Add new account to the CAP secuity group
                Add-AzureADGroupMember -ObjectId $($groupcheck.objectid) -RefObjectId $($Account.objectid)
                #Add a license to the new account
                try{
                    Add-License -SkuID $CAPSKUID -TargetUPN $cap.'User (UPN)' -Location $Location -ErrorAction stop
                }
                catch{
                    Write-Host "Error assigning license to user $UPN - continuing to next user" -ForegroundColor red
                    continue
                }
            }
            catch{
                Write-host "Error creating CAP user: $($cap.'User (UPN)') - its likely you have unsupported characters in UPN - please resolved in source" -ForegroundColor red
                $caperrors += $cap.'User (UPN)'
            }
        }
    }

    if($caperrors -gt 0){
            Write-host "$($caperrors.count) errors generated when creating CAP users - its likely you have unsupported characters in UPN - please resolved in source and re-run code" -ForegroundColor red
    }

    #endregion Check/Create CA Phone User Acounts

    #region Check/Configure CA Phone User Acount details
    
    <#Assign these CAP users enterprise voice and voice routing policy and IPPhone Policy - 
    I've found that sometimes you have to wait until the accounts are ready to be enabled and then a pause before you can add a number.
    You can keep running this loop until all errors resolve themselves. Ive avoided putting in a wait here as this will stop you moving onto further config.
    Instead just re-trying after 30 minutes seems better.
    #>

    foreach ($CAP in $source){
        if($($cap.'User (UPN)').length -eq 0){continue} #skip blank records from bottom of import if we get any
        Write-host "_____________________________________" -ForegroundColor DarkMagenta
        $upn = $cap.'User (UPN)'
        $csonlinevoiceuser = Get-CsOnlinevoiceUser -Identity $upn | select *

        #check the user has enterprise voice enabled
        if ($($csonlinevoiceuser.enterprisevoiceenabled) -ne "True"){
            Write-Host "Enabling user: $upn for ent voice" -ForegroundColor yellow
            Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true
        }
        else{
            Write-Host "User: $upn already has ent voice enabled" -ForegroundColor Green
        }

        #move onto checking and assigning a number

        Write-Host "Completing number assignment for user: $upn" -ForegroundColor Green
        if ($($cap.'Direct Dial').length -gt 5){
            if ($($csonlinevoiceuser.number) -ne $($cap.'Direct Dial').Replace("+","")){
                Assign-Number -TargetUPN $upn -Number $($cap.'Direct Dial')
            }
            if ($($csonlinevoiceuser.number) -eq $($cap.'Direct Dial').Replace("+","")){
                Write-Host "Number: $($cap.'Direct Dial') correctly associated with: $upn" -ForegroundColor yellow  
            }
        }
        if ($($cap.'External Number').Length -gt 5 -and $($cap.'Internal Extension').length -ge 1){
            $PhoneNumber = $($cap.'External Number' + ";ext=" + $cap.'Internal Extension') 
            Assign-Number -TargetUPN $upn -Number $PhoneNumber
        }

        #Assignment of various policies is completed below

        #assign call routing policy
        Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $($cap.'Call Routing Policy')

        #assign caller IP Policy
        if($cap.'Caller ID Policy'){
            #Check that a caller ID policy exists for the configuration
            if ($(Get-CsCallingLineIdentity -Identity $($cap.'Caller ID Policy'))){
                #Assign the policy
                Grant-CsCallingLineIdentity -PolicyName $($cap.'Caller ID Policy') -Identity $upn
            }
        }

        #assign IP Phone Policy
        if($cap.IPPhonePolicy){
            Grant-CsTeamsIpPhonePolicy -Identity $upn -PolicyName $($cap.IPPhonePolicy)
        } 
    }

    #endregion Check/Configure CA Phone User Acount details

    #region Change CAP account password utility
    
    #This utility enables the client to mass update the CAP password
    <#
    $cappassworderrors = @()

    foreach ($CAP in $source){
        if($($cap.'User (UPN)').length -eq 0){continue} #skip blank records from bottom of import if we get any
        try{
            get-azureaduser -ObjectId $cap.'User (UPN)' -ErrorAction Stop | Out-Null
            Write-host "Found user: $($cap.'User (UPN)')" -ForegroundColor Green
            Set-AzureADUserPassword -ObjectId $cap.'User (UPN)' -Password $(convertto-securestring -string $CAPPassword -AsPlainText -Force) -ForceChangePasswordNextLogin $false -EnforceChangePasswordPolicy $false
            Write-host "Updated user: $($cap.'User (UPN)')" -ForegroundColor Green
        }
        Catch{
            Write-host "Cannot find or update user: $($cap.'User (UPN)')" -ForegroundColor red
            $CAPPassword += $($cap.'User (UPN)')
        }
    }

    if($cappassworderrors -gt 0){
        Write-host ""
        Write-host "$($cappassworderrors.count) errors generated trying to update CAP account passwords" -ForegroundColor red
    }
    else {
        Write-host ""
        Write-host "No errors generated trying to update CAP account passwords" -ForegroundColor green
    }
    #>
    #endregion Change CAP account password utility

#endregion common area phone resource accounts

#region: call queues

<#
This region offers functionality but not developed so lots to do here..
This will 
-create a resource account and add it to a group (CA)
-associate an external number
-setup a call queue with the required delivery
This WILL NOT (currently)
-configure timeout and overflow actions - apart from basic options when using a 365 group
-deal with audio and welcome messages - apart from basic options when using a 365 group
#>

$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Call Queues)"
if ($workSheet -eq ""){$workSheet = "Call Queues"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

#https://docs.microsoft.com/en-us/microsoftteams/create-a-phone-system-call-queue-via-cmdlets
#obtain resource account licenses: https://docs.microsoft.com/en-us/microsoftteams/manage-resource-accounts#obtain-microsoft-teams-phone-resource-account-licenses

#Create and license the resource accounts and create the default TRA security group for TRA accounts

$GroupCheck = Get-AzureADGroup -SearchString $TRAGroupName
if (!$GroupCheck)
{ 
    $groupCheck = New-azureadgroup -DisplayName $TRAGroupName -SecurityEnabled $true -MailEnabled $false -MailNickName "NotSet" -Description "This group contains resource accounts for Teams"
}


#Account Config
Write-host "Checking that resource accounts are correctly configured for the call queues" -ForegroundColor Green

foreach ($QResource in $source){
    if ($qresource.'Call Queue Name' -eq $null){Continue} #sometimes the excel import has blanks
    Write-host "_____________________________________" -ForegroundColor DarkMagenta
    $Account = Get-CsOnlineApplicationInstance -Identity $QResource.'User (UPN)' -erroraction SilentlyContinue
    if ($Account -eq $null) {
        Write-host "Creating resource account for the call queue: $($QResource.'Call Queue Name') " -ForegroundColor Green
        $Account = New-CsOnlineApplicationInstance -UserPrincipalName $QResource.'User (UPN)' -DisplayName $QResource.DisplayName -ApplicationID "11cd3e2e-fccb-42ad-ad00-878b93575e07"
        Add-AzureADGroupMember -ObjectId $($groupcheck.objectid) -RefObjectId $($Account.objectid)
        #Add a license to the new account - pauses to ensure account is created
        Start-Sleep -Seconds 5
        try{
            Add-License -SkuID $TRASKUID -TargetUPN $QResource.'User (UPN)' -Location $Location -ErrorAction stop
        }
        catch{
            Write-Host "Error assigning license to user $UPN - continuing to next user" -ForegroundColor red
            continue
        }
    }
    else{
        Write-host "Found user: $($QResource.'User (UPN)')" -ForegroundColor Green
        $ADuser = Get-AzureADUser -ObjectId $($QResource.'User (UPN)') | select assignedlicenses,objectid
        if ($TRASKUID -notin $($aduser.AssignedLicenses.skuid)){
            Write-Host "Licence not found for: $($QResource.'User (UPN)')" -ForegroundColor yellow
            #Add a license to the new account
            try{
                Add-License -SkuID $TRASKUID -TargetUPN $QResource.'User (UPN)' -Location $Location -ErrorAction stop
            }
            catch{
                Write-Host "Error assigning license to user $UPN - continuing to next user" -ForegroundColor red
                continue
            }
        }
        else{Write-Host "Found Licence for: $($QResource.'User (UPN)')" -ForegroundColor Green}
    }

    Write-Host "Completing number assignment for Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green

    if ($($QResource.'Ext Phone Number').length -gt 5){
        #Test if the user is ready for a phone number
        

        if ($Account.PhoneNumber -ne $("tel:" + $QResource.'Ext Phone Number').trim()){
            #for smooth setup seems we need another pause here
            start-sleep -Seconds 5
            Assign-Number -TargetUPN $($QResource.'User (UPN)') -Number $($QResource.'Ext Phone Number')
        }
        else{
            Write-Host "Number: $($QResource.'Ext Phone Number') correctly associated with: $($QResource.'User (UPN)')" -ForegroundColor yellow  
        }
    }
}

#Create the Q
foreach ($QResource in $source){
    if ($qresource.'Call Queue Name' -eq $null){Continue} #sometimes the excel import has blanks
    Write-host "_____________________________________" -ForegroundColor DarkMagenta
    Write-host "Creating / Checking Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green

    $Name = $QResource.'Call Queue Name'
    $AgentAlertTime = $QResource.AlertTime
    $RoutingMethod = $QResource.'Routing Method'
    $PresenceRouting = [system.convert]::ToBoolean($($QResource.PresenceBasedRouting))
    $AllowOptOut = [system.convert]::ToBoolean($($QResource.AllowOptOut))

    #$audioFileQueueMusicID = ""
    #$content = Get-Content "d:\sales-hold-in-queue-music.wav" -Encoding byte -ReadCount 0
    #$audioFileSalesHoldInQueueMusicID = (Import-CsOnlineAudioFile -ApplicationID HuntGroup -FileName "sales-hold-in-queue-music.wav" -Content $content).ID

    $greeting = $QResource.Greeting
    $TimeOutTarget = $QResource.TimeoutHandling #id of target

    #delivery to a user list maintained in teams
    if(($($QResource.'Members of Queue UPN') -ne $null) -and ($($QResource.Group) -eq $null)){
        #Check if Queue already exists
        Write-host "Checking (uuser list) Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green
        $queue = get-cscallqueue -name $name
        if($queue){
            Write-host "Found Call Queue: $($QResource.'Call Queue Name') - for now please manually update configuration" -ForegroundColor Green
            continue
        }
        
        #CollectUserIDs
        Write-host "Creating (user list) Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green
        continue
        $users = $($QResource.'Members of Queue UPN').split(';').trim()
        $userids = @()
        foreach ($user in $users){
            #Add user ID's to array
            $CSUser = Get-CsOnlineUser -Identity $user
            if($csuser){$userids += $CSuser.identity}
        }

        New-CsCallQueue -Name $name `
        -PresenceBasedRouting $PresenceRouting `
        -UseDefaultMusicOnHold $true `
        -AgentAlertTime $AgentAlertTime `
        -RoutingMethod $RoutingMethod `
        -ConferenceMode $true `
        -LanguageID "en-GB" `
        -AllowOptOut $AllowOptOut `
        -User $userids

        #-TimeoutActionTarget $TimeOutTarget `
        #-TimeoutAction Forward `
        #-TimeoutThreshold 120 `
        #-OverflowAction Forward `
        #-OverflowActionTarget $TimeOutTarget `
        #-OverflowThreshold 200 `
        #-WelcomeMusicAudioFileId $greeting `

        #link the resource account
        Start-Sleep -Seconds 5 #just makesure the queue has been created
        $applicationInstanceID = (Get-CsOnlineUser -Identity $($QResource.'User (UPN)')).Identity
        $callQueueID = (Get-CsCallQueue -NameFilter $($QResource.'Call Queue Name')).Identity
        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue
    }


    #delivery to group
    if ($($QResource.Group) -ne $null){
        Write-host "Checking (unified group) Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green
        Write-host "Checking unified group: $($QResource.Group)" -ForegroundColor Green

        $group = Get-UnifiedGroup -Identity $($QResource.Group) -ErrorAction silentlycontinue
        $users = $($QResource.'Members of Queue UPN').Split(';') #the first user is made the group owner

        if ($group -eq $null)
        {
            Write-host "Creating unified group: $($QResource.group)" -ForegroundColor Green
            $group = New-UnifiedGroup -DisplayName $($QResource.Group) -AccessType private -Owner $users[0]  #-SuppressWarmupMessage
            start-sleep -Seconds 5

            Set-UnifiedGroup -Identity $($QResource.Group) -HiddenFromAddressListsEnabled $true #comment this out if customer doesn't want groups hidden

            $users = $($QResource.'Members of Queue UPN').Split(';')
            foreach ($user in $users){
                if ($user -eq ""){continue}
                #check user is voice enabled
                $CSUser = Get-CsOnlineVoiceUser -Identity $user
                if ($CSUser.enterprisevoiceenabled -ne "True"){
                    Add-UnifiedGroupLinks -Identity $($group.DisplayName) -LinkType members -links $user
                }
                else{
                    Write-host "Unable to add $user to unified group: $($QResource.group)" -ForegroundColor red
                }
            }
        }
        if ($group -ne $null)
        {
            Write-host "Found unified group: $($QResource.group)" -ForegroundColor Green
            if ($($group.HiddenFromAddressListsEnabled) -eq $False){
                Set-UnifiedGroup -Identity $($QResource.Group) -HiddenFromAddressListsEnabled $true #comment this out if customer doesn't want groups hidden
            }

            $CurrentMembers = $(Get-UnifiedGroupLinks -Identity $($QResource.Group) -LinkType members).windowsliveid
            $CurrentOwners = $(Get-UnifiedGroupLinks -Identity $($QResource.Group) -LinkType owners).windowsliveid
            $differences = Compare-Object -ReferenceObject $CurrentMembers -DifferenceObject $users

            if ($users[0] -notin $CurrentOwners){
                Add-UnifiedGroupLinks -Identity $($group.DisplayName) -LinkType owners -links $users[0]
            }

            foreach ($user in $($differences | ? -Property sideindicator -eq "=>").inputobject){
                if ($user -eq ""){continue}
                Write-host "Attempting to add $user to unified group: $($QResource.group)" -ForegroundColor green
                #check user is voice enabled
                $CSUser = Get-CsOnlineVoiceUser -Identity $user
                if ($CSUser.enterprisevoiceenabled -eq "True"){
                    Add-UnifiedGroupLinks -Identity $($group.DisplayName) -LinkType members -links $user
                }
                else{
                    Write-host "Unable to add $user to unified group: $($QResource.group)" -ForegroundColor red
                }
            }
            foreach ($user in $($differences | ? -Property sideindicator -eq "<=").inputobject){
                if ($user -eq ""){continue}
                Write-host "Attempting to remove $user to unified group: $($QResource.group)" -ForegroundColor yellow
                Remove-UnifiedGroupLinks -Identity $($group.DisplayName) -LinkType owner -links $user -Confirm:$false -ErrorAction silentlycontinue
                Remove-UnifiedGroupLinks -Identity $($group.DisplayName) -LinkType members -links $user -Confirm:$false
            }
        }

        $groupid = $(Get-UnifiedGroup -Identity $($QResource.Group)).ExternalDirectoryObjectId

        $queue = get-cscallqueue -name $name
        if($queue){
            Write-host "Found Call Queue: $($QResource.'Call Queue Name') - for now please manually update configuration" -ForegroundColor Green
            continue
        }


        New-CsCallQueue -Name $name `
        -PresenceBasedRouting $PresenceRouting `
        -UseDefaultMusicOnHold $true `
        -AgentAlertTime $AgentAlertTime `
        -RoutingMethod $RoutingMethod `
        -ConferenceMode $true `
        -LanguageID "en-GB" `
        -AllowOptOut $AllowOptOut `
        -DistributionLists $groupID `
        -OverflowAction SharedVoicemail `
        -OverflowActionTarget $groupID `
        -OverflowThreshold 30 `
        -TimeoutAction SharedVoicemail `
        -TimeoutActionTarget $groupID `
        -TimeoutThreshold 30 `
        -TimeoutSharedVoicemailTextToSpeechPrompt "We're sorry no one is available to take your call right now." `
        -OverflowSharedVoicemailTextToSpeechPrompt "We're sorry no one is available to take your call right now." `
        -EnableTimeoutSharedVoicemailTranscription $true

        
        
        #link the resource account
        Start-Sleep -Seconds 5 #just makesure the queue has been created
        $applicationInstanceID = (Get-CsOnlineUser -Identity $($QResource.'User (UPN)')).Identity
        $callQueueID = (Get-CsCallQueue -NameFilter $($QResource.'Call Queue Name')).Identity
        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue
    }

    #delivery to team
    if ($($QResource.'Deliver to Team') -ne $null){
        Write-host "Checking (delivery to Teams channel) Call Queue: $($QResource.'Call Queue Name')" -ForegroundColor Green
        $queue = get-cscallqueue -name $name
        if($queue){
            Write-host "Found Call Queue: $($QResource.'Call Queue Name') - for now please manually update configuration" -ForegroundColor Green
            continue
        }

        $group = Get-UnifiedGroup -Identity  $QResource.'Deliver to Team'
        $groupid = $group.ExternalDirectoryObjectId
        $groupGUI = $group.Guid

        $team = get-team -DisplayName $QResource.'Deliver to Team'
        #really werid error catch below as the -displayname perameter of the get-team command seems to bring back any string matches.
        if ($team.count -gt 1){
            $team = $team | ? -Property displayname -eq $QResource.'Deliver to Team'
        }

        $channel = $QResource.Channel
        if($channel -eq ""){$channel = "General"}
        $channelID = $(Get-TeamChannel -GroupId $($Team.groupid) | ? -Property DisplayName -eq $channel).id

        if(-not $ChannelID){
            Write-host "Could not find a channel ID Skipping Q creation" -ForegroundColor Red
            continue
        }

        #This doesn't actually work - i suspect a bug in the powershell.

        New-CsCallQueue -Name $name `
        -PresenceBasedRouting $PresenceRouting `
        -UseDefaultMusicOnHold $true `
        -AgentAlertTime $AgentAlertTime `
        -RoutingMethod $RoutingMethod `
        -ConferenceMode $true `
        -LanguageID "en-GB" `
        -AllowOptOut $AllowOptOut `
        -ChannelId $channelID `
        -DistributionLists $groupID `
        -OverflowAction SharedVoicemail `
        -OverflowActionTarget $groupID `
        -OverflowThreshold 30 `
        -TimeoutAction SharedVoicemail `
        -TimeoutActionTarget $groupID `
        -TimeoutThreshold 30 `
        -TimeoutSharedVoicemailTextToSpeechPrompt "We're sorry no one is available to take your call right now." `
        -OverflowSharedVoicemailTextToSpeechPrompt "We're sorry no one is available to take your call right now." `
        -EnableTimeoutSharedVoicemailTranscription $true
        
        #link the resource account
        Start-Sleep -Seconds 5 #just makesure the queue has been created
        $applicationInstanceID = (Get-CsOnlineUser -Identity $($QResource.'User (UPN)')).Identity
        $callQueueID = (Get-CsCallQueue -NameFilter $($QResource.'Call Queue Name')).Identity
        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue
    }
}

#endregion call queues

#region: AutoAttendant Resource Account Config

<#
This region only creates a resource account for the Autoattendants and adds them to a group (CA)
Autoattendants are likely to be minimal and vary too heavily to automate beyond this
#>

$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (AutoAttendant)"
if ($workSheet -eq ""){$workSheet = "AutoAttendant"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

#Create and license the resource accounts and add to the default TRA (Teams resource account) security group

$GroupCheck = Get-AzureADGroup -SearchString $TRAGroupName
if (!$GroupCheck)
{ 
    $groupCheck = New-azureadgroup -DisplayName $TRAGroupName -SecurityEnabled $true -MailEnabled $false -MailNickName "NotSet" -Description "This group contains resource accounts for Teams"
}

Write-host "Checking/Creating resource accounts for the AutoAttendants" -ForegroundColor Green

foreach ($AAResource in $source){
    
    $Account = Get-CsOnlineApplicationInstance -Identity $($AAResource.'User (UPN)') -erroraction SilentlyContinue
    if (!$Account) {
        Write-host "Creating AA resource account: $($AAResource.name) " -ForegroundColor Green
        $Account = New-CsOnlineApplicationInstance -UserPrincipalName $AAResource.'User (UPN)' -DisplayName $AAResource.DisplayName -ApplicationID "11cd3e2e-fccb-42ad-ad00-878b93575e07" 
        Add-AzureADGroupMember -ObjectId $($groupcheck.objectid) -RefObjectId $($Account.objectid)
        #Add a license to the new account
        Start-Sleep -Seconds 5
           try{
            Add-License -SkuID $TRASKUID -TargetUPN $AAResource.'User (UPN)' -Location $Location -ErrorAction stop
        }
        catch{
            Write-Host "Error assigning license to user $UPN - continuing to next user" -ForegroundColor red
            continue
        }
    }
        else{
        Write-host "Found user: $($AAResource.'User (UPN)')" -ForegroundColor Green
    }

    Write-Host "Completing number assignment for: $($AAResource.Name)" -ForegroundColor Green

    if ($Account.PhoneNumber -ne $("tel:" + $AAResource.'Ext Phone Number')){
        Assign-Number -TargetUPN $($AAResource.'User (UPN)') -Number $($AAResource.'Ext Phone Number')
    }
    else{
        Write-Host "Number: $($AAResource.'Ext Phone Number') correctly associated with: $($AAResource.'User (UPN)')" -ForegroundColor yellow  
    }
}

#endregion AutoAttendant Resource Account Config

#region: teams rooms - Not developed
$workSheet = Read-Host "Please provide name of the source Excel worksheet - or leave blank to accept the default (Teams Rooms)"
if ($workSheet -eq ""){$workSheet = "Teams Rooms"}
$source = Import-Excel -Path $excelFile -WorkSheetName $workSheet

#Write code here for teams rooms!!! https://docs.microsoft.com/en-us/microsoftteams/rooms/with-office-365?tabs=exchange-online%2Cazure-active-directory2-password%2Cactive-directory2-license#create-a-resource-account

#endregion teams rooms - Not developed