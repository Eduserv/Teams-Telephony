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

$Location = "GB"
$symbols = '!@#$%^&*'.ToCharArray()
$characterList = 'a'..'z' + 'A'..'Z' + '0'..'9' + $symbols

$skus = Get-MgSubscribedSku -Select Id,skuPartNumber

$sku = $skus | ? { $_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER_FACULTY" }


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
"sbcinit@$CustFQDN"
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

$question = $host.UI.PromptForChoice("Has Teams Resource Account Licenses (free) been procured?", "Resource account licenses are used on the SBC init account", ([System.Management.Automation.Host.ChoiceDescription]"&Yes",[System.Management.Automation.Host.ChoiceDescription]"&No"), 0)

if ($question -gt 0 -or !$sku -or $sku.length -eq 0) {
    throw "Please ensure a Teams Resource Account License is available before continuing"
}

Write-Host "Creating SBC Provisioning Accounts"

foreach ($User in $provisioningAccounts) {
    Get-MgUser -UserId $User -OutVariable $usertest -ErrorAction SilentlyContinue
    if (!$usertest) {
        New-MgUser -OutVariable account -AccountEnabled:$true -Department "SBC Provisionig Account" -UserPrincipalName $User -PasswordProfile $PasswordProfile -UsageLocation $Location
        Start-Sleep -s 5
        Set-MgUserLicense -UserId $account.Id -AddLicenses @(@{ SkuId = $skuId } ) -RemoveLicenses @()
    } else {
        Write-host "Found user: $User" -ForegroundColor Green
    }
}

#endregion

#region:2. Configure Teams with basic config

$completedSBC = $false

Set-CsOnlinePSTNUsage -Identity Global -Usage @{add="UnrestrictedPstnUsage"}
while (!$completedSBC) {
    $voicerouteerror = $null
    New-CsOnlineVoiceRoute -Identity "UnrestrictedPstnUsage" -NumberPattern ".*" -OnlinePstnGatewayList $CustFQDN -OnlinePstnUsages "UnrestrictedPstnUsage" -ErrorAction Continue -ErrorVariable voicerouteerror
    if ($voicerouteerror -ilike "*Cannot find specified Gateway*") {
        Write-Host "Gateway initialization in progress, we will retry in 5 minutes"
        foreach ($i in 0..300) {
            $timeleft = ([timespan]::fromSeconds(300 - $i)).toString()
            Write-Progress -Activity "Waiting for SBC Gateway - recheck in $timeleft" -SecondsRemaining (300 - $i) -PercentComplete ($i / 300 * 100)
            Start-Sleep -s 1
        }
    }
    $completed = (!$voicerouteerror)
}

New-CsOnlineVoiceRoute -Identity "UnrestrictedPstnUsage" -NumberPattern ".*" -OnlinePstnGatewayList $CustFQDN -OnlinePstnUsages "UnrestrictedPstnUsage"
New-CsOnlineVoiceRoutingPolicy "No Restrictions" -OnlinePstnUsages "UnrestrictedPstnUsage"

#endregion

#region:3. Finish configuring our provisioning accounts

foreach ($User in $provisioningAccounts) {
   
    #Enable users for enterprise voice

    if ($(Get-CsOnlineVoiceUser -Identity $user).enterprisevoiceenabled -ne "True") {
        Write-Host "Enabling user: $user for ent voice" -ForegroundColor yellow
        Set-CsPhoneNumberAssignment -Identity $user -EnterpriseVoiceEnabled $true
    } else {
        Write-Host "User: $user already has ent voice enabled" -ForegroundColor Green
    }

    Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName "No Restrictions"

}

#assign the 2 testing numbers to the accounts

Write-Host "Assigning $($TestNumbers[0]) to $($provisioningAccounts[0])"
Set-CsPhoneNumberAssignment -Identity $provisioningAccounts[0] -PhoneNumber $TestNumbers[0] -PhoneNumberType DirectRouting

Write-Host "Assign $(TestNumbers[0]) to another account via the Teams Admin Portal"

#endregion