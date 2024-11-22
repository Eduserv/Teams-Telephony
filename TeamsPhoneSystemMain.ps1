<#
Jisc standard teams telephony delivery currently includes the following elements. #

Modified by Nick

1. Holiday definition for uk bank holidays.
2. Emergency calling routing rules and locations.
    https://docs.microsoft.com/en-us/microsoftteams/what-are-emergency-locations-addresses-and-call-routing#emergency-call-enablement-for-direct-routing
    https://shawnharry.co.uk/2020/03/11/configuring-e911-with-microsoft-teams-direct-routing-audiocodes/
3. Normalization rules add to default dial plans for UK.
4. Voice routing policy. - (Note that the customer will have to have full call support from the SIP provider to use all of the functionality).
    Standard Routing
    Premium Routing
    Premium Plus International Routing

#>

try { 
    Get-InstalledModule MicrosoftTeams -ErrorAction Stop
} catch { 
    Install-Module MicrosoftTeams -Scope CurrentUser
}

Connect-MicrosoftTeams

$question = $host.UI.PromptForChoice("What type?", "Select the type of script run this is", ([System.Management.Automation.Host.ChoiceDescription]"&Holidays only",  [System.Management.Automation.Host.ChoiceDescription]"&Full"), 0)

#region:1. Create bank holiday definition
#download from gov.uk
$hols = ConvertFrom-Json (Invoke-WebRequest -Uri "https://www.gov.uk/bank-holidays.json").Content

$options = @()
foreach ($region in ($hols | Get-Member -MemberType NoteProperty)) {
    $options += [System.Management.Automation.Host.ChoiceDescription] "&$($region.Name)"
}

$opt = $host.UI.PromptForChoice("Choose Region", "Select the Region to load bank holiday dates for", $options, 0)

$DateRanges = @()
foreach ($e in $hols."$(($hols | Get-Member -MemberType NoteProperty)[$opt].Name)".events | where-object { (Get-Date -Date $_.date) -gt (Get-Date) }) {
    $d = Get-Date -Date $e.date
    $DateRanges += New-CsOnlineDateTimeRange -Start $d.ToString("d/M/yyyy") -End ($d.addDays(1)).ToString("d/M/yyyy")
}

#Add Additional dates to the Date Range
#$DateRanges += New-CsOnlineDateTimeRange -Start "dd/mm/yyyy" -End "dd+1/mm/yyyy"
$HolidayDef = Get-CsOnlineSchedule | ? Name -eq "UK Bank Holidays"

if ($HolidayDef -eq $null) {
    Write-Host "Creating Schedule"
    $HolidayDef = New-CsOnlineSchedule -Name "UK Bank Holidays" -FixedSchedule -DateTimeRanges $DateRanges
} else {
    Write-Host "Updating Schedule"
    $HolidayDef.FixedSchedule.DateTimeRanges += $DateRanges
    Set-CsOnlineSchedule -Instance $HolidayDef
}

#Following section will only run if a full run is 
if ($question -gt 0) {
#endregion

#region:2. Emergency Policies

#emergency policies are made up of 2 components
#1. Calling Policies
#2. Calling routing policies
#Both of these have existing 'Global (Org-wide default)' policies that we will configure.
#3. Emergency locations - to be defined 


#1. Calling Policies
#https://docs.microsoft.com/en-us/microsoftteams/manage-emergency-calling-policies

$EmergencyNotifiactionGroup = Read-Host "What is the UPN of the group to notify in an emergency? (Nothing will not set this)"
if ($EmergencyNotifiactionGroup -eq "") {
    $EmergencyNotifiactionGroup = $null
}
$Disclaimer = "Remember to configure your emergency location."
$Description = "This default global policy has been updated to enable External location lookup mode (users to configure their emergency address outside the corporate network). Notification to: $EmergencyNotifiactionGroup"

Set-CsTeamsEmergencyCallingPolicy -Identity global -Description $Description -NotificationGroup $EmergencyNotifiactionGroup -NotificationMode NotificationOnly -ExternalLocationLookupMode Enabled -EnhancedEmergencyServiceDisclaimer $Disclaimer

#2. Calling routing policies
#https://docs.microsoft.com/en-us/microsoftteams/manage-emergency-call-routing-policies

$EmergencyNumber = "999"
$Description = "This default global policy has been updated to include US (911) and European (112) emergency dial masks."
$EmergencyNumbers = @()
$pstnusage = "UnrestrictedPstnUsage"

$EmergencyNumbers += New-CsTeamsEmergencyNumber -EmergencyDialMask "911;112" -EmergencyDialString $EmergencyNumber -OnlinePSTNUsage $pstnusage

Set-CsTeamsEmergencyCallRoutingPolicy -Identity Global -AllowEnhancedEmergencyServices $true -Description $Description -EmergencyNumbers $EmergencyNumbers

#3. Emergency locations - https://shawnharry.co.uk/2020/03/11/configuring-e911-with-microsoft-teams-direct-routing-audiocodes/
#endregion

#region:3. UK normalization rules - add to standard global dial plan

    Write-Host "Creating normalization rules"
    $DPParent = "Global"
    $NR = @()

    $translation = Read-Host -Prompt "Enter the local area code e.g. +441329, entering nothing here will skip this step"
    if ($translation -ne "") {
        #Get a local routing for the local area to normalize the dailing of the local number
        $NR += New-CsVoiceNormalizationRule -Name "UK-Local" -Parent $DPParent -Pattern '^(([2-8]\d\d|9[0-8]\d|99[0-8])\d{3})$' -Translation ($translation + '$1') -InMemory -Description "Local number normalization"
    }
    $NR += New-CsVoiceNormalizationRule -Name "UK-Non-Geographic-Local" -Parent $DPParent -Pattern '^(([2-8]\d\d|9[0-8]\d|99[0-8])\d{1,5})$' -Translation '+443$1' -InMemory -Description "Local number normalization for Non-Geographic, United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-TollFree' -Parent $DPParent -Pattern '^0((80(0\d{6,7}|8\d{7}|01111)|500\d{6}))\d*$' -Translation '+44$1' -InMemory -Description "TollFree number normalization for United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-Premium' -Parent $DPParent -Pattern '^0((9[018]\d|87[123]|70\d)\d{7})$' -Translation '+44$1' -InMemory -Description "Premium number normalization for United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-Mobile' -Parent $DPParent -Pattern '^0((7([1-57-9]\d{8}|624\d{6})))$' -Translation '+44$1' -InMemory -Description "Mobile number normalization for United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-National' -Parent $DPParent -Pattern '^0((1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8}|8((4[2-5]|70)\d{7}|45464\d)))\d*(\D+\d+)?$' -Translation '+44$1' -InMemory -Description "National number normalization for United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-Service' -Parent $DPParent -Pattern '^(1(47\d|70\d|800\d|1[68]\d{3}|\d\d)|999|[\*\#][\*\#\d]*\#)$' -Translation '$1' -InMemory -Description "Service number normalization for United Kingdom"
    $NR += New-CsVoiceNormalizationRule -Name 'UK-International' -Parent $DPParent -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{5,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory -Description "International number normalization for United Kingdom"


    #If you are in a multi site setup you may want to create multiple dial plans for each site with different local normalization rules, by default adjust the Global with the UK generics
    Set-CsTenantDialPlan -Identity Global -NormalizationRules @{add=$NR} -Description "Policy updated to include UK normalization rules."

#endregion

#Region: IPPhonePolicies
<#
The default "Global" policy is configured for UserSignIn and will allow user to signout
"CAP" policy is configured for common areas phones with directory access, and will hide user signout under the admin menu
"CAP-Secure" doesn't have directory access , and will hide user signout under the admin menu
#>
    Write-Host "Creating Common Area Phone Policies"
    New-CsTeamsIPPhonePolicy -Identity "CommonAreaPhone" -Description "Common Area Phone User Policy" -SignInMode CommonAreaPhoneSignIn -AllowHotDesking:$false
    New-CsTeamsIPPhonePolicy -Identity "CommonAreaPhone-Secure" -Description "Common Area Phone User Policy without Directory Search" -SignInMode CommonAreaPhoneSignIn -AllowHotDesking:$false -SearchOnCommonAreaPhoneMode 'Disabled' -AllowHomeScreen 'Disabled'
    New-CsTeamsIPPhonePolicy -Identity "UserPhone" -Description "IP Phone User Policy - Allow Hot Desking" -SignInMode usersignin -AllowHotDesking:$True -AllowBetterTogether 'Enabled'

#endregion

#region:4. Voice Routing

    #Create Voice Routing Policies - these are the policies that must be assigned to users

    Write-Host 'Creating voice policies'
    New-CsOnlineVoiceRoutingPolicy "Standard Routing" -Description "Allows UK calls excluding premium rate numbers" -WarningAction:SilentlyContinue | Out-Null
    New-CsOnlineVoiceRoutingPolicy "Premium Routing" -Description "Allows UK calls including premium rate numbers" -WarningAction:SilentlyContinue | Out-Null
    New-CsOnlineVoiceRoutingPolicy "Standard Routing plus International Routing" -Description "Allows Standard and internal calls excluding premium rate numbers" -WarningAction:SilentlyContinue | Out-Null
    New-CsOnlineVoiceRoutingPolicy "Premium plus International Routing" -Description "Allows Premium and internal calls" -WarningAction:SilentlyContinue | Out-Null

    #Create PSTN Usages

    Write-Host 'Creating PSTN usages'
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Local"} -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Service"} -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-National"} -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Mobile"} -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Premium"} -WarningAction:SilentlyContinue | Out-Null
    Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-International"} -WarningAction:SilentlyContinue | Out-Null

    # Assign PSTN usage to voice routing policy
    # It seems to take a while for PSTN usages to become available for usage, so if we get an error, wait a minute and try again.
    Write-Host 'Assigning PSTN usages to voice policies'

    $StandardRoutingUsages = "UK-Non-Geographic-Local","UK-Non-Geographic-Service","UK-Non-Geographic-National","UK-Non-Geographic-Mobile"
    $StandardPlusIntRoutingUsages = "UK-Non-Geographic-Local","UK-Non-Geographic-Service","UK-Non-Geographic-National","UK-Non-Geographic-Mobile","UK-Non-Geographic-International"
    $PremiumRoutingUsages = "UK-Non-Geographic-Local","UK-Non-Geographic-Service","UK-Non-Geographic-National","UK-Non-Geographic-Mobile","UK-Non-Geographic-Premium"
    $PremiumPlusRoutingUsages = "UK-Non-Geographic-Local","UK-Non-Geographic-Service","UK-Non-Geographic-National","UK-Non-Geographic-Mobile","UK-Non-Geographic-Premium","UK-Non-Geographic-International"

    Set-CsOnlineVoiceRoutingPolicy -Identity "Standard Routing" -OnlinePstnUsages @{Add=$StandardRoutingUsages}
    Set-CsOnlineVoiceRoutingPolicy -Identity Global -OnlinePstnUsages @{Add=$StandardRoutingUsages}
    Set-CsOnlineVoiceRoutingPolicy -Identity "Standard Routing plus International Routing" -OnlinePstnUsages @{Add=$StandardPlusIntRoutingUsages}
    Set-CsOnlineVoiceRoutingPolicy -Identity "Premium Routing" -OnlinePstnUsages @{Add=$PremiumRoutingUsages}
    Set-CsOnlineVoiceRoutingPolicy -Identity "Premium plus International Routing" -OnlinePstnUsages @{Add=$PremiumPlusRoutingUsages}



    $CustomerSBCAddress = Read-Host "Enter the customer SBC FQDN (example: jsl.sbc.jisc.ac.uk)"

    if ($CustomerSBCAddress -ne "") {
        Write-Host "Creating voice routes"
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Local" -Priority 1 -OnlinePstnUsages "UK-Non-Geographic-Local" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+440?3(([2-8]\d\d|9[0-8]\d|99[0-8])\d{1,5})' -Description "Local routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Mobile" -Priority 2 -OnlinePstnUsages "UK-Non-Geographic-Mobile" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+44(7([1-57-9]\d{8}|624\d{6}))$' -Description "Mobile routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-TollFree" -Priority 3 -OnlinePstnUsages "UK-Non-Geographic-Local" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+44(80(0\d{6,7}|8\d{7}|01111)|500\d{6})$' -Description "TollFree routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Premium" -Priority 4 -OnlinePstnUsages "UK-Non-Geographic-Premium" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+44(9[018]\d|87[123]|70\d)\d{7}$' -Description "Premium routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-National" -Priority 5 -OnlinePstnUsages "UK-Non-Geographic-National" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+440?(1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8}|8((4[2-5]|70)\d{7}|45464\d))' -Description "National routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-International" -Priority 7 -OnlinePstnUsages "UK-Non-Geographic-International" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(44))([2-9]\d{6,14})))' -Description "International routing for Non-Geographic, United Kingdom" | Out-Null
        New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Service" -Priority 6 -OnlinePstnUsages "UK-Non-Geographic-Service" -OnlinePstnGatewayList $CustomerSBCAddress -NumberPattern '^\+?(1(47\d|70\d|800\d|1[68]\d{3}|\d\d)|999|[\*\#][\*\#\d]*\#)$' -Description "Service routing for Non-Geographic, United Kingdom" | Out-Null
    }
#endregion

}
Write-Host "Jisc Telephony Script Completed"
