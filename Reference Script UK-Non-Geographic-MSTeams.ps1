<#
.SYNOPSIS
A script to automatically create custom normalization rules for a Microsoft Teams Enterprise Voice deployment.

.DESCRIPTION
Automates the creation of Microsoft Teams Enterprise Voice dialplans/voice routes/policies etc. for various countries.

This script generated for Non-Geographic, United Kingdom.

.PARAMETER OverrideAdminDomain
OPTIONAL The FQDN your Office365 tenant. Use if your admin account is not in the same domain as your tenant (ie. doesn't use a @tenantname.onmicrosoft.com address)

.PARAMETER PSTNGateway
OPTIONAL The FQDN of a PSTN gateway to apply the script to.
If a value is not provided and multiple PSTN gateways exist, script will ask during execution.

.PARAMETER DPOnly
OPTIONAL. Only create dial plan
This option is useful for when all required routes/PSTN usages already exist, and you require separate dialplans for different groups.
Command line only option.

.EXAMPLE
.\UK-Non-Geographic-MSTeams.ps1
Runs the script in interactive mode. Script will prompt user for information when required.

.LINK
https://ucken.blogspot.com/2012/01/complete-guide-to-lync-optimizer.html

.NOTES
 -Works on Microsoft Teams environments with Enterprise Voice

To import the rules into Microsoft Teams, please save this file as a .ps1 (Powershell script). 
Run the program from Microsoft Teams Powershell by typing .\UK-Non-Geographic-MSTeams.ps1. 
#>

# The below settings are for applying command line options for unattended script application
param (
	# Input the PSTN Gateway name. Only necessary if multiple PSTN gateways are assigned to a mediation pool
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[string] $PSTNGateway,
	# Create only a dial plan. No routes/PSTN usages etc.
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[switch] $DPOnly,

	# Input the OverrideAdminDomain. Use if you normally have to enter your onmicrosoft.com domain name when signing onto O365
	[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
	[ValidateNotNullOrEmpty()]
	[string] $OverrideAdminDomain
)

# $ErrorActionPreference can be set to SilentlyContinue, Continue, Stop, or Inquire for troubleshooting purposes
$Error.Clear()
$ErrorActionPreference = 'SilentlyContinue'

If ((Get-PSSession | Where-Object -FilterScript {$_.ComputerName -like '*.online.lync.com' -Or $_.ComputerName -like '*.teams.microsoft.com'}).State -eq 'Opened') {
	Write-Host 'Using existing session credentials'}
Else {
	#$O365Session = New-CsOnlineSession -OverrideAdminDomain $OverrideAdminDomain
	#Import-PSSession $O365Session -AllowClobber
	Connect-MicrosoftTeams
}

# Prompt user to create either tenant-global or tenant-user dialplans.
Write-Host
$Global = New-Object System.Management.Automation.Host.ChoiceDescription '&Global','Create tenant global dial plan.'
$User = New-Object System.Management.Automation.Host.ChoiceDescription '&User','Create tenant user-level dial plan'
$Skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip','Skip dial plan creation'
$choices = [System.Management.Automation.Host.ChoiceDescription[]]($Global,$User,$Skip)
$message = 'Create global or user-level dial plan?'
$DialPlanSelect = $Host.UI.PromptForChoice($caption,$message,$choices,1)
Write-Host
If ($DialPlanSelect -eq 1) {
	Write-Host "Creating tenant user dial plan"
	$DPParent = "UK-Non-Geographic"
	New-CsTenantDialPlan $DPParent -Description "Normalization rules for Non-Geographic, United Kingdom"
}
ElseIf ($DialPlanSelect -eq 0) {
	Write-Host 'Using tenant global dial plan'
	$DPParent = "Global"
}

$ExtDPParent = $DPParent

If ($DialPlanSelect -ne 2) {
	Write-Host "Creating normalization rules"
	$NR = @()
	$NR += New-CsVoiceNormalizationRule -Name "UK-Non-Geographic-Local" -Parent $DPParent -Pattern '^(([2-8]\d\d|9[0-8]\d|99[0-8])\d{1,5})$' -Translation '+443$1' -InMemory -Description "Local number normalization for Non-Geographic, United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-TollFree' -Parent $DPParent -Pattern '^0((80(0\d{6,7}|8\d{7}|01111)|500\d{6}))\d*$' -Translation '+44$1' -InMemory -Description "TollFree number normalization for United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-Premium' -Parent $DPParent -Pattern '^0((9[018]\d|87[123]|70\d)\d{7})$' -Translation '+44$1' -InMemory -Description "Premium number normalization for United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-Mobile' -Parent $DPParent -Pattern '^0((7([1-57-9]\d{8}|624\d{6})))$' -Translation '+44$1' -InMemory -Description "Mobile number normalization for United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-National' -Parent $DPParent -Pattern '^0((1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8}|8((4[2-5]|70)\d{7}|45464\d)))\d*(\D+\d+)?$' -Translation '+44$1' -InMemory -Description "National number normalization for United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-Service' -Parent $DPParent -Pattern '^(1(47\d|70\d|800\d|1[68]\d{3}|\d\d)|999|[\*\#][\*\#\d]*\#)$' -Translation '$1' -InMemory -Description "Service number normalization for United Kingdom"
	$NR += New-CsVoiceNormalizationRule -Name 'UK-International' -Parent $DPParent -Pattern '^(?:\+|00)(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{5,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory -Description "International number normalization for United Kingdom"

	Set-CsTenantDialPlan -Identity $DPParent -NormalizationRules @{add=$NR} -Description "Policy updated to include UK normalization rules."
}

# If DPOnly option selected (meaning to only create dialplan), quit the program here.
If ($DPOnly) {
	Write-Host 'Finished!'
	Exit
}

# Check for existence of PSTN gateways and prompt to add PSTN usages/routes
If (($PSTNGateway -eq $NULL) -or ($PSTNGateway -eq '')) {
	$PSTNGW = Get-CsOnlinePSTNGateway
	If (($PSTNGW.Identity -eq $NULL) -and ($PSTNGW.Count -eq 0)) {
		Write-Host
		Write-Host 'No PSTN gateway found. If you want to configure Direct Routing, you have to define at least one PSTN gateway Using the New-CsOnlinePSTNGateway command.' -ForegroundColor Yellow
		Exit
	}

	If ($PSTNGW.Count -gt 1) {
		$PSTNGWList = @()
		Write-Host
		Write-Host "ID    PSTN Gateway"
		Write-Host "==    ============"
		For ($i=0; $i -lt $PSTNGW.Count; $i++) {
			$a = $i + 1
			Write-Host ($a, $PSTNGW[$i].Identity) -Separator "     "
		}

		$Range = '(1-' + $PSTNGW.Count + ')'
		Write-Host
		$Select = Read-Host "Select a primary PSTN gateway to apply routes" $Range

		If (($Select.ToInt32($NULL) -gt $PSTNGW.Count) -or ($Select -lt 1)) {
			Write-Host 'Invalid selection' -ForegroundColor Red
			Exit
		}
		Else {
			$PSTNGWList += $PSTNGW[$Select-1]
		}

		$Select = Read-Host "OPTIONAL - Select a secondary PSTN gateway to apply routes (or 0 to skip)" $Range

		If (($Select.ToInt32($NULL) -gt $PSTNGW.Count) -or ($Select -lt 0)) {
			Write-Host 'Invalid selection' -ForegroundColor Red
			Exit
		}
		ElseIf ($Select -gt 0) {
			$PSTNGWList += $PSTNGW[$Select-1]
		}
	}
	Else { # There is only one PSTN gateway
		$PSTNGWList = Get-CSOnlinePSTNGateway
	}
}
Else {
	$PSTNGWInputList = $PSTNGateway.Split(',').Trim()
	$PSTNGWList = @()
	ForEach ($PSTNGWInput in $PSTNGWInputList) {
		$PSTNGW = Get-CSOnlinePSTNGateway $PSTNGWInput -ErrorAction SilentlyContinue
		If (($PSTNGW.Identity -eq $NULL)) {
			Write-Host
			Write-Host "Could not find $PSTNGWInput in the tenant. Verify the name using Get-CsOnlinePSTNGateway." -ForegroundColor Yellow
			Write-Host 'Will attempt to use the PSTN gateway as entered, in case tenant is using Microsoft SuperTrunks.' -ForegroundColor Yellow
			$PSTNGWList += New-CsOnlinePSTNGateway -Identity $PSTNGWInput -SipSignalingPort 5061 -InMemory -ErrorAction SilentlyContinue
		}
		Else {
			$PSTNGWList += $PSTNGW
		}
	}
}

$UK_LocalList = ,'UK-Non-Geographic-Local'
$UK_MobileList = ,'UK-Non-Geographic-Mobile'
$UK_PremiumList = ,'UK-Non-Geographic-Premium'
$UK_NationalList = ,'UK-Non-Geographic-National'
$UK_InternationalList = ,'UK-Non-Geographic-International'

If ((Get-CsOnlineVoiceRoutingPolicy | Where-Object {$_.Identity -notlike 'UK-Non-Geographic-*'}).Count -gt 0) {
	If ($PSCmdlet.MyInvocation.BoundParameters["LeastCostRouting"].IsPresent -eq $NULL) {
		Write-Host
		$yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes','Will configure the voice policies for least-cost/failover routing'
		$no = New-Object System.Management.Automation.Host.ChoiceDescription '&No','Will not configure the voice policies for least-cost/failover routing'
		$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
		$message = 'Configure voice policies for least-cost/failover routing?'
		$FailoverRouting = $Host.UI.PromptForChoice($caption,$message,$choices,1)
		Write-Host
	}
	Else {
		If ($LeastCostRouting -eq $TRUE) {
			$FailoverRouting = 0
		}
		else {
			$FailoverRouting = 1
		}
	}

	If($FailoverRouting -eq 0) {
		Write-Host "Calculating least-cost/failover routing tables"

		# If user entered LCRSite list, split the string into an array
		If (($LCRSites -ne '') -and ($LCRSites -ne $NULL)) {$LCRList = $LCRSites -split ','}

		# Create Powershell lists with the words Local, Mobile, National, International in all supported languages
		$LocalLang = 'Lokal','Teghakan','Lokální','Dangdì','Lokaal','Local','Kohalik','Paikallinen','Locales','Adgilobrivi','Topiká','Lokalno','Helyi','Locali','Rokaru','Jibang-ui','Vietejs','Vietinis','Lokalen','Lokalny','Mestnyy','Miestne','Lokalni','Lokalt','Yerel','Mistsevyy','Tempatan'
		$MobileLang = 'ILëvizshëm','Ssharzhakan','Mobilní','Yídòng','Mobiel','Mobile','Mobiilne','Käsipuhelin','Mobiluri','Kinitó','Mobitel','Ponsel','Cellulari','Geitaidenwa','Mobail','Mobils','Mobilus','Mobilni','Komórkowy','Celular','Mobilnyy','Mobilné','Móvil','Medunarodni','Bimbit'
		$NationalLang = 'Kombëtar','Azgayin','Národní','Guó','National','Nationaal','Riiklik','Valtakunnallinen','Nationale','Erovnuli','Ethnikí','Nacionalni','Nemzeti','Nasional','Nazionali','Nashonaru','Guggaui','Valsts','Nacionalinis','Nacionalna','Nasjonal','Krajowy','Natsionalnyy','Celoštátna','Državni','Nationellt','Ulusal','Negara'
		$InternationalLang = 'Ndërkombëtar','Mijazgayin','Mezinárodní','GuójìDe','International','Internationaal','Rahvusvaheline','Kansainvälinen','Internationales','SaertAshoriso','Diethní','Internacionala','Nemzetközi','Internasional','Internazionali','Intanashonaru','Gugje','Starptautisks','Tarptautinis','Meg´unaroden','Internasjonal','Miedzynarodowy','Internationala','Mezhdunarodnyy','Medzinárodné','Internationellt','Uluslararasi','Mizhnarodnyy','Antarabangsa'

		If (($LCRSites -ne '') -and ($LCRSites -ne $NULL)) {
			# Create a list of all local, national and international PSTN usages for least-cost routing
			$UK_LocalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like "UK-*-Local" -and $_ -ne "UK-Non-Geographic-Local" -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
			$Int_LocalList = (Get-CsOnlinePstnUsage).Usage | Where {$LocalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*' -and $_ -ne 'Local' -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}

			$UK_MobileList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-Mobile' -and $_ -ne "UK-Non-Geographic-Mobile" -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
			$Int_MobileList = (Get-CsOnlinePstnUsage).Usage | Where {$MobileLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*' -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}

			$UK_PremiumList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-Premium' -and $_ -ne "UK-Non-Geographic-Premium" -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
			$UK_NationalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -Like 'UK-*-National' -and $_ -ne "UK-Non-Geographic-National" -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
			$Int_NationalList = (Get-CsOnlinePstnUsage).Usage | Where {$NationalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*' -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}

			$UK_InternationalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-International' -and $_ -ne "UK-Non-Geographic-International" -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
			$Int_InternationalList = (Get-CsOnlinePstnUsage).Usage | Where {$InternationalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*' -and $LCRList -cmatch ($_.SubString(0,$_.Length - ($_.Length-$_.LastIndexOf('-'))))}
		}
		Else {
			# Create a list of all local, national and international PSTN usages for least-cost routing
			$UK_LocalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like "UK-*-Local" -and $_ -ne "UK-Non-Geographic-Local"}
			$Int_LocalList = (Get-CsOnlinePstnUsage).Usage | Where {$LocalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*' -and $_ -ne 'Local'}
			$UK_MobileList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-Mobile' -and $_ -ne "UK-Non-Geographic-Mobile"}
			$Int_MobileList = (Get-CsOnlinePstnUsage).Usage | Where {$MobileLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*'}
			$UK_PremiumList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-Premium' -and $_ -ne "UK-Non-Geographic-Premium"}
			$UK_NationalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -like 'UK-*-National' -and $_ -ne "UK-Non-Geographic-National"}
			$Int_NationalList = (Get-CsOnlinePstnUsage).Usage | Where {$NationalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*'}
			$UK_InternationalList += (Get-CsOnlinePstnUsage).Usage | Where-Object {$_ -Like 'UK-*-International' -and $_ -ne "UK-Non-Geographic-International"}
			$Int_InternationalList = (Get-CsOnlinePstnUsage).Usage | Where {$InternationalLang -cmatch ($_.SubString($_.Length - ($_.Length-$_.LastIndexOf('-')-1),($_.Length-$_.LastIndexOf('-')-1)))+'$' -and $_ -notlike 'UK-*'}
		}
	}
}

Write-Host 'Creating voice policies'
New-CsOnlineVoiceRoutingPolicy "UK-Non-Geographic-Local" -Description "Allows local calls from Non-Geographic, United Kingdom" -WarningAction:SilentlyContinue | Out-Null

# Only create National/International policies if user did not select LocalOnly command line option
If (!$LocalOnly) {
	New-CsOnlineVoiceRoutingPolicy "UK-Non-Geographic-National" -Description "Allows local-national calls from Non-Geographic, United Kingdom" -WarningAction:SilentlyContinue | Out-Null
	New-CsOnlineVoiceRoutingPolicy "UK-Non-Geographic-International" -Description "Allows local-national-international calls from Non-Geographic, United Kingdom" -WarningAction:SilentlyContinue | Out-Null
}

Write-Host 'Creating PSTN usages'
Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Local"} -WarningAction:SilentlyContinue | Out-Null
Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Service"} -WarningAction:SilentlyContinue | Out-Null

# Only create following PSTN usages if user did not select LocalOnly command line option
If (!$LocalOnly) {
	Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-National"} -WarningAction:SilentlyContinue | Out-Null
	Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Mobile"} -WarningAction:SilentlyContinue | Out-Null
	Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-Premium"} -WarningAction:SilentlyContinue | Out-Null
	Set-CsOnlinePSTNUsage -Identity global -Usage @{Add="UK-Non-Geographic-International"} -WarningAction:SilentlyContinue | Out-Null
}

Write-Host 'Assigning PSTN usages to voice policies'
# It seems to take a while for PSTN usages to become available for usage, so if we get an error, wait a minute and try again.
$Iteration = 0
Do {
	Try {
		$PSTNUsageSetError = $False
		Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$UK_NationalList} -ErrorAction Stop
	}
	Catch {
		$PSTNUsageSetError = $True
		$Iteration++
		$Time = 60
		ForEach($i in (1..$Time)) {
			$Percentage = $i / $Time
			$Remaining = New-TimeSpan -Seconds ($Time - $i)
			$Message = "Round $Iteration`: PSTN usages not ready. Waiting for {1} before trying again. This may 10 minutes or longer." -f $Percentage, $Remaining
			Write-Progress -Activity $Message -PercentComplete ($Percentage * 100) -Status 'Waiting...'
			Start-Sleep 1
		}
	}
} While ($PSTNUsageSetError)

Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-Local" -OnlinePstnUsages @{Add=$UK_LocalList} | Out-Null
If ($Int_LocalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-Local" -OnlinePstnUsages @{Add=$Int_LocalList} | Out-Null}
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-Local" -OnlinePstnUsages @{Add="UK-Non-Geographic-Service"} | Out-Null
If ($UK_LocalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Replace=$UK_LocalList} | Out-Null}
Else {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages $NULL -WarningAction:SilentlyContinue | Out-Null}
If ($UK_MobileList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$UK_MobileList} | Out-Null}
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$UK_NationalList} | Out-Null
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add="UK-Non-Geographic-Service"} | Out-Null
If ($Int_LocalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$Int_LocalList} | Out-Null}
If ($Int_MobileList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$Int_MobileList} | Out-Null}
If ($Int_NationalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-National" -OnlinePstnUsages @{Add=$Int_NationalList} | Out-Null}
If ($UK_LocalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Replace=$UK_LocalList} | Out-Null}
Else {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages $NULL -WarningAction:SilentlyContinue | Out-Null}
If ($UK_MobileList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$UK_MobileList} | Out-Null}
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$UK_NationalList} | Out-Null
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$UK_PremiumList} | Out-Null
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add="UK-Non-Geographic-Service"} | Out-Null
If ($Int_LocalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$Int_LocalList} | Out-Null}
If ($Int_MobileList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$Int_MobileList} | Out-Null}
If ($Int_NationalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$Int_NationalList} | Out-Null}
Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$UK_InternationalList} | Out-Null
If ($Int_InternationalList -ne $NULL) {Set-CsOnlineVoiceRoutingPolicy -Identity "UK-Non-Geographic-International" -OnlinePstnUsages @{Add=$Int_InternationalList} | Out-Null}

# Prompt user if they want to apply PSTN usages to the Global voice policy.
Write-Host
Write-Host 'If desired, a set of dialing permissions and routes can be applied globally to all users.'
$Local = New-Object System.Management.Automation.Host.ChoiceDescription '&Local','Allow Local dialing via the Global voice policy'
$National = New-Object System.Management.Automation.Host.ChoiceDescription '&National','Allow National dialing via the Global voice policy'
$International = New-Object System.Management.Automation.Host.ChoiceDescription '&International','Allow International dialing via the Global voice policy'
$Skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip','No changes to the Global voice policy'
$choices = [System.Management.Automation.Host.ChoiceDescription[]]($Local,$National,$International,$Skip)
$message = 'Assign PSTN usages to the Global Voice Policy?'
$VPSelect = $Host.UI.PromptForChoice($caption,$message,$choices,3)
Write-Host

# Configure Global voice policy as per user selection above
Switch ($VPSelect) {
	0 {Set-CsOnlineVoiceRoutingPolicy -Identity Global -OnlinePstnUsages @{Replace=(Get-CsOnlineVoiceRoutingPolicy UK-Non-Geographic-Local).OnlinePstnUsages} -Description "Allows local calls from Non-Geographic, United Kingdom"}
	1 {Set-CsOnlineVoiceRoutingPolicy -Identity Global -OnlinePstnUsages @{Replace=(Get-CsOnlineVoiceRoutingPolicy UK-Non-Geographic-National).OnlinePstnUsages} -Description "Allows local-national calls from Non-Geographic, United Kingdom"}
	2 {Set-CsOnlineVoiceRoutingPolicy -Identity Global -OnlinePstnUsages @{Replace=(Get-CsOnlineVoiceRoutingPolicy UK-Non-Geographic-International).OnlinePstnUsages} -Description "Allows local-national-international calls from Non-Geographic, United Kingdom"}
}


Write-Host "Creating voice routes"
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Local" -Priority 0 -OnlinePstnUsages "UK-Non-Geographic-Local" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+440?3(([2-8]\d\d|9[0-8]\d|99[0-8])\d{1,5})' -Description "Local routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Mobile" -Priority 2 -OnlinePstnUsages "UK-Non-Geographic-Mobile" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+44(7([1-57-9]\d{8}|624\d{6}))$' -Description "Mobile routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-TollFree" -Priority 3 -OnlinePstnUsages "UK-Non-Geographic-Local" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+44(80(0\d{6,7}|8\d{7}|01111)|500\d{6})$' -Description "TollFree routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Premium" -Priority 4 -OnlinePstnUsages "UK-Non-Geographic-Premium" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+44(9[018]\d|87[123]|70\d)\d{7}$' -Description "Premium routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-National" -Priority 5 -OnlinePstnUsages "UK-Non-Geographic-National" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+440?(1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8}|8((4[2-5]|70)\d{7}|45464\d))' -Description "National routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-International" -Priority 7 -OnlinePstnUsages "UK-Non-Geographic-International" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(44))([2-9]\d{6,14})))' -Description "International routing for Non-Geographic, United Kingdom" | Out-Null
New-CsOnlineVoiceRoute -Name "UK-Non-Geographic-Service" -Priority 6 -OnlinePstnUsages "UK-Non-Geographic-Service" -OnlinePstnGatewayList $PSTNGWList.Identity -NumberPattern '^\+?(1(47\d|70\d|800\d|1[68]\d{3}|\d\d)|999|[\*\#][\*\#\d]*\#)$' -Description "Service routing for Non-Geographic, United Kingdom" | Out-Null

Write-Host 'Creating outbound translation rules'
$OutboundTeamsNumberTranslations = New-Object 'System.Collections.Generic.List[string]'
$OutboundPSTNNumberTranslations = New-Object 'System.Collections.Generic.List[string]'
New-CsTeamsTranslationRule -Identity "UK-Non-Geographic-AllCalls" -Pattern '^\+(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{5,14})(;ext=\d+)?$' -Translation '+$1$2' -Description "" | Out-Null
$OutboundTeamsNumberTranslations.Add("UK-Non-Geographic-AllCalls")

Write-Host 'Adding translation rules to PSTN gateways'
ForEach ($PSTNGW in $PSTNGWList) {
	Set-CsOnlinePSTNGateway -Identity $PSTNGW.Identity -OutboundTeamsNumberTranslationRules $OutboundTeamsNumberTranslations -OutboundPstnNumberTranslationRules $OutboundPSTNNumberTranslations -ErrorAction SilentlyContinue
}

Write-Host 'Finished!'

# SIG # Begin signature block
# MIINEQYJKoZIhvcNAQcCoIINAjCCDP4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUV6SNYtjj2HCs2bP7g0JfEqAp
# 3S6gggpTMIIFGzCCBAOgAwIBAgIQCVcmswfJ1koJSkb+PKEcYzANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDIxMzAwMDAwMFoXDTIzMDIx
# NzEyMDAwMFowWDELMAkGA1UEBhMCQ0ExEDAOBgNVBAgTB09udGFyaW8xDzANBgNV
# BAcTBkd1ZWxwaDESMBAGA1UEChMJS2VuIExhc2tvMRIwEAYDVQQDEwlLZW4gTGFz
# a28wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCsCGnJWjBqNX+R8Pyv
# 24IX7EnQkYm8i/VOd5dpxUMKOdnq+oEi2fr+tkWHtSgggbjTdcXP4l6fBxBzuGB2
# q12eLaa136Um4KAmYRnuqJ2IXfdEyW8/Zib7FVzUV41dwRBVH/VF+QZOHxwcL0MJ
# 5OwiRSLiMWYqWk7c+8UIFpDe17Pjevy8g2o0RcTAhyDeEZ1FPAIFk/nkirB5psMz
# mC5TfCKkuxQWOg3/F78KnvBxuVl7q9QcS2BeJXrospvQ130qRMOjrcO6suuRjtrT
# iuMt3CjKtStnqKAY/2yPV1Gvlg4itoO1quANvoNgYB66B3zQZMBGicdwnq0nkG7B
# vPENAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32ZXUO
# WDAdBgNVHQ4EFgQUHAoKnWsI9RGj62kGJANJvR36UXYwDgYDVR0PAQH/BAQDAgeA
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6Ly9j
# cmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Axhi9o
# dHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDBM
# BgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3
# dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2MCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUHMAKG
# Qmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1cmVk
# SURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUA
# A4IBAQAVlWAZaJ+IYUINcKhTdojQ3ViHdfcuzgTMECn1bCbr4DxH/SDmNr1n3Nm9
# ZIodfPE5fjFuqaYqQvfzFutqBTiX5hStT284WpGjrPwSuiJDRoI3iBTjyy5OABpi
# kpgdncJeyNvEO2ermcBrVw4t4AUZKfYsyjxfXaX8INcvHdNGhTiN5x4SjGSXxSvx
# hr7F9aLeE0mG+5yDlr5HbfPbyqLWdvLP4UcQ9WrJOmN0wa7qanrErr3ZeuDZQebL
# zEesJy1VCY2bqTEI8/fyTqnlLjut7i9dvp84zKomX30lqy9R81WUas9XruMLfgVR
# 3BVuBoyVtdx4AmgVzHoznDWs/vh/MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1
# b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtE
# aWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgx
# MDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBT
# SEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLX
# cep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSR
# I5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXi
# TWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5
# Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8
# vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYD
# VR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4
# oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJv
# b3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCow
# KAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZI
# AYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaA
# FEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPz
# ItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRu
# pY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKN
# JK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmif
# z0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN
# 3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKy
# ZqHnGKSaZFHvMYICKDCCAiQCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UE
# AxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQCVcm
# swfJ1koJSkb+PKEcYzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUMHXs5WotV7eFfXJjZ2fyymr8
# EAAwDQYJKoZIhvcNAQEBBQAEggEAmLF0FQ1LqgOD8RqXZUIkelaiZB15h2IaeVbD
# Sjpw2TlDi9NOnbhHegRPSBebvVKnmHtod/W1ryyUwtnACCM77Urn8Ij5FLI6PIkE
# J5Wv2Bi8O9tWSNhvQuTFbf4bEl2PFq6v/wURyYpplKhrS2JCD0FvilHYtYIi4obY
# dVTS6RzGKVIXY9Q7XckRblZWp8jcBn8U7Clc0gj6UPywaTxy7F+8cyN5D8wV0k1n
# o1MtE6BMMAJSEO++isMewzQAQzXV3RYwGD8MyFA8XqxJhSPsImlXdymAlNLfAdWR
# gehQF27JfKLBWbkUCCJM98AEmUw8M+5jCEz4f3jEVXPsAzDjkA==
# SIG # End signature block
