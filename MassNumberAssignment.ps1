<#
Mass Assigns Numbers

Modified by Nick

#>

Write-Host "This script will first get a list of location IDs, please use this in the csv file LocationID field"

Connect-MicrosoftTeams

Get-CsOnlineLisLocation | ft Description,LocationID

Read-Host "Please complete the LocationID column in the CSV file, press any key to continue"

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
$OpenFileDialog.title = "Select the mass numbers csv file"
if ($OpenFileDialog.ShowDialog() -eq "OK") {
    $csv = Import-CSV $OpenFileDialog.filename
    $options = @()
    foreach ($prop in $csv[0].psobject.properties) {
        $options += [System.Management.Automation.Host.ChoiceDescription] "&$($prop.Name)"
    }
    
    $upn = $options[($host.UI.PromptForChoice("UPN Column", "Select the column with the UPN/Email in", $options, 0))].Label -replace "&",""
    $phone = $options[($host.UI.PromptForChoice("Phone Number Column", "Select the column with the phone number/extension in", $options, 0))].Label -replace "&",""
    $location = $options[($host.UI.PromptForChoice("LocationID Column", "Select the column with the location id in", $options, 0))].Label -replace "&",""

    $i = 0
    foreach ($a in $csv) {
        $i++;
        Write-Progress -Activity "$i / $($csv.Count) Assigning $($a."$upn") $($a."$phone")" -PercentComplete ($i/$csv.Count * 100)
        Set-CsPhoneNumberAssignment -Identity $a."$upn" -PhoneNumber $a."$phone" -LocationId $a."$location" -PhoneNumberType DirectRouting
    }

} else {
    throw "Mass Numbers CSV file open dialog cancelled"
}
