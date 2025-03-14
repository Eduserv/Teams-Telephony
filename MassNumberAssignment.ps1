<#
Mass Assigns Numbers

Modified by Nick

#>
Start-Transcript
Write-Host "Nick's Teams Phone mass number assignment script - please have your CSV file ready with UPN + Phone Number"

Connect-MicrosoftTeams

Read-Host "Please complete the LocationID column in the CSV file, press enter to continue"

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

    $i = 0
    foreach ($a in $csv) {
        $i++;
        if ($a."$upn" -ne "") {
            Write-Progress -Activity "$i / $($csv.Count) Assigning $($a."$upn") $($a."$phone")" -PercentComplete ($i/$csv.Count * 100)
            try {
                Set-CsPhoneNumberAssignment -Identity $a."$upn" -PhoneNumber $a."$phone" -PhoneNumberType DirectRouting -ErrorAction Stop
            } catch {
                Write-Error "Error setting phone number on $($a."$upn")"
                Write-Error $_
            }
        }
    }

} else {
    throw "Mass Numbers CSV file open dialog cancelled"
}
Stop-Transcript
