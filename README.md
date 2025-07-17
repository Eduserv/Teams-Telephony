# Jisc Cloud's Teams Phone System Scripts

> Before running the SBC configuration script ensure the SBC sub domain has been verified in your tenant as per the SBC onboarding document.

### Start by setting up some test users and [initial SBC configuration]((./TeamsSBCConfigurationQuickScript.ps1))
You can run [TeamsSBCConfigurationQuickScript.ps1](./TeamsSBCConfigurationQuickScript.ps1) for this purpose.
>ℹ️ Jisc cloud will likely not use this

## [Main setup](./TeamsPhoneSystemMain.ps1)
Download and run [TeamsPhoneSystemMain.ps1](./TeamsPhoneSystemMain.ps1). This will guide you through the Jisc SBC and base Teams Phone setup.
>ℹ️ Jisc cloud will likely get you run in Holiday mode first and then full after that. You will need your SBC FQDN and have already setup a user on this domain.

### [Mass number assignment](./MassNumberAssignment.ps1)
1. Create a CSV file with 2 columns:
    1. User Principal Name (UPN)
    2. Telephone Number (full number with extension if needed)
2. Download and run [MassNumberAssignment.ps1](./MassNumberAssignment.ps1)
3. Select the CSV above and then select the columns as instructed
> ⚠️ This will overwrite any pre-assigned numbers