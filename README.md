# JCS Teams Phone System Scripts

Before running the SBC configuration script ensure the SBC has been configured as per the SBC service document.

When provisioning the SBC system run [TeamsSBCConfigurationQuickScript](./TeamsSBCConfigurationQuickScript.ps1), this will create the user accounts and license them.

Once SBC testing has been completed run [TeamsPhoneSystemMain](./TeamsPhoneSystemMain.ps1).

You can run [TeamsPhoneSystemMain](./TeamsPhoneSystemMain.ps1) seperately and use the Holiday option to configure the holidays.