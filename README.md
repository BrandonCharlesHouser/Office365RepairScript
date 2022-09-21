# Office365RepairScript
A script that checks if Office 365 is installed as an x86/x64 application, runs a repair, and then reinstalls Teams.

I have it reinstall Teams as there is a bug in the organization I developed this in.
Intermittently up to a day after an Office 365 repair some users will lose Teams, a reinstall remedies this.
Reinstalling after every repair reduces the chances of this ocurring to 0%

I cannot trace the cause of this bug as I am not a System's Administrator for my organization, and I am given only 30 minutes maximum to aid each user.

---
Use the -help/-h parameters to display a help page when calling OfficeRepairAndTeamsCaller.bat from the cmd line.
---
OfficeRepairAndTeamsCaller.bat -help

OfficeRepairAndTeamsCaller.bat -h  

---
Use -Background to force close all Office 365 apps and run the repair in the background.
Once the repair completes a normal hands off Teams install occurs.
---
OfficeRepairAndTeamsCaller.bat -Background
