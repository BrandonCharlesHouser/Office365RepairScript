param (
    # If set to false ALL Office 365 apps will be force closed and repair will occurr in the background. User will not have any progress notifications until Teams begins reinstalling and the script completes.
    [switch] $Background,
    [switch] $help,
    [switch] $h
)

function DisplayHelp {
    # Get-Help OfficeRepair -Detailed
    write-host @"
NAME
OfficeRepair

DETAILS
Author: Brandon Houser
Date:   June 28, 2010
Github: https://github.com/BrandonCharlesHouser/


SYNOPSIS
Office Repair Script 1.0


DESCRIPTION
Automatically detects if Office 365 is installed as an x86 or x64 application. Then performs a repair and
reinstalls Teams when complete. (Reinstalls Teams to address bug where Teams will disappear after a repair.)


SYNTAX
OfficeRepair [-Background] [-help] [-h] [<CommonParameters>]


PARAMETERS
-Background [<SwitchParameter>]
    Force closes all Office 365 applications and performs a background repair.

-help / h [<SwitchParameter>]
    Opens help menu.

<CommonParameters>
    This cmdlet supports the common parameters: Verbose, Debug,
    ErrorAction, ErrorVariable, WarningAction, WarningVariable,
    OutBuffer, PipelineVariable, and OutVariable. For more information, see
    about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

-------------------------- EXAMPLE 1 --------------------------

C:\System32>OfficeRepairAndTeamsCaller.bat
Performs a standard repair equivalent to a repair started from control panel


-------------------------- EXAMPLE 2 --------------------------

C:\System32>OfficeRepairAndTeamsCaller.bat -Background
Force closes all Office 365 applications and performs a background repair.

"@
}

function Write-HostAndMsgBox($MessageText) {
    Start-Process -FilePath "CMD.exe" -ArgumentList '/c', 'msg', '*', $MessageText
    Throw ($MessageText)
}
if ($help.IsPresent -or $h.IsPresent) {
    write-host "Office Repair 1.0 Help Guide"
    write-host "Script by Brandon Houser: https://github.com/BrandonCharlesHouser/"

    DisplayHelp
}
else {

    # Check if running in elevated session
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $AmIAdmin = $($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))

    if (-not $AmIAdmin) {
        $AdminErrorText = "You must run script as Administrator."
        Write-HostAndMsgBox $AdminErrorText
    }

    #Set Variables
    $UserDownloads = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
    $TeamsInstallerLocation = $(join-path -Path $UserDownloads -ChildPath "Teams_windows.exe")
    $TeamsInstallerWebLocation = "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x409&culture=en-us&country=US"

    # This variable sets a time span, if the script detects that the Repair program ran for less than the set timespan then teh rest of the script is aborted
    $MinimumRunTime = New-TimeSpan -Minutes 5

    # Finds whether user has x86 or x64 version of Office365 installed
    $OfficeBitVersion = ($(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" | Select-Object Platform).platform)

    # Writes to terminal whether you have 32/64 bit Office and throws an exception/aborts script if an unexpected result appears
    if ("x86" -eq $OfficeBitVersion) {
        Write-Host "32-Bit Office365 installation found."
    }
    elseif ("x64" -eq $OfficeBitVersion) {
        Write-Host "64-Bit Office365 installation found."
        }
    else {Throw ("Office365 version not found. Manually run repair.")}

    # Builds the Platform parameter for OfficeClickToRun.exe
    $OfficePlatform = $("platform=$OfficeBitVersion")

    # Checks if Background Flag is set then builds the DisplayLevel and forceappshutdown parameters for OfficeClickToRun.exe
    if ($Background.IsPresent) {
        $OfficeDisplayLevel = @("DisplayLevel=False", "forceappshutdown=True")
    } Else {
        $OfficeDisplayLevel = @("DisplayLevel=True")
    }

    # Builds the argument list for OfficeClickToRun.exe
    $OfficeRepairArgList = @("scenario=Repair", "$OfficePlatform", "culture=en-us", "RepairType=FullRepair") + $OfficeDisplayLevel

    # Writes command to be run to the Terminal
    Write-Host ('Start-Process "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" -ArgumentList -NoNewWindow -Wait -PassThru ' + $OfficeRepairArgList)
    Set-Location -Path "C:\Program Files\Microsoft Office 15\ClientX64\"

    # Stores Process data in Object $process
    $process = Start-Process "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" -NoNewWindow -Wait -PassThru -ArgumentList $OfficeRepairArgList
    # Subtracts Process Start Time and Exit Time to get the run time as a TimeSpan Object  
    $processRunTime = ($process.StartTime - $process.ExitTime)




    # If Repair Program errors out or quits unexpectedly it will return an error code other than 0
    # Checks if Error Code is not 0 AND if Process Run Time equals Minimum Run Time
    if ($($process.ExitCode -eq 0) -and $($processRunTime -ge $MinimumRunTime)) {
        #Teams Download and Install - Repairs sometimes cause Teams to disappear sometimes even a day later, reinstalling prevents this$TeamsInstallerWebLocation = "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x409&culture=en-us&country=US"
        Invoke-WebRequest -UseBasicParsing -Uri $TeamsInstallerWebLocation -outfile $TeamsInstallerLocation
        Start-Process -FilePath $TeamsInstallerLocation -NoNewWindow -Wait -PassThru

        # Writes Message to terminal and creates a Message Box with the same text
        $ExitText = "Office repaired"
        Write-HostAndMsgBox $ExitText
    }
    Else {
        # Writes Message to terminal and creates a Message Box with the same text
        $ExitText = "Office Not Repaired"
        Write-Host $ExitText
        if (-not ($process.ExitCode) -eq 0) {
            $ExitCode = "Error Code: $($process.ExitCode)"
            Write-HostAndMsgBox $ExitCode
        }
        Write-HostAndMsgBox $ExitText
    }
}
