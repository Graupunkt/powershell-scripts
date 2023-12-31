#TASKBAR SETTINGS
$Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$Icon_Alignment = "TaskbarAI" 
$Key_Teams = "TaskbarMn"
$Key_Widgets = "TaskbarDa"
$Key_News = "TaskbarDn"
$Key_Taskview = "ShowTaskViewButton"
$Key_Search = "SearchboxTaskbarMode"
$KeyFormat = "DWord"
$Value = "0"

if(!(Test-Path $Path)){New-Item -Path $Path -Force}
Set-ItemProperty -Path $Path -Name $Icon_Alignment -Value $Value -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_Teams -Value $Value -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_Taskview -Value $Value -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_Widgets -Value $Value -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_News -Value $Value -Type $KeyFormat
#Set-ItemProperty -Path $Path -Name $Key_Search -Value $Value -Type $KeyFormat

#DARK MODE
$Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
$Key_1 = "ColorPrevalence"
$Key_2 = "EnableTransparency"
$Key_3 = "AppsUseLightTheme"
$Key_4 = "SystemUsesLightTheme"
$KeyFormat = "DWord"

if(!(Test-Path $Path)){New-Item -Path $Path -Force}
Set-ItemProperty -Path $Path -Name $Key_1 -Value 0 -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_2 -Value 1 -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_3 -Value 0 -Type $KeyFormat
Set-ItemProperty -Path $Path -Name $Key_4 -Value 0 -Type $KeyFormat

#EXPLORER
Set-Itemproperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'HideFileExt' -value 0

#RIGHT CLICK LEGACY CONEXT MENU
# REMOVE = reg delete "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" /f
reg.exe add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

#REMOVE FILE GROPUING
$Bags = 'HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags'
$DLID = '{885A186E-A440-4ADA-812B-DB871B942259}'
(Get-ChildItem $bags -recurse | ? PSChildName -eq $DLID ) | Remove-Item
gps explorer | spps

#SHOW DEFENDER ICON ON TASKBAR
# LIST OF APP NAMES = HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\
# HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.Defender.SecurityCenter /v Enabled /f
#REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.Defender.SecurityCenter" r /v Enabled /t REG_DWORD /d 1 /f
$Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.Defender.SecurityCenter"
$Key_1 = "Enabled"
$KeyFormat = "DWord"
if(!(Test-Path $Path)){New-Item -Path $Path -Force}
Set-ItemProperty -Path $Path -Name $Key_1 -Value 0 -Type $KeyFormat

#UPDATE EXPLORER / SYSTEM
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class User32 {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            UIntPtr wParam,
            IntPtr lParam,
            SendMessageTimeoutFlags fuFlags,
            uint uTimeout,
            out UIntPtr lpdwResult
        );

        [Flags]
        public enum SendMessageTimeoutFlags : uint {
            SMTO_NORMAL = 0x0000,
            SMTO_BLOCK = 0x0001,
            SMTO_ABORTIFHUNG = 0x0002,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
        }
    }
"@
function Send-WMSettingChange {
    param(
        [string]$area = "Environment",
        [string]$param = $null
    )

    $hwndBroadcast = [System.IntPtr]::Zero
    $WM_SETTINGCHANGE = 0x001A
    $SMTO_ABORTIFHUNG = 0x0002

    $result = [User32]::SendMessageTimeout($hwndBroadcast, $WM_SETTINGCHANGE, [UIntPtr]::Zero, [IntPtr]::Zero, [User32.SendMessageTimeoutFlags]::SMTO_ABORTIFHUNG, 5000, [ref]$null)

    if ($result -eq [UIntPtr]::Zero) {
        Write-Host "Failed to send WM_SETTINGCHANGE message."
    } else {
        Write-Host "WM_SETTINGCHANGE message sent successfully."
    }
}

# Example: Send WM_SETTINGCHANGE for the "Environment" area
Send-WMSettingChange -area "Environment"

# STUFF FOR DEFAULT USER
<# HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\FileZilla.Client.AppID /v Enabled /f
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize]
"ColorPrevalence"=dword:00000000
"EnableTransparency"=dword:00000001
"AppsUseLightTheme"=dword:00000000
"SystemUsesLightTheme"=dword:00000000

#darmode.cmd
reg load "hku\Default" "C:\Users\Default\NTUSER.DAT" 
reg import DarkMode\DarkMode.reg
reg unload "hku\Default

#>




<#

## THIS IS TO SET THE SETTINGS FOR ANY NEW USER IN THE DEFAULT SETTINGS
REG LOAD HKLM\Default C:\Users\Default\NTUSER.DAT
 
# Removes Task View from the Taskbar
New-itemproperty "HKLM:\Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value "0" -PropertyType Dword
 
# Removes Widgets from the Taskbar
New-itemproperty "HKLM:\Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Value "0" -PropertyType Dword
 
# Removes Chat from the Taskbar
New-itemproperty "HKLM:\Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Value "0" -PropertyType Dword
 
# Default StartMenu alignment 0=Left
New-itemproperty "HKLM:\Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAl" -Value "0" -PropertyType Dword
 
# Removes search from the Taskbar
reg.exe add "HKLM\Default\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /v SearchboxTaskbarMode /t REG_DWORD /d 0 /f
 
REG UNLOAD HKLM\Default
#>