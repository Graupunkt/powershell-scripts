<#
####################
### INFORMATIONS ###
####################
###############################################################
### ------------------------------------- GENERIC DETAILS ---------------------------------#
### SCRIPT NAME       : 1_Windows10_All In One Konfiguration Script                        #
### AUTHOR            : Graupunkt                                                          #
### Discord           : Graupunkt                                                          #
### CREATION DATE     : 10.03.2020                                                         # 
### LAST CHANGE       : 31.12.2023                                                         #
### VERIFIED OS       : Windows 10                                                         #
### VERIFIED VERSIONS : (1903),(1909), (2004), 20H2, 22H2                                  #
############################################################################################
### ---------------------------- CHANGE LOG (EU DATE FORMAT) --------------------------------------------------------------------------------------------------------#
### 29.06.2022 - Removed PCVisit and N++ from default installation because of errors
### 28.01.2022 - Renamed Script to Script_Ersteinrichtung Windows Cleint OS, Fixed Download for Notepad++ and added to default Install, Added Teamviewer to default install, Removed Java from Defaults, Replaced Certification Error with new method (not breaking each other)
### 22.11.2021 - Added new function, set Default Password for local or domain user
### 21.09.2021 - Added Customobjects for Datagroups, with Function Name, Display Name and Description
### 21.09.2021 - Solved 2x TRUE Messages in Powershell window, which where issued by the move-window function
### 06.08.2021 - TightVNC added ViewOnly password
### 30.07.2021 - Changed Method to Download Installfiles for Firefox, TightVNC, Air, Silverlight, 7Zip and Java, 
###            - Added VLC Install Function, Modified Erstinstallation Selection, Updated Function for Setting Wallpaper
### 28.07.2021 - Added Function to Remove Taskbar Weather Widget, Added Function to Disable QuickEdit Mode in Console, Moved Office Shortcuts Creation to Appearance #
###            - Added multiple Descriptions for functions 
### 20.03.2020 - Bugfixes, Silent Progression, Status Updates, Install Section added                                                                                 #
### 10.03.2020 - Added Script Informations in Header, Generic Content                                                                                                #
######################################################################################################################################################################
### --------------------------------------------------- CREDITS ---------------------------------------------------#
### https://www.tenforums.com/tutorials/7030-backup-restore-task-manager-settings.html [FOR EDITING HEX VALUES]
### https://github.com/Disassembler0/Win10-Initial-Setup-Script/blob/master/Win10.psm1
### https://github.com/unixuser011/WindowsConfig/blob/master/WindowsConfig.ps1
### https://github.com/CHEF-KOCH/Windows-10-hardening/blob/master/WD/Windows%20Defender%20-%20EMET%20Mode.ps1
### Multiple Forums, Technet and others Solutions found in past years
####################################################################################################################

#######################
### TODO / BUG LIST ###
#######################
<# W
Zeile 1146, Zeichen 5 - Der Typ [Wallpaper.Setter wurd enicht gefunden]
Set-ItemProperty Der Pfad HKCU Software Microsoft Windows CurrentVerion Internet Settings Ranges Range 6 kann nicht gefunden werden
Windows Defender - Manipulationsschutz aktivieren


# REPLACE OPENJDK WITH OLD JAVA RE
# Setting Powerplan to high power - try catch exception into write-log with result
# PROVIDE DOWNLOAD LCOATIOn FOR OEM LOGO, ACCOUNT LOGO, BGINFO SETTINGSFILE
# FiX TASKBAR UNPIN 
# FIX 3RD PARTY REMOVAL FROM START MENU
#>
#####################
###     HOWTO     ###
#####################
# DISABLE ANY FUNCTION THROUGH UNCOMMENT WITH # AT THE BEGINNING IN THE SELECTION VARIABLE
# THE FUNCTIONS LINE NUMBER IS SHOWN FOR FURTHER MODIFICATIONS
# A DESCRIPTION FOR EACH FUNCTION IS SHOWN
# AN !EXCLAMATION MARK IS SHOWN AT THE DESCRIPTIONS BEGINNING IF THIS FUNCTION DESERVES AS PREREQUEST FOR OTHER FUNCTIONS
# AN $DOLLAR SIGN IS SHOWN IF THIS FUNCTIONS NEEDS USER INTERACTION
#>

#!GLOBAL FUNCTIONS
#Runs the current Script with Admin Permissions. Is needed for most functions
Function AdminPermissions {
    If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Start-Process powershell.exe "-noProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Exit
    }
}
AdminPermissions

#Allow Net WebClient to download files, even if TLS or certi mismatch occurs
$ErrorActionPreference = 'Continue'
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

#REMOVE ALREADY EXISTING DATAGROUP VARIABLES
$DGVariables = Get-Variable -Name "DataGroup*" -Scope GLOBAL
#if($DGVariables){Get-Variable -Include $DGVariables.Name | Remove-Variable}

#Creates an empty array, primarly used for torubleshooting
$DataGroupPreconfig = @(

)

$DataGroupTest = @(
                         #Name of the Function                                   #Name to Display on GUI                          #Descriptions and Details for this function
    [PSCustomObject] @{  Function="Install_Browser_Chrome";                      DisplayName="Browser - Google Chrome";           Description="Installs Chrome"},
    [PSCustomObject] @{  Function="Install_Browser_Edge";                        DisplayName="Browser - Microsoft Edge";          Description="Installs or Updates Microsft Edge to recent chrome version, Stops if Windows ist not Up2Date"},
    [PSCustomObject] @{  Function="Install_Browser_Firefox";                     DisplayName="Browser - Mozilla Firefox";         Description="Installs Firefox"},
    [PSCustomObject] @{  Function="Install_RemoteControl_TightVNC";              DisplayName="RemoteControl - TightVNC";          Description="Installs and Configures TightVNC"},
    [PSCustomObject] @{  Function="Install_RemoteControl_PCVisit";               DisplayName="RemoteControl - PCVisit";           Description="Installs PCVisit"},
    [PSCustomObject] @{  Function="Install_RemoteControl_Teamviewer";            DisplayName="RemoteControl - TeamViewer";        Description="Installs Teamviewer"},
    [PSCustomObject] @{  Function="Install_Editor_NotepadPlusPlus";              DisplayName="Tool - Notepad++";                  Description="Installs Notepad++"},
    [PSCustomObject] @{  Function="Install_Runtimes_Java";                       DisplayName="Runtime - Java OpenJDK";            Description="Installs Java OpenJDK 32bit and 64bit"},
    [PSCustomObject] @{  Function="Install_Runtimes_Silverlight";                DisplayName="Runtime - Silverlight";             Description="Installs Microsoft Silverlight"},
    [PSCustomObject] @{  Function="Install_Runtimes_Air";                        DisplayName="Runetime - Adobe Air";              Description="Installs Adobe Air"},
    [PSCustomObject] @{  Function="Install_Archive_7Zip";                        DisplayName="Tool - 7Zip";                       Description="Installs 7Zip"},
    [PSCustomObject] @{  Function="Install_Backup_VeeamAgentforWindows";         DisplayName="Backup - Veeam Agent";              Description="Installs Backup Software Veeam Agent for Windows 3.0"},
    [PSCustomObject] @{  Function="Install_Generic_KeePass";                     DisplayName="Tool - KeePass";                    Description="Installs KeePass2 with RPC Plugin, German Language File and installs Keepass plugin in chrome"},
    [PSCustomObject] @{  Function="Install_Tool_Everything";                     DisplayName="Tool - Everything";                 Description="Installs and configures Everything, a proper search and index tool"},
    [PSCustomObject] @{  Function="Install_Generic_NetTime";                     DisplayName="Tool - NetTime";                    Description="Installs and configures NetTime with NTP ptbtime1/2/3.de and Refresh Settings for Domains"},
    [PSCustomObject] @{  Function="Install_Generic_DesktopRestore";              DisplayName="Tool - DesktopRestore";             Description="Installs a Desktop Integration Tool to save and restore desktop icons (changing monitor setups (office and home)"},
    [PSCustomObject] @{  Function="Install_RemoteControl_SolarwindsAgent";       DisplayName="RemoteControl - Solarwinds";        Description="Installs Solarwinds nCentral Windows Agent, silently"},
    [PSCustomObject] @{  Function="Install_VPN_FortiClient";                     DisplayName="VPB - FortiClient";                 Description="Installs FortiClient for VPN"},
    [PSCustomObject] @{  Function="Install_UblockFF";                            DisplayName="BrowserPlugin - Ublock Firefox";    Description="Installs UBlockOrigin Addon for Firefox"},
    [PSCustomObject] @{  Function="Install_UblockChrome";                        DisplayName="BrowserPlugin - Ublock Chrome";     Description="Installs UBlockOrigin Addon for Chrome"},
    [PSCustomObject] @{  Function="Install_UblockEdge";                          DisplayName="BrowserPlugin - Ublock Edge";       Description="Installs UBlockOrigin Addon for Edge"},
    [PSCustomObject] @{  Function="Install_PowershellV7";                        DisplayName="Google Chrome";                     Description="Installs Powershell v7"},
    [PSCustomObject] @{  Function="Install_VSCode";                              DisplayName="Tool - Visual Studio Code";         Description="Installs Visual Studio Code"}
)
    
#Extract Function, Displayname and Description from Datagroup / Hashtable (working)
    #$DataGroupTest.Function
    #$DataGroupTest.DisplayName
    #$DataGroupTest.Description

$DataGroupInstallation = @(
    "Install_Browser_Chrome"                       #Installs Chrome    
    "Install_Browser_Edge"                         #Installs or Updates Microsft Edge to recent chrome version, Stops if Windows ist not Up2Date
    "Install_Browser_Firefox"                      #Installs Firefox
    "Install_RemoteControl_TightVNC"               #Installs and Configures TightVNC
    "Install_RemoteControl_PCVisit"                #Installs PCVisit
    "Install_RemoteControl_Teamviewer"             #Installs Teamviewer
    "Install_Editor_NotepadPlusPlus"               #Installs Notepad++
    "Install_Runtimes_Java"                        #Installs Java OpenJDK 32bit and 64bit
    "Install_Runtimes_Silverlight"                 #Installs Microsoft Silverlight
    "Install_Runtimes_Air"                         #Installs Adobe Air
    "Install_Archive_7Zip"                         #Installs 7Zip
    "Install_Backup_VeeamAgentforWindows"          #Installs Backup Software Veeam Agent for Windows 3.0
    "Install_Generic_KeePass"                      #Installs KeePass2 with RPC Plugin, German Language File and installs Keepass plugin in chrome
    "Install_Tool_Everything"                      #Installs and configures Everything, a proper search and index tool
    "Install_Generic_NetTime"                      #Installs and configures NetTime with NTP ptbtime1/2/3.de and Refresh Settings for Domains
    "Install_Generic_DesktopRestore"               #Installs a Desktop Integration Tool to save and restore desktop icons (changing monitor setups (office and home)
    "Install_RemoteControl_SolarwindsAgent"        #Installs Solarwinds nCentral Windows Agent, silently
    "Install_VPN_FortiClient"                      #Installs FortiClient for VPN
    "Install_UblockFF"                             #Installs UBlockOrigin Addon for Firefox
    "Install_UblockChrome"                         #Installs UBlockOrigin Addon for Chrome
    "Install_UblockEdge"                           #Installs UBlockOrigin Addon for Edge
    "Install_PowershellV7"                         #Installs Powershell v7
    "Install_VSCode"                               #Installs Visual Studio Code
)

$DataGroupTools = @(
    "Diagnose_Internet_PingPlotter"                #Network Tool to Analyse Ping or Connection Problems
    "Install_Tool_JDAST"                           #Network Tool to Visualize and save historical data on internet speeds
    "Diagnose_Netzwerk_PingInfoView"               #Network Tool to audit packetloss, allows to ping multiple targets at once and keeps historical data
)

$DataGroupConfiguration = @(
    "Config_Explorer_DisableThumbsdbOnNetwork"     #Disables the thumbs.db files on network shares (thumbs.db creates preview pictures of pictures and videos)
    "Config_Explorer_AddPhotoViewerOpenWith"       #Adds legacy photoviewer to context of open with for pictures
    "Config_IE_DisableFirstrunWizard"              #Disables FirstRunWizard for Internet Explorer
    "Config_OneDrive_DisableAutorun"               #Removes OneDrive from Startup
    "Config_Privacy_AllowMicrophoneAccess"         #Allows Apps to use the Microphone
    "Config_System_EnableClipboardHistory"         #Enables Windows Clipboard History (Open with WIN+V)
    "Config_System_EnableStorageSense"             #Enables StorageSense, Microsofts Automated Drive Cleanup Tool
    "Config_System_DisableRebootOnBluescreen"      #Disables autoamtic reboot if a bluescreen happens
    "Config_System_ShowShutdownOptionsLockscreen"  #Enables the Option to Shutdown the PC from Loginscreen (no login needed)
    "Config_System_ShowFileOperationsDetails"      #File Transfers in Windows Explorer are minimal by default, this setting shows full details
    "Config_System_DisableSearchUnknownExtensions" #
    "Config_System_EnableNumlockOnStartup"         #Enables NUM on Startup, sometimes NUM LED is off, but NUM is still working for logn/pin
    "Config_Network_EnableLinkedConnections"       #Allows Administrative Users to access network drives
    "Config_Network_AddPrivateNetworksIntranet"    #Adds 10.x.x.x, 172.16.x.x and 192.168.x.x to Local Intranet and Websites with protocols HTTP/S to trusted sites
    "Config_Network_SetUnknownNetworksPrivate"     #
    "Config_Firewall_AllowPing"                    #Allows Ping Replies on IPv4 and IPv6
    "Config_Power_DisableFastboot"                 #Disables Fastboot
    "Config_Power_PowerPlanHigh"                   #Selects High performance power plan
    "Config_Power_DisableStandbyOnAC"              #Disables Standby when running with power supply
    "Config_Power_DisableStandbyOnBattery"         #Disables Standby when running on Battery
    "Config_Power_DisableMonitorTimeoutAC"         #Disables Monitor Timeout when running with power supply
    "Config_Power_DisableMonitorTimeoutOnBattery"  #Disables Monitor Timeout when running on Battery
    "Config_Power_DisableAllStandbyOptions"        #Disables Monitor Timeout, Disk Timeout, Standby and Hibernate on AC or Battery
    "Config_Remote_AllowRdpWithBlankPassword"      #Enable Microsoft Remote Access and allows to connect if the administrator has a blank password
    "Config_WinDefender_DisableMsAccountWarning"   #
    "Config_WinDefender_DisableOneDriveWarning"    #
    "Config_WinUpdates_DisabledDriverUpdates"      #Disables Drivers update via Windows Update
    "Config_WinUpdates_EnableMicrosoftUpdate"      #Enables Windows to download and update additional Microsoft products
    "Config_WinUpdates_DisableNightlyWakeUp"       #Disables that Windows update can wake up PC from sleep to install updates
    "Config_WinUpdates_EnableRestartSignOn"        #Enables Windows Update to log into Windows to complete updates on reboot/start
    "Config_WinUpdates_SystemDriveAsTempDir"       #Denies MSI Installer to use the biggest free drive (usually an hdd) as temporary directory for installations (this is a default behavior..) and only use C: drive
)

$DataGroupSecurity = @(
    "Hardening_Network_DisableWiFiSense"           #
    "Hardening_Store_DisableSuggestions"           #Disable App Suggestions for Windows Store
    "Hardening_System_DisableAutoplay"             #Disables AutoPlay for optical Devices
    "Hardening_System_DisableAutorun"              #Disables AutoRun for USB Sticks and optical Devices
    "Hardening_System_DisableTailoredExperience"   #
    "Hardening_System_DisableAdvertisingID"        #
    "Hardening_System_DisallowLanguage"            #
    "Hardening_System_DisableErrorReporting"       #
    "Hardening_System_DisableUserExpAndTelemtry"   #
    "Hardening_System_DisableSharedExperience"     #
    "Hardening_System_DisableWindowsScriptHost"    #Disables Windows Script Host
    "Hardening_CMD_DisableQuickEdit"               #Disables QuickEdit Mode for CMD Console Windows, that might prevent a script from running, if console accidentally is clicked
    "Hardening_Defender_EnableSmartScreen"         #Enables SmartScreen Protection for brwosing
)

$DataGroupDesign = @(
    "Appearance_Desktop_ShowBginfo"                #Installs and configures BGinfo with Default Background and preconfigured settings.bgi
    "Appearance_Desktop_ResetBackgroundToDefault"  #Reset Background to Windows Default, after Reboot
    "Appearance_Desktop_ShowComputer"              #Shows Computer on Desktop
    "Appearance_Desktop_OpenExplorerWithMyPC"      #Changes Default View from Explorer to "My PC View"
    "Appearance_Desktop_EnableDarkTheme"           #Enables the Windows Dark Theme
    "Appearance_Desktop_ShowBuildNumber"           #Shows the Current Build Number on Desktop
    "Appearance_Desktop_HideBuildNumber"           #Hides the Current Build Number from Desktop
    "Appearance_Explorer_RenameSystemDrive"        #Renames C: Drive to System
    "Appearance_Explorer_RenameDataDrive"          #Renames D: Drive to Daten
    "Appearance_Explorer_ShowFileExtensions"       #Shows File Extension for known files
    "Appearance_Explorer_ShowFullPath"             #Explorer always show the full path to current directory (documents or other personal fodlers)
    "Appearance_Explorer_ExpandCurrentFolder"      #Explorer Navigation View always expands to current folder
    "Appearance_Explorer_DisableAeroShake"         #Disables Aero Shake to minimize Windows
    "Appearance_ControlPanel_SmallSymbols"         #Changes Control panel to small icons view
    "Appearance_Taskbar_ChangeSymbols"             #Removes Cortana, Contacts and Taskview from Taskbar. Changes Searchbar to Search Icon
    "Appearance_Taskbar_UnpinUnwantedIcons"        #Unpins Icons from Taskbar, based on a Blacklist
    "Appearance_Taskbar_Remove2ndKeyboardLayout"   #Removes the secoand EN-US Keyboard Layout from System
    "Appearance_Taskbar_ShowTimeWithSeconds"       #Adds seconds to clock in taskbar
    "Appearance_Taskbar_ShowDayname"               #Adds seconds to clock in taskbar
    "Appearance_Taskbar_ShowAllIconsSystray"       #Shows all Icons in Systray
    "Appearance_TaskManager_CustomSettings"        #Changes Task Manager View to Detailed, Shows CPU View as Logical Proeessors and Shows Kernel Usage on Cores
    "Appearance_WinDefender_ShowAlwaysSystray"     #Shows Windows Defender always in SysTray
    "Appearance_Edge_DisableShortcutCreation"      #Removes the additional "shortcut" on every new shortcut that is created
    "Appearance_System_SetOEMInformations"            #Sets OEM Information at system properties, like logo and supprot contact and more # SET YOUR DETAILS IN FUNCTION FIRST; BEFORE ENABLING
    "Appearance_Account_ProfileLogo"               #Sets the Profile Picture of the current user
    "Appearance_Powershell_RunAsAdminContext"      #Adds a new contextmenu entry for powershell files "run with powershell ise as admin"
    "Appearance_System_DisableMouseShadow"         #Disables Shadow under Mouse Pointer, causes weird issues with mouse movement if used on TerminalServers. Seems to be enabled by default for all users
    "Appearance_Taskbar_DisableNewsInterestsWeather" # Disables Weather Widget on Taskbar
    "Appearance_CopyOfficeToDesktop"                #Copys Office Shortcuts to Desktop
)

$DataGroupCleanup = @(
    "Cleanup_Microsoft_FaxPrinter"                 #Removes Fax Printer 
    "Cleanup_Microsoft_XPSWriter"                  #Removes Microsoft XPS Writer
    "Cleanup_Microsoft_FaxAndScanServices"         #Removes Fax and ScanServices (not more present in Win10 - 2004)
    "Cleanup_Microsoft_StoreApps"                  #Removes unwanted apps from startmenu via Whitelist
        # Uninstalls Hello.Face, Powershell ISE, Druckverwaltungskonsole, Wordpad, Editor, Wordpad and Windows-Speicherverwaltung - add to Whitelist
        # apps and features -> optional features -> install to get them back
        # Does not unpin Windows Store and xxx from Taskbar
        # Install SNMP?
    "Cleanup_ThirdParty_McAfeeLive"                #$Tries to remove McAfee LiveSafe (with GUI only, no silent uninstall available)
    "Cleanup_ThirdParty_MultiAvRemoval"            #Uses ESET's uninstall tool to remove the most common A/V products in silent mode (does not recognize mcafee livesafe)
    "Cleanup_RemoteControl_SolarwindsAgent"        #  
    "Uninstall_Tool_OfficeRemoval"                 #Removes currently installed Microsoft Office from PC (liek click to run, pre installed
)

$DataGroupUpdates = @(
    "Update_StoreApps_All"                        #Tries to force an update of all installed Windows Store Apps
    "Update_Windows_CurrentlyAvailable"           #Scans and installs currently via windows update available updates
    "Update_Windows_WsusOffline"                  #Uses a prepared network share to run and install all windows updates incl. multiple reboots
    "Update_Windows_PSWindowsUpdate"              #
)

$DataGroupAdministrative = @(
    "Domain_DisallowLocalLogin"                   #
    "Domain_JoinComputer"                         #$Requires manual provision of domain admin credentials
    "Generic_Useraccount_CreateAdmin"             #
    "Generic_Useraccount_DefaultPassword"         #Sets our internal Default Password for the current user
    "Generic_Useraccount_HideUsersLoginscreen"    #Hides specific Default Users from Win10 Loginscreen
    "Generic_Explorer_Restart"                    #Forces a Refresh of Windows Explorer/Desktop, closes all open explorers windows
    "PreConfig_Network_SetHostnameByMac"          #
    "PreConfig_Network_SetNetworkSettingsbyMac"   #
    "PreConfig_SystemDrive_CreateDefaultDirs"     #
    "Config_System_HostnameSeriennummer"          #
)
function Create_DynamicFormMainframe{
    
    # LOAD PRERQUESTS FOR SYSTEM WINDOWS FORMS
    try {
        #[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  | Out-Null
        #[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  | Out-Null
        Add-Type -assembly System.Windows.Forms
        Add-Type -assembly System.Drawing
    }
    catch {
        Write-Warning -Message 'Unable to load required assemblies'
        return
    }
    if(!$script:MainForm -eq $null){Clear-Variable -Name $script:MainForm}
    

    #GET ALL VARIABLES (ARRAYS) THAT CONTAIN DATAGROUP IN THEIR NAME 
    $AllDGVariables = Get-Variable -Include "DataGroup*"
    $script:FinalGroupsForForm = @()
    foreach ($DGVariable in $AllDGVariables){
        $script:FinalGroupsForForm += New-Object psobject -Property @{
            Groups = $DGVariable.Name.Replace("DataGroup","")
            Options = $DGVariable.Value
        }
    }
    #$FinalGroupsForForm | FT -Property Groups, Options

    #Count Groups
    $NumberFormGroups = $FinalGroupsForForm.Groups.Count
    $NumberFormGroups = 3 

    #Count max Options
    [System.Collections.ArrayList]$ListOfOptions = @()
    foreach ($FormGroup in $FinalGroupsForForm){
        $ListOfOptions.add($FormGroup.Options.Count) | Out-Null
    }

    #DEFINE SIZE, BORDERS & CO
    $OptionHeight = 24
    $OptionWidth = 295
    $OptionSpacer = 5
    $GroupBoxWidth = $OptionWidth + 60
    $GroupBoxSpacing = 20
    $script:MainFormVerticalSpace = 150
    $script:MainFormVerticalBorder = 10
    $script:MainFormHorizontalBorder = 10


    $NumberFormMostOptions = $ListOfOptions | Sort-Object -Descending | Select-Object -First 1
    $script:MainFormWidth = $GroupBoxSpacing + $script:MainFormHorizontalBorder + $NumberFormGroups * $GroupBoxWidth
    $script:MainFormHeight = $script:MainFormVerticalSpace + $script:MainFormVerticalBorder + $NumberFormMostOptions * $OptionHeight
    #CHECK IF GROUPS OR OPTIONS EXCEED SCREENSIZE
    $ScreenResolution = [System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.DeviceName -like "*DISPLAY1"}  #GET MAX SCREEN RESOLUTION
    $MaxFormSizeWidth = $ScreenResolution[0].WorkingArea.Width         #GET MAX WIDTH FROM PRIMARY SCREEN
    $MaxFormSizeHeight = $ScreenResolution[0].WorkingArea.Height       #GET MAX HEIGHT FROM PRIMARY SCREEN

    #IF FORM EXCEEDS SCREEN SIZE
    if ($script:MainFormWidth -gt $MaxFormSizeWidth -OR $script:MainFormHeight -gt $MaxFormSizeHeight){
        Write-Warning "Current form will exceed screen limits $($MaxFormSizeWidth):$($MaxFormSizeHeight) Form $($script:MainFormWidth):$($script:MainFormHeight)"
    }

    #Create Main Form
    $TotalFormWidth = $script:MainFormWidth
    $TotalFormHeight = $script:MainFormHeight + $script:MainFormVerticalBorder + $script:MainFormVerticalBorder * 6 
    $script:MainForm = New-Object System.Windows.Forms.Form                                    # Create new object
    $script:MainForm_Drawing_Size = New-Object System.Drawing.Size
    $script:MainForm_Drawing_Size.Width = $TotalFormWidth
    $script:MainForm_Drawing_Size.Height = $TotalFormHeight
    $script:MainForm.Size = $script:MainForm_Drawing_Size
    $script:MainForm.Text = "Window All in One Script"                                         # Window Title
    $script:MainForm.StartPosition = 'CenterScreen'                                            # Position on Screen
    $script:MainForm.DataBindings.DefaultDataSourceUpdateMode = 0

   

    #Create Tabs
    $TabControl = New-object System.Windows.Forms.TabControl
    $script:TabConfiguration  = New-Object System.Windows.Forms.TabPage
    $script:TabInstallation   = New-Object System.Windows.Forms.TabPage
    $script:TabDesign         = New-Object System.Windows.Forms.TabPage
    $script:TabAdministrative = New-Object System.Windows.Forms.TabPage
    $script:TabSecurity       = New-Object System.Windows.Forms.TabPage
    $script:TabCleanup        = New-Object System.Windows.Forms.TabPage
    $script:TabUpdates        = New-Object System.Windows.Forms.TabPage
    $script:TabTools          = New-Object System.Windows.Forms.TabPage

    #Tab Control 
    $tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 0
    $System_Drawing_Point.Y = 70
    $tabControl.Location = $System_Drawing_Point
    $tabControl.Name = "tabControl"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = $TotalFormHeight #- $FormExtraHeightTabs
    $System_Drawing_Size.Width = $TotalFormWidth
    $tabControl.Size = $System_Drawing_Size
    $script:MainForm.Controls.Add($tabControl)

    #Configuration
    $TabConfiguration.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabConfiguration.UseVisualStyleBackColor = $True
    $TabConfiguration.Name = "Configuration"
    $TabConfiguration.Text = "Configuration"
    $TabConfiguration.AutoScroll = $True
    #SET CORRECT HEIGHT AND WIDTH
    $TabConfiguration.AutoScrollMinSize = New-Object System.Drawing.Size($script:MainFormHeight,$script:MainFormWidth)
    $tabControl.Controls.Add($script:TabConfiguration)

    #Installation
    $TabInstallation.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabInstallation.UseVisualStyleBackColor = $True
    $TabInstallation.Name = "Installation"
    $TabInstallation.Text = "Installation"
    $tabControl.Controls.Add($script:TabInstallation)

    #Design
    $TabDesign.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabDesign.UseVisualStyleBackColor = $True
    $TabDesign.Name = "Design"
    $TabDesign.Text = "Design"
    $TabDesign.AutoScroll = $True
    #SET CORRECT HEIGHT AND WIDTH
    $TabDesign.AutoScrollMinSize = New-Object System.Drawing.Size($script:MainFormHeight,$script:MainFormWidth)
    $tabControl.Controls.Add($script:TabDesign)

    #Administrative
    $TabAdministrative.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabAdministrative.UseVisualStyleBackColor = $True
    $TabAdministrative.Name = "Administrative"
    $TabAdministrative.Text = "Administrative"
    $tabControl.Controls.Add($script:TabAdministrative)

    #Security
    $TabSecurity.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabSecurity.UseVisualStyleBackColor = $True
    $TabSecurity.Name = "Security"
    $TabSecurity.Text = "Security"
    $tabControl.Controls.Add($script:TabSecurity)

    #Cleanup
    $TabCleanup.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabCleanup.UseVisualStyleBackColor = $True
    $TabCleanup.Name = "Cleanup"
    $TabCleanup.Text = "Cleanup"
    $tabControl.Controls.Add($script:TabCleanup)

    #Update
    $TabUpdates.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabUpdates.UseVisualStyleBackColor = $True
    $TabUpdates.Name = "Update"
    $TabUpdates.Text = "Update"
    $tabControl.Controls.Add($script:TabUpdates)

    #Tools
    $TabTools.DataBindings.DefaultDataSourceUpdateMode = 0
    $TabTools.UseVisualStyleBackColor = $True
    $TabTools.Name = "Tools"
    $TabTools.Text = "Tools"
    $tabControl.Controls.Add($script:TabTools)

    #DYNAMIC BOXES, FOR EACH DATAGROUP ONE BOX TO THE RIGHT
    $GroupBoxCounter = 0
    foreach ($Group in $FinalGroupsForForm){
        $GroupBoxCounter++
        $ListOfGroupOptions = $Group.Options.Count

        #SET GROUPBOX 
        $groupBox = New-Object System.Windows.Forms.GroupBox
        #SET POSITION OF GROUPBOX
        $GroupBox_Drawing_Size = New-Object System.Drawing.Size
        #$GroupBox_Drawing_Size.Width = $GroupBoxSpacing + (($GroupBoxCounter -1) * $GroupBoxWidth)
        $GroupBox_Drawing_Size.Width = $GroupBoxSpacing
        $GroupBox_Drawing_Size.Height = $script:MainFormVerticalSpace 
        $GroupBox_Drawing_Size.Height = 10
        $groupBox.Location = $GroupBox_Drawing_Size
        $groupBox.text = $Group.Groups
        $groupBox.Name = "GroupBox_$($Group.Groups)"
        #SET SIZE OF OF GROUPBOX 
        $GroupBox_Option_Drawing_Size = New-Object System.Drawing.Size
        $GroupBox_Option_Drawing_Size.Width = $OptionWidth + $GroupBoxSpacing
        $GroupBox_Option_Drawing_Size.Height = $GroupBoxSpacing + 2 * $OptionSpacer + ($OptionHeight) * $ListOfGroupOptions
        $groupBox.size = $GroupBox_Option_Drawing_Size
        #$script:MainForm.Controls.Add($groupBox)
        #$script:LegacyOverview.Controls.Add($groupBox)
        switch -wildcard ($groupBox.Name){
            "*Configuration"     {$script:TabConfiguration.Controls.Add($groupBox)}
            "*Installation*"     {$script:TabInstallation.Controls.Add($groupBox)}
            "*Design*"           {$script:TabDesign.Controls.Add($groupBox)}
            "*Administrative*"   {$script:TabAdministrative.Controls.Add($groupBox)}
            "*Security*"         {$script:TabSecurity.Controls.Add($groupBox)}
            "*Cleanup*"          {$script:TabCleanup.Controls.Add($groupBox)}
            "*Updates*"          {$script:TabUpdates.Controls.Add($groupBox)}
            "*Tools*"            {$script:TabTools.Controls.Add($groupBox)}
        }

        #SET OPTIONS
        $OptionCounter = 0
        foreach ($Option in $Group.Options){
            $OptionCounter++
            $script:OptionBox = New-Object System.Windows.Forms.CheckBox 
            #SET POSITION OF OPTION
            $Option_Drawing_Size = New-Object System.Drawing.Size
            $Option_Drawing_Size.Width = $OptionSpacer
            $Option_Drawing_Size.Height = $OptionCounter * $OptionHeight
            $OptionBox.Location = $Option_Drawing_Size
            #SET SIZE OF OPTION
            $OptionBox.size = New-Object System.Drawing.Size($OptionWidth,$OptionHeight) 
            $OptionBox.Text = $Option
            $OptionBox.Name = "Option_$($Option)"
            $groupBox.Controls.Add($OptionBox) 
        }
    }

    # DEFINE RUN BUTTON
    $RunButton = New-Object System.Windows.Forms.Button 
    $RunButton.Location = New-Object System.Drawing.Size(30,10) 
    $RunButton.Size = New-Object System.Drawing.Size(80,40) 
    $RunButton.Text = "RUN" 
    $script:MainForm.Controls.Add($RunButton)
    $RunButton.Add_Click({
        $script:FormProperlyClosed = $true
        $script:MainForm.Close()
    })

    #DEFINE CANCEL BUTTON
    #$CancelButton = New-Object System.Windows.Forms.Button 
    #$CancelButton.Location = New-Object System.Drawing.Size(130,10) 
    #$CancelButton.Size = New-Object System.Drawing.Size(80,40) 
    #$CancelButton.Text = "CANCEL" 
    #$script:MainForm.Controls.Add($CancelButton)
    #$CancelButton.Add_Click({
    #    $script:MainForm.Close()
    #})

    # DEFINE LOAD BUTTON, TO GIVE PREREQUESTS
    $script:LoadButton = New-Object System.Windows.Forms.Button 
    $script:LoadButton.Location = New-Object System.Drawing.Size(130,10) 
    $script:LoadButton.Size = New-Object System.Drawing.Size(100,40) 
    $script:LoadButton.Text = "Erstinstallation" 
    $MainForm.Controls.Add($script:LoadButton)
}
Create_DynamicFormMainframe

# DONT RUN FUNCTIONS IF FORM GETS CLOSED OR ANy OTHER ERROR BEFORE RUN WAS CLICKED

#ADD ALL FUNCTIONS TO A SINGLE ARRAY
$CollectionOfTabs = @()
$CollectionOfTabs  += $script:TabConfiguration
$CollectionOfTabs  += $script:TabInstallation
$CollectionOfTabs  += $script:TabDesign
$CollectionOfTabs  += $script:TabAdministrative
$CollectionOfTabs  += $script:TabSecurity
$CollectionOfTabs  += $script:TabCleanup
$CollectionOfTabs  += $script:TabUpdates
$CollectionOfTabs  += $script:TabTools

#SET COLORS FOR INACTIVE CHECKBOXES
foreach ($Box in $CollectionOfTabs.Controls){    #FOR TAB1 FOR MAIN SCRIPT USE #foreach ($Box in $script:MainForm.Controls){
    foreach ($Checkbox in $Box.Controls){
        if($Checkbox.Name -like "*PreConfig_Network_SetHostnameByMac"){$Checkbox.ForeColor = [System.Drawing.Color]::Gray}
        if($Checkbox.Name -like "*PreConfig_Network_SetNetworkSettingsbyMac"){$Checkbox.ForeColor = [System.Drawing.Color]::Gray}
        #if($Checkbox.Name -like "*Install_RemoteControl_PCVisit"){$Checkbox.ForeColor = [System.Drawing.Color]::Red}
        #if($Checkbox.Name -like "*Install_RemoteControl_Teamviewer"){$Checkbox.ForeColor = [System.Drawing.Color]::Red}
        if($Checkbox.Name -like "*Cleanup_Microsoft_StoreApps"){$Checkbox.ForeColor = [System.Drawing.Color]::Red}
    }
}

#ENABLE CHECKBOXES FOR BUTTON LOAD
$script:LoadButton.Add_Click({
    foreach ($Box in $CollectionOfTabs.Controls){    #FOR TAB1 FOR MAIN SCRIPT USE #foreach ($Box in $script:MainForm.Controls){
        foreach ($Checkbox in $Box.Controls){
            #if($Checkbox.Name -like "*Config_System_HostnameSeriennummer"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*PreConfig_SystemDrive_CreateDefaultDirs"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_DisableUAC"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_DisableQuickBoot"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_PowerPlanHigh"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_CopyOfficeToDesktop"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_RemoteControl_TightVNC"){$Checkbox.Checked = $true}  # SET YOUR DETAILS IN FUNCTION FIRST; BEFORE ENABLING
            #if($Checkbox.Name -like "*Install_RemoteControl_PCVisit"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Runtimes_Java"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Mediaplayer_VLC"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_RemoteControl_Teamviewer"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Browser_Edge"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Browser_Firefox"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Browser_Chrome"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Editor_NotepadPlusPlus"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Install_Archive_7Zip"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Install_Tool_Everything"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Taskbar_ChangeSymbols"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Taskbar_UnpinUnwantedIcons"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Taskbar_Remove2ndKeyboardLayout"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Taskbar_ShowTimeWithSeconds"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Desktop_ResetBackgroundToDefault"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Desktop_ShowComputer"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Desktop_OpenExplorerWithMyPC"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Desktop_EnableDarkTheme"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_RenameSystemDrive"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_RenameDataDrive"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_ShowFileExtensions"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_ShowFullPath"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_ExpandCurrentFolder"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Explorer_DisableAeroShake"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_ControlPanel_SmallSymbols"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_TaskManager_CustomSettings"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Appearance_Edge_DisableShortcutCreation"){$Checkbox.Checked = $true}
            #if($Checkbox.Name -like "*Appearance_System_SetOEMInformations"){$Checkbox.Checked = $true} # SET YOUR DETAILS IN FUNCTION FIRST; BEFORE ENABLING
            if($Checkbox.Name -like "*Appearance_Taskbar_DisableNewsInterestsWeather"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Cleanup_Microsoft_FaxPrinter"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Cleanup_Microsoft_XPSWriter"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Explorer_DisableThumbsdbOnNetwork"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Explorer_AddPhotoViewerOpenWith"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_IE_DisableFirstrunWizard"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_EnableClipboardHistory"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_EnableStorageSense"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_DisableRebootOnBluescreen"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_ShowShutdownOptionsLockscreen"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_ShowFileOperationsDetails"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_DisableSearchUnknownExtensions"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_System_EnableNumlockOnStartup"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Network_EnableLinkedConnections"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Network_AddPrivateNetworksIntranet"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Firewall_AllowPing"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_PowerPlanHigh"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableFastboot"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableStandbyOnAC"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableMonitorTimeoutAC"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableStandbyOnBattery"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableMonitorTimeoutOnBattery"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_Power_DisableAllStandbyOptions"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_WinDefender_DisableMsAccountWarning"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_WinUpdates_EnableMicrosoftUpdate"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_WinUpdates_DisableNightlyWakeUp"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_WinUpdates_EnableRestartSignOn"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Config_WinUpdates_SystemDriveAsTempDir"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Fix_Powershell_IseFilenameComma"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_Network_DisableWiFiSense"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_Store_DisableSuggestions"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableAutoplay"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableAutorun"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableTailoredExperience"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableAdvertisingID"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisallowLanguage"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableErrorReporting"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableUserExpAndTelemtry"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_System_DisableSharedExperience"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Hardening_Defender_EnableSmartScreen"){$Checkbox.Checked = $true}
            if($Checkbox.Name -like "*Generic_Explorer_Restart"){$Checkbox.Checked = $true}
        }
    }
})

#SHOW FORM TO USER
$script:MainForm.ShowDialog()| Out-Null


#ADD ALL CHECKED CHECKBOXE FUNCTIONS TO EXECUTION LIST
$script:ExecutionList = @()
foreach ($Box in $CollectionOfTabs.Controls){
    foreach ($Checkbox in $Box.Controls){
        if($Checkbox.Checked){$ExecutionList += $Checkbox.Name.replace('Option_',"")}
    }
}




### FORM END ###


#!GLOBAL TEMP FOLDER
$global:TempFolder = "C:\Temp"
If (Test-Path $global:TempFolder) { }Else {New-Item -ItemType Directory $global:TempFolder | Out-Null}

#! LOGFILE
#Scriptname from Powershell
if(!$ScriptFileName){$ScriptFileName = $MyInvocation.MyCommand.Name}
#Scriptname from ISE
if(!$ScriptFileName){
    $ScriptFileName = ([Environment]::GetCommandLineArgs()[1]).Split('\')[-1]
    $ScriptFileName = $ScriptFileName.replace("-NoProfile","") -replace '\s',''
}
#Scriptname from VSCODe
if(!$ScriptFileName){$ScriptFileName = ($psEditor.GetEditorContext().CurrentFile.Path).Split('\')[-1]}

$logfile = "C:\ADMIN\LOGS\Protokoll_$ScriptFileName.log"
If (Test-Path "C:\ADMIN\LOGS") { }Else {New-Item -ItemType Directory "C:\ADMIN\LOGS" | Out-Null}
function Write-Log([string]$Value1, [int]$level = 0) {
        #Create Mutex, to access the same file multiple times
        $mtx = New-Object System.Threading.Mutex($false, "LogMutex")

        $logdate = get-date -format "yyyy-MM-dd HH:mm:ss"
        # VERBOSE, DEBUG, NORMAL INFO
        if ($level -eq 0) {
            $logtext = "[INFO] " + $Value1 
            $text = "[" + $logdate + "] - " + $logtext
            Write-Host $text -ForegroundColor Green #-Separator ":"
        }
        # WARNINGS
        if ($level -eq 1) {
            $logtext = "[WARNING] " + $Value1
            $text = "[" + $logdate + "] - " + $logtext
            Write-Host $text -ForegroundColor Yellow #-Separator ":"
        }
        # ERRORS
        if ($level -eq 2) {
            $logtext = "[ERROR] " + $Value1
            $text = "[" + $logdate + "] - " + $logtext
            Write-Host $text -ForegroundColor Red #-Separator ":"
        }

        # DEBUG
        if ($level -eq 9) {
            $logtext = "[DEBUG] " + $Value1
            $text = "[" + $logdate + "] - " + $logtext
            Write-Host $text -ForegroundColor Gray #-Separator ":"
        }

        #WRITE ALL LOGS INTO A FILE

        #WAITS FOR MUTEX TO BE ACCCESSABLE, IN CASE OTHER PROCESSES ARE USING THE MUTUX TO WRITE TO LOGFILE 
        if ($mtx.WaitOne()){
            #$text >> $using:logfile
            $text >> $logfile
            [void]$mtx.ReleaseMutex() 
        }
        else {
            #Mutex timed out, but not timeout is specified by default
        }
        $mtx.Dispose()
    }
#Write-Log "Grün, Info" 0
#Write-Log "Gelb, Warning" 1
#Write-Log "Rot, Fehler" 2
#Write-Log "Grau, Debug" 9

#SET CURRENT WINDOWS SIZE
Function Set-WindowSize{
    param(
        [int]$newXSize,
        [int]$newYSize
       )
    [Console]::WindowHeight=$newXSize;
    [Console]::WindowWidth=$newYSize;
    #[Console]::WindowHeight=15
    #[Console]::WindowWidth=$newYSize
}    

function Move-Window {
    param(
        [int]$newX,
        [int]$newY
    )
    BEGIN {
    $signature = @'

[DllImport("user32.dll")]
public static extern bool MoveWindow(
    IntPtr hWnd,
    int X,
    int Y,
    int nWidth,
    int nHeight,
    bool bRepaint);

[DllImport("user32.dll")]
public static extern IntPtr GetForegroundWindow();

[DllImport("user32.dll")]
public static extern bool GetWindowRect(
    HandleRef hWnd,
    out RECT lpRect);

public struct RECT
{
    public int Left;        // x position of upper-left corner
    public int Top;         // y position of upper-left corner
    public int Right;       // x position of lower-right corner
    public int Bottom;      // y position of lower-right corner
}

'@
    
    Add-Type -MemberDefinition $signature -Name Wutils -Namespace WindowsUtils
    
    }
    PROCESS{
        $phandle = [WindowsUtils.Wutils]::GetForegroundWindow()
    
        $o = New-Object -TypeName System.Object
        $href = New-Object -TypeName System.RunTime.InteropServices.HandleRef -ArgumentList $o, $phandle
    
        $rct = New-Object WindowsUtils.Wutils+RECT
    
        [WindowsUtils.Wutils]::GetWindowRect($href, [ref]$rct) | Out-Null
        
        $width = $rct.Right - $rct.Left
        $height = 700
    <#
        $height = $rct.Bottom = $rct.Top
        
        $rct.Right
        $rct.Left
        $rct.Bottom
        $rct.Top
        
        $width
        $height
    #>
        [WindowsUtils.Wutils]::MoveWindow($phandle, $newX, $newY, $width, $height, $true) | Out-Null
    
    }
}

#SET WINDOW SIZE AN DPOSITION
$ScreenResolution = [System.Windows.Forms.Screen]::AllScreens
$WindowSizeX4k = 75
$WindowSizeY4k = 220
$WindowSizeXHD = 50
$WindowSizeYHD = 180
$MaxX = $ScreenResolution[0].WorkingArea.Width
$MaxY = $ScreenResolution[0].WorkingArea.Height
#if($MaxX -gt 2000){Set-WindowSize $WindowSizeX4k $WindowSizeY4k}
#elseif($MaxX -gt 1300){Set-WindowSize $WindowSizeXHD $WindowSizeYHD}
if($script:FormProperlyClosed){Move-Window 25 25} 


#Determines Windows Version, Windows Build Number, Language Type and Bittype and provides global parameters for these
Function Determine_BittypeLanguageVersion {
    # IF WINDOW IS RESIZED THIS KILLS LINE ALIGNMENT, DISABLE TO KEEP LINES WRAPPED TO WINDOW SIZE
    [console]::bufferwidth = 32766
    Write-Log "Determining Host Bittype, Language, Version and Release" 0
    $global:language = $null
    $global:bittype = $null
    $global:WindowsVersion = $null
    $global:WindowsRelease = $null
    
    $global:language=Get-WinSystemLocale | Select-Object -ExpandProperty Name
    $global:bittype=(Get-CimInstance -ClassName win32_operatingsystem).OSArchitecture
    $global:WindowsVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" -Name CurrentVersion).CurrentVersion
    $global:WindowsRelease = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" -Name ReleaseID).ReleaseId
    $global:WindowsBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" -Name CurrentBuild).CurrentBuild
    $global:WindowsInstallDate = ((Get-Item $env:windir).creationtime).ToString('yyyy.MM.dd-HH:mm:ss')
    switch($global:WindowsVersion){
        6.3 { $VersionHumanReadable = "Windows 10" }
        6.2 { $VersionHumanReadable = "Windows 8" }
        6.1 { $VersionHumanReadable = "Windows 7" }
        6.0 { $VersionHumanReadable = "Windows Vista" }
        5.2 { $VersionHumanReadable = "Server 2003" }
        5.1 { $VersionHumanReadable = "Windows XP" }
    }
    Write-Log "Bittyp:$global:bittype Sprache:$global:language Version:$global:WindowsVersion ($VersionHumanReadable) Release:$global:WindowsRelease Build:$global:WindowsBuild Installation:$global:WindowsInstallDate" 9
    Start-Sleep -Seconds 1
}
if($script:FormProperlyClosed){Determine_BittypeLanguageVersion}

#Allows to Download directly from cloud server, ignoring hostname not matching the ssl certificate
Function Resolve-DownloadURL {
    #RESET ALL PARAMETERS TO NULL
    #$URLHasExtension = $FileExtension = $request = $response = $LoopIteration = $URLResolvedSuccessfully = $FileName = $ExtractedFiles = $null

    #PERMALINK TO DIRECT FILE LINK
    if ($global:bittype -eq "64-Bit") {$URLRaw = $Global:DownloadURL64}
    if ($global:bittype -eq "32-Bit") {$URLRaw = $Global:DownloadURL32}

    #CHECK IF LINK IS POINTING DIRECTLY TO A FILE, IF NOT START RESOLVING PROCESS 
    $FileExtension = [IO.Path]::GetExtension($URLRaw)
    $URLHasExtension = $FileExtension -eq ".zip" -OR $FileExtension -eq ".exe" -OR $FileExtension -eq ".msi"
    if (!$URLHasExtension){$URLHasExtension = $false}

    #RESOLVE URL TO DOWNLOADLINK, LIMIT: 5
    $LoopIteration = 0
    while ($URLHasExtension -eq $false -AND $LoopIteration -lt 5){
        $LoopIteration ++
        Write-Log "$Global:ApplicationName - Resolving Link to Downloadfile #$LoopIteration" 9
        $request = [System.Net.WebRequest]::Create($URLRaw)
        $request.AllowAutoRedirect=$false
        $request.Timeout=5000;
        $response=$request.GetResponse()
        If ($response.StatusCode -eq "Found")
        {
            $URLRaw = $response.GetResponseHeader("Location")
        }
        $Global:FileExtension = [IO.Path]::GetExtension($URLRaw)
        $URLHasExtension = $Global:FileExtension -eq ".zip" -OR $FileExtension -eq ".exe" -OR $FileExtension -eq ".msi"
        Start-Sleep -Seconds 1
    }
    if ($URLHasExtension){
        $Global:FileName = [System.IO.Path]::GetFileName($URLRaw)
        $Global:ExeFileName = $Global:FileName
        $Global:FileExtension = [IO.Path]::GetExtension($URLRaw)
        $URLResolvedSuccessfully = $true
        $Global:DowloadFinalURL = $URLRaw 
    }

    #DEBUG
    #Write-log "URLRAW4 = $UrlRaw" 9
    #Write-log "FileExtension = $FileExtension" 9
    #Write-log "Global:ExeFileName2 = $Global:ExeFileName" 9
    #Write-log "Global:FileName2 = $Global:FileName" 9
    #Write-log "Global:FileExtension = $Global:FileExtension" 9
}
Function Get-FileWebclient {
    param(
        [Parameter(Mandatory=$true)]
        $url, 
        $destinationFolder="$global:TempFolder",
        [switch]$includeStats
    )

    Write-Log "Downloading $Global:FileName from $Global:DowloadFinalURL" 9

    #ENABLE DOWNLOADS FROM HOSTS WITH DIFFERENT CERTIFICATE THAN HOSTNAME
    if ("TrustAllCertsPolicy" -as [type]) {} else {
        Add-Type "using System.Net;using System.Security.Cryptography.X509Certificates;public class TrustAllCertsPolicy : ICertificatePolicy {public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem) {return true;}}"
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    }

    #WAIT UNTIl ASYNC DOWNLOAD IS FINISHED
    $global:WaitForDownload = 0

    $wc = New-Object Net.WebClient
    $wc.UseDefaultCredentials = $true
    $file = $url | Split-Path -Leaf
    $destination = Join-Path $destinationFolder $file
    $start = Get-Date 
    $null = Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged `
        -MessageData @{start=$start;includeStats=$includeStats} `
        -SourceIdentifier WebClient.DownloadProgressChanged -Action { 
            filter Get-FileSize {
	            "{0:N2} {1}" -f $(
	            if ($_ -lt 1kb) { $_, 'Bytes' }
	            elseif ($_ -lt 1mb) { ($_/1kb), 'KB' }
	            elseif ($_ -lt 1gb) { ($_/1mb), 'MB' }
	            elseif ($_ -lt 1tb) { ($_/1gb), 'GB' }
	            elseif ($_ -lt 1pb) { ($_/1tb), 'TB' }
	            else { ($_/1pb), 'PB' }
	            )
            }
            $elapsed = ((Get-Date) - $event.MessageData.start)
            #calculate average speed in Mbps
            $averageSpeed = ($EventArgs.BytesReceived * 8 / 1MB) / $elapsed.TotalSeconds
            $elapsed = $elapsed.ToString('hh\:mm\:ss')
            #calculate remaining time considering average speed
            $remainingSeconds = ($EventArgs.TotalBytesToReceive - $EventArgs.BytesReceived) * 8 / 1MB / $averageSpeed
            $receivedSize = $EventArgs.BytesReceived | Get-FileSize
            $totalSize = $EventArgs.TotalBytesToReceive | Get-FileSize        
            if ($EventArgs.ProgressPercentage -eq 1 -OR $EventArgs.ProgressPercentage -eq 10 -OR $EventArgs.ProgressPercentage -eq 20 -OR $EventArgs.ProgressPercentage -eq 30 -OR $EventArgs.ProgressPercentage -eq 40 -OR $EventArgs.ProgressPercentage -eq 50 -OR $EventArgs.ProgressPercentage -eq 60 -OR $EventArgs.ProgressPercentage -eq 70 -OR $EventArgs.ProgressPercentage -eq 80 -OR $EventArgs.ProgressPercentage -eq 90 -OR $EventArgs.ProgressPercentage -eq 100){
                Write-Progress -Activity (" {0:N2} Mbps" -f $averageSpeed) `
                -Status ("{0} of {1} ({2}% in {3})" -f $receivedSize,$totalSize,$EventArgs.ProgressPercentage,$elapsed) `
                -SecondsRemaining $remainingSeconds `
                -PercentComplete $EventArgs.ProgressPercentage
            }
            if ($EventArgs.ProgressPercentage -eq 100){
                 Write-Progress -Activity (" {0:N2} Mbps" -f $averageSpeed) `
                -Status 'Done' -Completed
                if ($event.MessageData.includeStats.IsPresent){
                    $global:WaitForDownload = 1
                    #([PSCustomObject]@{Name=$destination;TotalSize=$totalSize;Time=$elapsed}) | Out-Host
                }
            }
        }    
    $null = Register-ObjectEvent -InputObject $wc -EventName DownloadFileCompleted `
         -SourceIdentifier WebClient.DownloadFileCompleted -Action { 
            Unregister-Event -SourceIdentifier WebClient.DownloadProgressChanged
            Unregister-Event -SourceIdentifier WebClient.DownloadFileCompleted
            Get-Item $destination | Unblock-File
        }  
    try  {  
        $wc.DownloadFileAsync($url, $destination)  
    }  
    catch [System.Net.WebException]  {  
        Write-Warning "Download of $url failed"  
    }   
    finally  {    
        $wc.Dispose()
        
    } 
 while (!$global:WaitForDownload -eq 1){}
 }
 Function Extract-Archive {
    #param(
    #[Parameter(Mandatory=$true)]
    #$Archive
    #)

    Write-Log "Extracting Archive $Archive" 9
    $Archive = join-path "$Global:TempFolder\" "$Global:FileName"
    #Powershell V7 and newer
    #$ExtractedFiles = (Expand-Archive -LiteralPath "$global:TempFolder\$FileName" -DestinationPath $global:TempFolder -Force -PassThru).Name
    #REPLACE FILENAME WITH EXECUTABLE FOR SETUP
    #$Global:FileName = $ExtractedFiles | Where-Object {[IO.Path]::GetExtension($_)} | Select-Object -First 1
       
    #Powershell V5 and earlier
    if ($Global:FileExtension -eq ".zip"){Expand-Archive -Path $Archive -DestinationPath "C:\Temp" -Force}
    if ($Global:FileExtension -eq ".zip"){$global:ExeFileName = (Get-ChildItem $global:TempFolder *.exe).Name | Select-Object -First 1}
}
Function Clear-TempDir {        
    Start-Sleep -s 3
    Remove-Item "$global:TempFolder\*.*" -Recurse -Force
}
Function Set_GlobalTempDir {
    $Global:DowloadFinalURL = $null
    Push-Location $global:TempFolder
    $global:TempFolder = "C:\Temp"
    If (!(Test-Path "C:\Temp")) {
        New-Item -Path "C:\Temp" -ItemType "directory" -Force | Out-Null
    }

}


##################
### PRE-CONFIG ###
##################

Function PreConfig_Network_SetHostnameByMac {
    Write-Log "Setting Hostname..."
    #################
    ### PARAMETER ###
    #################
    #IMPORT HOSTNAME CSV
    $ScriptDir = Split-Path $psise.CurrentFile.FullPath
    Push-Location $ScriptDir
    $CSV=Import-Csv 0_Table-for-MAC-Hostname-IP.csv -Delimiter ";"
}


##############
### SCRIPT ###
##############
#GET ALL MACS FROM ALL NICS
#$Mac=(Get-WmiObject win32_networkadapterconfiguration).macaddress

#VALIDATE CSV LIST FOR MATCHING MAC
<#
$Match=@()
$Mac | 
ForEach-Object {
    if($CSV.MACAddress -contains $_)
    {
	    $Match+=$_
    }
}
#>

#SET HOSTNAME FROM CSV
<#
$Result = $CSV|Where-Object{$_.MACAddress -match $Match}
$Hostname = $Result.HostName
#>

#GET AD CREDENTIALS
<#
Get-Content 0_AD-Credentials.txt | Where {$_ -notmatch '^#.*'} | Foreach-Object{
   $var = $_.Split('=')
   Set-Variable -Name $var[0] -Value $var[1]
}
$AD_Username
$AD_Password
#>

#RENAME PC ACCORDING TO MAC ADRESS
#Rename-Computer -NewName "$Hostname" # -Restart
#}
Function PreConfig_Network_SetNetworkSettingsbyMac {
    Write-Log "Configuring network settings..."
    #IMPORT NETWORK SETTINGS FORM CONFIG
    $ScriptDir = Split-Path $psise.CurrentFile.FullPath
    Push-Location $ScriptDir

    #GET NETWORK SETTINGS
    Get-Content 0_NETWORK-SETTINGS.txt | Where {$_ -notmatch '^#.*'} | Foreach-Object{
    $var = $_.Split('=')
    Set-Variable -Name $var[0] -Value $var[1]
    }

    #GET IP ADDRESS FROM CSV VIA HOSTNAME
    $CSV=Import-Csv 0_Table-for-MAC-Hostname-IP.csv -Delimiter ";"
    #$Result = $CSV|?{$_.HostName -match $env:computername}
    #GET ALL MACS FROM ALL NICS
    $Mac=(Get-WmiObject win32_networkadapterconfiguration).macaddress

    #VALIDATE CSV LIST FOR MATCHING MAC
    $Match=@()
    $Mac | 
    ForEach-Object {
        if($CSV.MACAddress -contains $_)
        {
            $Match+=$_
        }
    }

    #SET HOSTNAME FROM CSV
    $Result = $CSV|?{$_.MACAddress -match $Match}
    $IPv4 = $Result.IPAddress


    ##############
    ### SCRIPT ###
    ##############
    ###################### IPv4 KONFIGURATION ####################
    # AKTIVE NETZWERKSCHNITTSTELLE ERMITTELN
    #$adapter = Get-NetAdapter | where {$_.status -eq 'up' -and $_.LinkSpeed -eq '1 Gbps'}| Select-Object –first 1
    $adapter = Get-NetAdapter -Physical

    # AKTUELLE IPV4 ADRESSE ENTFERNEN
    If (($adapter | Get-NetIPConfiguration).IPv4Address.IPAddress) {
    $adapter | Remove-NetIPAddress -AddressFamily IPv4 -Confirm:$false
    }

    #AKTUELLE NETZMASKE ENTFERNEN
    If (($adapter | Get-NetIPConfiguration).Ipv4DefaultGateway) {
    $adapter | Remove-NetRoute -AddressFamily IPv4 -Confirm:$false
    }

    #NETZWERKEINSTELLUNGEN SETZEN
    $adapter | New-NetIPAddress `
    -AddressFamily $IPv4Type `
    -IPAddress $IPv4 `
    -PrefixLength $IPv4Netzmaske `
    -DefaultGateway $IPv4Gateway 

    #DNS ADRESSE SETZEN 
    $adapter | Set-DnsClientServerAddress -ServerAddresses $IPv4DNS 
    $adapter | Set-DnsClientServerAddress -ServerAddresses "192.168.178.254"

    #NETZWERKERKENNUNGSPROFIL FESTLEGEN
    timeout /t 6
    $adapter | set-netconnectionprofile -networkcategory Private
}
Function PreConfig_SystemDrive_CreateDefaultDirs {
    Write-Log "Creating Default Directories..."
    $OrdnerAdmin = ("BGinfo","Scripte","Treiber","Dokumentation","Software","Treiber","Lizenzen","Logo")
    ForEach ($directory in $OrdnerAdmin) {new-item -ItemType directory -Force -Path "C:\ADMIN\$directory" | Out-Null}
    
    #CREATE TEMP DIRECTORY
    If (!(Test-Path "C:\Temp")) {
        New-Item -Path "C:\Temp" -ItemType "directory" -Force | Out-Null
    }
        $global:TempFolder = "C:\Temp"
}

##########################
###      FUNCTIONS     ###
###       CLEANUP      ###
##########################
Function Cleanup_ThirdParty_McAfeeLive {
    Write-Log "Starting McAfee LiveSafe Uninstall Progress..."
    Start-Process "C:\Program Files\McAfee\MSC\mcuihost.exe" -ArgumentList "/body:misp://MSCJsRes.dll::uninstall.html /id:uninstall" -ErrorAction SilentlyContinue
}
Function Cleanup_ThirdParty_MultiAvRemoval {
    Write-Log "Trying to remove multiple 3rdParty AV Solutions..."
    #DOWNLOAD PARAMETER
    $EsetDownloadURL32 = "http://download.eset.com/com/eset/tools/installers/av_remover/latest/avremover_nt32_enu.exe"
    $EsetDownloadURL64 = "http://download.eset.com/com/eset/tools/installers/av_remover/latest/avremover_nt64_enu.exe"
    $EsetFileName = "avremover_ntxx_enu.exe"
    #DOWNLOAD
    if ($global:bittype -eq "64-bit") {Start-BitsTransfer -Source $EsetDownloadURL32 -Destination "C:\Windows\Temp\$EsetFileName"}
    if ($global:bittype -eq "32-bit") {Start-BitsTransfer -Source $EsetDownloadURL64 -Destination "C:\Windows\Temp\$EsetFileName"}
    #INSTALL
    Start-Process "C:\Windows\Temp\$EsetFileName" -ArgumentList "--silent --accepteula --avr-disable" -NoNewWindow -Wait
    #CLEANUP
    Remove-Item "C:\Windows\Temp\$EsetFileName"
}
Function Cleanup_Microsoft_StoreApps {
    Write-Log "Removing Windows Store Bloatware..."

    $WhiteListedApps = @(
    "Microsoft.DesktopAppInstaller",
    "Microsoft.GetHelp"
    "Microsoft.Getstarted"
    "Microsoft.Messaging"
    "Microsoft.MicrosoftEdge"
    "Microsoft.MicrosoftStickyNotes"
    "Microsoft.MSPaint"
    "Microsoft.Print3D"
    "Microsoft.StorePurchaseApp"
    "Microsoft.WebMediaExtensions"
    "Microsoft.Windows.Photos"
    "Microsoft.WindowsAlarms"
    "Microsoft.WindowsCalculator"
    "Microsoft.WindowsCamera" 
    "Microsoft.WindowsCommunicationsApps"
    "Microsoft.WindowsFeedbackHub"
    "Microsoft.WindowsPhotos"
    "Microsoft.WindowsMaps"
    "Microsoft.WindowsSoundRecorder "
    "Microsoft.WindowsStore"
    "Microsoft.YourPhone"
    "RealtekSemiconductorCorp.RealtekAudioControl"
    )
    # Loop through the list of appx packages
    
    $AppArrayList = Get-AppxPackage -PackageTypeFilter Bundle -AllUsers | Select-Object -Property Name, PackageFullName | Sort-Object -Property Name
    foreach ($App in $AppArrayList) {
    # If application name not in appx package white list, remove AppxPackage and AppxProvisioningPackage
    if (($App.Name -in $WhiteListedApps)) {
        Write-Log "Skipping excluded application package: $($App.Name)" 9
        Write-Log -ForegroundColor Gray "Skipping excluded application package: $($App.Name)" 9
    }
    else {
        # Gather package names
        $AppPackageFullName = Get-AppxPackage -Name $App.Name | Select-Object -ExpandProperty PackageFullName
        $AppProvisioningPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $App.Name } | Select-Object -ExpandProperty PackageName
        # Attempt to remove AppxPackage
        if ($AppPackageFullName -ne $null) {
            try {
                #Write-Log "Removing application package: $($App.Name)"
                Write-Log "Removing application package: $($App.Name)"
                Remove-AppxPackage -Package $AppPackageFullName -ErrorAction Stop | Out-Null
            }
            catch [System.Exception] {
                #Write-Log "Removing AppxPackage failed: $($_.Exception.Message)"
                Write-Log "Removing AppxPackage failed: $($_.Exception.Message)" 2
            }
        }
        else {
            #Write-Log "Unable to locate AppxPackage for app: $($App.Name)"
            Write-Log "Unable to locate AppxPackage for app: $($App.Name)" 2
        }
        # Attempt to remove AppxProvisioningPackage
        if ($AppProvisioningPackageName -ne $null) {
            try {
                #Write-Log "Removing application provisioning package: $($AppProvisioningPackageName)"
                Write-Log "Removing application provisioning package: $($AppProvisioningPackageName)"
                Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Online -ErrorAction Stop | Out-Null
            }
            catch [System.Exception] {
                #Write-Log "Removing AppxProvisioningPackage failed: $($_.Exception.Message)"
                Write-Log "Removing AppxProvisioningPackage failed: $($_.Exception.Message)" 2
            }
        }
        else {
            #Write-Log "Unable to locate AppxProvisioningPackage for app: $($App.Name)"
            Write-Log "Unable to locate AppxProvisioningPackage for app: $($App.Name)" 2
        }
    }
}
# White list of Features On Demand V2 packages
# Packages that will be removed: App.Support.QuickAssist, Hello.Face, Hello.Face.Migration, MathRecognizer, OneCoreUAP.OneSync,OpenSSH.Client,XPS.Viewer
Write-Log "Starting Features on Demand V2 removal process"
#$WhiteListOnDemand = "NetFX3|Tools.Graphics.DirectX|Tools.DeveloperMode.Core|Language|Browser.InternetExplorer|Media.WindowsMediaPlayer"
$WhiteListOnDemand = "NetFX3|Tools.Graphics.DirectX|Tools.DeveloperMode.Core|Language|Browser.InternetExplorer|Media.WindowsMediaPlayer|Microsoft.WIndows.WordPad|Microsoft.Windows.Powershell.ISE|Windows.Client.ShellComponents|Hello.Face"
# Get Features On Demand that should be removed
$OnDemandFeatures = Get-WindowsCapability -Online | Where-Object { $_.Name -notmatch $WhiteListOnDemand -and $_.State -like "Installed"} | Select-Object -ExpandProperty Name
foreach ($Feature in $OnDemandFeatures) {
    try {
        Write-Log "Removing Feature on Demand V2 package: $($Feature)"
        Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -like $Feature } | Remove-WindowsCapability -Online -ErrorAction Stop | Out-Null
    }
    catch [System.Exception] {
        Write-Log "Removing Feature on Demand V2 package failed: $($_.Exception.Message)" 2
    }
}

}
Function Cleanup_Microsoft_XPSWriter {
    Write-Log "Removing Windows XPS Writer Printer..."
    $XPSPrinterInstalled = Get-Printer -Name "Microsoft XPS Document Writer" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
    if($XPSPrinterInstalled){
        #Disable-WindowsOptionalFeature -Online -FeatureName "Printing-XPSServices-Features" -NoRestart -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
        Remove-Printer -Name "Microsoft XPS Document Writer" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
    }
}

Function Cleanup_Microsoft_FaxPrinter {
    Write-Log "Removing Windows Fax Printer..."
    Remove-Printer -Name "Fax" -ErrorAction SilentlyContinue
}
Function Cleanup_Microsoft_FaxAndScanServices {
    Write-Log "Disabling Windows Fax and Scan Service..."
    Disable-WindowsOptionalFeature -Online -FeatureName "FaxServicesClientPackage" -NoRestart -WarningAction SilentlyContinue | Out-Null
}
Function Cleanup_BrowserAndDownloadHistory {}
Function Cleanup_RemoteControl_SolarwindsAgent {
    Write-Log "Installing Solarwinds Agent..."

    #DOWNLOAD PARAMETER
    #$SWADownloadURL = "https://xxx.de/dms/FileDownload?customerID=142&softwareID=101"
    $SWAFileName = "WindowsAgentSetup.exe"

    #DOWNLOAD
    Start-BitsTransfer -Source $SWADownloadURL -Destination "$global:TempFolder\$SWAFileName" 

    #INSTALL
    Start-Process "taskkill" -ArgumentList "/IM agent.exe /f" -ErrorAction SilentlyContinue
    Start-Process "$global:TempFolder\$SWAFileName" -ArgumentList "/uninstall" -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$SWAFileName"
}

##########################
###      FUNCTIONS     ###
### DESKTOP APPEARANCE ###
##########################
Function Appearance_CopyOfficeToDesktop {
    Write-Host "Coping Office Shortcuts to Desktop"

    #PARAMETERS
    $OfficePathLinks = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    $OfficeApps = ("Word", "Excel", "PowerPoint", "Outlook", "OneNote")
    
    #COMMAND
    foreach ($App in $OfficeApps){
        write-Host $App
        if(Test-Path "$OfficePathLinks\$App.lnk"){Copy-Item -Path  "$OfficePathLinks\$App.lnk" "C:\ProgramData\Desktop"}
        }
}
############################################################
Function Appearance_ControlPanel_SmallSymbols {
    Write-Log "Setting Control Panel Overview to Small Symbols..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel")) {
	    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" -Name "StartupPage" -Type DWord -Value 1
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" -Name "AllItemsIconView" -Type DWord -Value 1
}
Function Appearance_Desktop_ShowBginfo {
    Write-Log "Installing BGInfo..."
    $BginfoDownloadUrl = "https://download.sysinternals.com/files/BGInfo.zip"
    $BgsettingsDownloadUrl ="https://www.xxx.de/settings.bgi"
    $BGInfoFolder = "C:\ADMIN\BGinfo"
    $BginfoRegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    $BginfoRegkey = 'BGInfo'
    $BginfoRegkeyValue = 'C:\ADMIN\BGInfo\Bginfo.exe C:\ADMIN\BGInfo\settings.bgi /timer:0 /nolicprompt'

    If (!(Test-Path $BGInfoFolder)) {
        New-Item -Path $BGInfoFolder -ItemType "directory" -Force | Out-Null
    }

    #Download
    Start-BitsTransfer -Source $BgsettingsDownloadUrl -Destination "$BGInfoFolder\settings.bgi"
    Start-BitsTransfer -Source $BginfoDownloadUrl     -Destination "$global:TempFolder\bginfo.zip"
    Expand-Archive -LiteralPath "$global:TempFolder\bginfo.zip" -DestinationPath $BGInfoFolder -Force
    Remove-Item "$BGInfoFolder\Eula.txt"
    
    #Autorun
    New-ItemProperty -Path $BginfoRegPath -Name $BginfoRegkey -PropertyType "String" -Value $BginfoRegkeyValue -ErrorAction SilentlyContinue
    Start-Process "$BGInfoFolder\Bginfo.exe" -ArgumentList "$BGInfoFolder\settings.bgi /nolicprompt /timer:0"
    
    #Cleanup
    Remove-Item "$global:TempFolder\bginfo.zip" -Recurse -ErrorAction SilentlyContinue
}

Function Appearance_Desktop_ResetBackgroundToDefault{
    $img = "$env:windir\Web\Wallpaper\Windows\img0.jpg"
    Add-Type -TypeDefinition @" 
using System; 
using System.Runtime.InteropServices;
public class Params
{ 
    [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
    public static extern int SystemParametersInfo (Int32 uAction, 
                                                    Int32 uParam, 
                                                    String lpvParam, 
                                                    Int32 fuWinIni);
}
"@
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02  
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
    [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $img, $fWinIni) | Out-Null
}
Function Appearance_Desktop_ShowComputer {
    Write-Log "Enabling Computer Desktop Symbol..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu")) {
    	New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Type DWord -Value 0
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")) {
    	New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Type DWord -Value 0
}
Function Appearance_Desktop_OpenExplorerWithMyPC {
    Write-Log "Setting Default Explorer View..."
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v LaunchTo /t REG_DWORD /d "00000001" /f
}
Function Appearance_Desktop_EnableDarkTheme {
    Write-Log "Enabling Windows Dark Theme..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Type DWord -Value 0
}
Function Appearance_Desktop_ShowBuildNumber {
    Write-Log "Showing Build Number on Desktop..."
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "PaintDesktopVersion" -Type DWord -Value 1
}
Function Appearance_Desktop_HideBuildNumber {
    Write-Log "Hiding Build Number from Desktop..."
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "PaintDesktopVersion" -Type DWord -Value 0
}
Function Appearance_Edge_DisableShortcutCreation {
    Write-Log "Disabling Auto Shortcut Creation for Microsoft Edge..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "link" -Type Binary -Value ([byte[]](0,0,0,0))
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name "DisableEdgeDesktopShortcutCreation" -Type DWord -Value 1
    }
Function Appearance_Explorer_RenameSystemDrive {
    Write-Log "Renaming C: Drive to System..."
    label C: System 
}
Function Appearance_Explorer_RenameDataDrive {
    Write-Log "Renaming D: Drive to Daten..."
    if (Get-PSDRive D -ErrorAction SilentlyContinue| Where-Object{$_.Provider.name -eq "FileSystem"}) {label D: "Daten"}  
    }
Function Appearance_Explorer_ShowFileExtensions {
    Write-Log "Showing File Extensions for known Filetypes..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Type DWord -Value 0
}
Function Appearance_Explorer_ShowFullPath {
    Write-Log "Settings Explorer to always show full path..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState")) {
	    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Name "FullPath" -Type DWord -Value 1
}
Function Appearance_Explorer_ExpandCurrentFolder {
    Write-Log "Setting Explorer to follow current folder in navigation panel..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "NavPaneExpandToCurrentFolder" -Type DWord -Value 1
}
Function Appearance_Explorer_DisableAeroShake {
    Write-Log "Disabling Aero Shake..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "NoWindowMinimizingShortcuts " -Type DWord -Value 1
    #Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "NoWindowMinimizingShortcuts " -Type DWord -Value 1
}
Function Appearance_PowershellISE_CustomizeView {}
Function Appearance_Taskbar_ChangeSymbols {
    #REMOVES CONTACTS AND TASKVIEW FROM TASKBAR, CHANGES SEARCH BAR TO ICON
    Write-Log "Modifying Taskbar Appearance..."
    REG ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People" /V PeopleBand                           /T REG_DWORD /D 0 /F
    REG ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"        /V ShowTaskViewButton                   /T REG_DWORD /D 0 /F
    REG ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"        /V ShowCortanaButton                    /T REG_DWORD /D 0 /F
    REG ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Search"                   /V SearchboxTaskbarMode                 /T REG_DWORD /D 1 /F
    REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\PenWorkspace"             /V PenWorkspaceButtonDesiredVisibility  /T REG_DWORD /D 0 /F
}
Function Appearance_Taskbar_UnpinUnwantedIcons {
    Write-Log "Removing Unwanted Default Taskbar Icons..."
    #UNPINS ALL UNWANTED ICONS FROM TASKBAR
    #BLACKLIST FOR TASKBAR ICONS TO BE REMOVED
    
    #GET NAMES FOR CURRENT ITEMS
    #((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()) | Select-object Name
    
    #BLACKLIST
    $BlackList=@(
    "Microsoft Store"
    "Lenovo Vantage"
    "Mail"
    "Lenovo Welcome"
    "McAfee LiveSafe"
    "Microsoft Edge"
    )

    #UNPIN APPS (GERMAN)
    $ErrorActionPreference= 'silentlycontinue'
    foreach ($entry in $BlackList){
        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | Where-Object{$_.Name -eq $entry}).Verbs() | Where-Object{$_.Name.replace('&','') -match 'Von Taskleiste lösen'} | ForEach-Object{$_.DoIt(); $exec = $true} 
    }
    $ErrorActionPreference= 'Continue'
    #### DEBUG ###
    <#GET ALL APP NAMES
    #((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items())
    ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()) | Select-object Name
    #GET VERBS FOR APPNAME
    # $appname = "Store"
    # ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs()
    #UNPIN APP (ENGLISH)
    # ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Unpin from taskbar'} | %{$_.DoIt(); $exec = $true}
    #>
}
Function Appearance_Taskbar_Remove2ndKeyboardLayout {
    # REMOVES ADDITONAL SECOND EN-US KEYBOARD LAYOUT
    Write-Log "Removing 2nd Keyboard Layout en-us..."
    try {
        $langs = Get-WinUserLanguageList
        Set-WinUserLanguageList ($langs | Where-Object {$_.LanguageTag -ne "en-US"}) -Force -ErrorAction SilentlyContinue
    }
    catch{
        Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 2
    }
}
Function Appearance_Taskbar_ShowTimeWithSeconds {
    Write-Log "Displaying time with seconds..."

    try {
        If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced")) {
            New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" | Out-Null
        }
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowSecondsInSystemClock" -Type DWord -Value 1
    }
    catch{
        Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 2
    }
}
Function Appearance_Taskbar_ShowDayname {
    Write-Log "Displaying Name of Day in Taskbar..."
    Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name "sShortDate" -Type String -Value "ddd. dd.MM.yyyy"
}
Function Appearance_Taskbar_ShowAllIconsSystray {
    Write-Log "Showing all Icons in Systray..."
    If (!(Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer")) {
	    New-Item -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "EnableAutoTray" -Type DWord -Value 0
}
Function Appearance_TaskManager_CustomSettings {
    Write-Log "Setting Task Manager Custom Appearance..."
    #RUN TASK MANAGER HIDDEN AND GET ITS PROPERTIES
    $taskmgr = Start-Process -WindowStyle Hidden -FilePath taskmgr.exe -PassThru
	    Do {
		Start-Sleep -Milliseconds 100
		$preferences = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\TaskManager" -Name "Preferences" -ErrorAction SilentlyContinue
	} Until ($preferences)
	
    #STOP HIDDEN TASK MANAGER PROCESS
    Stop-Process $taskmgr

    #SETTING CUSTOM SETTINGS
    $preferences.Preferences[28]   = 0 #Change Default View to Detailed
    $preferences.Preferences[3480] = 1 #Change CPU View to Logical Processor
    $preferences.Preferences[4008] = 1 #Show Kernel Usage in CPU View
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\TaskManager" -Name "Preferences" -Type Binary -Value $preferences.Preferences

    #DEBUG
    <#TO GET OUTPUT OF LINE CHANGES
    $taskmgr = Start-Process -WindowStyle Hidden -FilePath taskmgr.exe -PassThru
	Start-Sleep -Milliseconds 100
    Do {$preferences = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\TaskManager" -Name "Preferences" -ErrorAction SilentlyContinue} Until ($preferences)
	Stop-Process $taskmgr
    $preferences.Preferences
    #>
    #COUNTING STARTS WITH 0
    #THERE EXISTS 4824 ENTRIES
    #GET CURRENT SETTINGS COPY THEM TO https://www.diffchecker.com/diff WITH CTRL+A, THEN USE CLS TO CLEAR CONSOLE CHANGE TASK MANAGER, GET RESULTS AGAIN AND PASTE EVERYTHING IN SECOND COLUMN, GET CHANGES
    #OBJECT LINE       DEF   MOD  DESCRIPTION
    #28     (0029)     1     0    Show Task Manager Details 
    #3480   (3481)     1     0    Show Task Manager Details 
}
Function Appearance_System_SetOEMInformations {
    #DOWNLOAD LOGO
    Write-Log "Setting OEM Informations..."
    $OEMURL_Logo = "https://www.xxx.de/OEM-Logo.bmp"
    $OEMLogo_File = "C:\ADMIN\Logo\OEM_Logo.bmp"
    $OEMDirectory = Split-Path -Path "$OEMLogo_File"
    If(!(test-path $OEMDirectory)){New-Item -ItemType Directory -Force -Path $OEMDirectory | Out-Null}
    #$OEMwebclient = New-Object System.Net.WebClient
    #$OEMwebclient.DownloadFile($OEMURL_Logo,$OEMLogo_File)
    # bmp file, 24-bit color, 120x120 or larger will be scaled

    #SET SUPPORT SETTINGS
    $OEMPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
    Set-ItemProperty -Path $OEMPath -Name "HelCustomized" -Type Dword -Value "0"
    Set-ItemProperty -Path $OEMPath -Name "Manufacturer" -Type ExpandString -Value "Your Company Name"
    Set-ItemProperty -Path $OEMPath -Name "Model" -Type ExpandString -Value ""
    #Set-ItemProperty -Path $OEMPath -Name "Logo" -Type ExpandString -Value "C:\ADMIN\LOGO\Logo.bmp"
    Set-ItemProperty -Path $OEMPath -Name "SupportPhone" -Type ExpandString -Value "+49 (xxxx) xxxx-xx"
    Set-ItemProperty -Path $OEMPath -Name "SupportURL" -Type ExpandString -Value "https://www.xxx.de"
    Set-ItemProperty -Path $OEMPath -Name "SupportHours" -Type ExpandString -Value "Montag - Freitag von xx:00 bis xx:00 Uhr"
}

Function Appearance_Account_ProfileLogo {
    #PARAMETER
    $AccountLogoFilePath = "C:\ADMIN\Logo\Account_Logo.jpg"
    $AccountLogoDownloadURL = "https://www.xxx.de/Account_Logo.jpg"
    $DestinationFilePath = "C:\ProgramData\Microsoft\User Account Pictures\user.bmp"

    #DOWNLOAD
    $OEMwebclient = New-Object System.Net.WebClient
    $OEMwebclient.DownloadFile($AccountLogoDownloadURL,$AccountLogoFilePath)

    #KONFIG
    Rename-Item -Path $DestinationFilePath -NewName "user-ori.bmp"
    Copy-Item -Path "$AccountLogoFilePath" -Destination $DestinationFilePath
}
Function Appearance_WinDefender_ShowAlwaysSystray {
    Write-Log "Showing Always Windows Defender Icon in Systray..."
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Systray" -Name "HideSystray" -ErrorAction SilentlyContinue
    If ([System.Environment]::OSVersion.Version.Build -eq 14393) {
    	Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "WindowsDefender" -Type ExpandString -Value "`"%ProgramFiles%\Windows Defender\MSASCuiL.exe`""
    } ElseIf ([System.Environment]::OSVersion.Version.Build -ge 15063 -And [System.Environment]::OSVersion.Version.Build -le 17134) {
    	Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "SecurityHealth" -Type ExpandString -Value "%ProgramFiles%\Windows Defender\MSASCuiL.exe"
    } ElseIf ([System.Environment]::OSVersion.Version.Build -ge 17763) {
    	Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "SecurityHealth" -Type ExpandString -Value "%windir%\system32\SecurityHealthSystray.exe"
    }
}

Function Appearance_Powershell_RunAsAdminContext {
    Write-Log "Adding Powershell ISE as Admin to Contextmenu..."

    try{
    New-PSDrive -Name "HKCR" -PSProvider "Registry" -Root "HKEY_CLASSES_ROOT" | Out-Null
    If (!(Test-Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin")) {
	    New-Item -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin" -Name '(Default)' -Value "ISE als Administrator"
    Set-ItemProperty -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin" -Name "Extended" -Type String -Value "-"
    Set-ItemProperty -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin" -Name "HasLUAShield" -Type String -Value ""
    Set-ItemProperty -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin" -Name "Icon" -Type String -Value "PowerShell_ISE.exe"
    $Value = 'PowerShell -windowstyle hidden -Command "Start-Process cmd -ArgumentList ''/s,/c,start PowerShell_ISE.exe ""%1""''  -Verb RunAs"'
    Set-ItemProperty -Path "HKCR:\Microsoft.PowerShellScript.1\Shell\PowerShellISEAsAdmin\command" -Name '(Default)' -Value $Value
    Remove-PSDrive "HKCR"
    }
    catch{
        Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 2
    }

}

function Appearance_System_DisableMouseShadow {
        #PARAMETERS
        $ActionName = "Appearance Settings - Removing Mouse pointer shadow..."

        #COMMAND #1
        Write-Log $ActionName
        $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
        $RegName = "VisualFXSetting"
        $RegKey = "DWord"
        $RegValue = "3"
    
        If (!(Test-Path $RegPath)){New-Item -Path $RegPath -Force | Out-Null}
        Set-ItemProperty -Path $RegPath -Name $RegName -Type $RegKey -Value $RegValue

        #COMMAND #2
        $RegPath = "HKCU:\Control Panel\Desktop"
        $RegName = "UserPreferencesMask"
        $RegKey = "Binary"
        $RegValue = "([byte[]](0x90, 0x00, 0x30, 0x80, 0x10, 0x00, 0x00, 0x00))"
        #aus = 90 00 30 80 10 00 00 00 
        #ein = 90 20 03 80 10 00 00 00 
    
        If (!(Test-Path $RegPath)){New-Item -Path $RegPath -Force | Out-Null}
        Set-ItemProperty -Path $RegPath -Name $RegName -Type $RegKey -Value $RegValue
}

function Appearance_Taskbar_DisableNewsInterestsWeather {
        #PARAMETERS
        $ActionName = "Appearance Settings - Disabling weather widget..."

        #COMMAND #1
        Write-Log $ActionName
        $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds"
        $RegName = "ShellFeedsTaskbarViewMode"
        $RegKey = "DWord"
        $RegValue = "2"
    
        If (!(Test-Path $RegPath)){New-Item -Path $RegPath -Force | Out-Null}
        Set-ItemProperty -Path $RegPath -Name $RegName -Type $RegKey -Value $RegValue
}
##########################
###      FUNCTIONS     ###
###       CONFIG       ###
##########################
Function Config_System_HostnameSeriennummer{
    #PARAMETERS
    $ActionName = "Setting Hostname if Terra PC"
    $TerraFilePath = "C:\SNr.txt"
    $TerraSerial = if (Test-path $TerraFilePath){Get-Content $TerraFilePath}

    #COMMAND
    Write-Log $ActionName

    #DETERMINE IF DESKTOP OR NOTEBOOK VIA CHASSIS TYPE OR IF BATTERY IS PRESENT
    #if(Get-WmiObject -Class win32_systemenclosure -ComputerName "localhost" | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14}){$Hostname = "NB"+$SnrFilepath}
    if(Get-WmiObject -Class win32_battery -ComputerName "localhost"){
        $Hostname = "NB-"+$TerraSerial
        Write-Log "Setting Hostname to $Hostname" 9
        Rename-Computer -NewName "$Hostname"
    }
    elseif($TerraSerial){
        $Hostname = "AP-"+$TerraSerial
        Write-Log "Setting Hostname to $Hostname" 9
        Rename-Computer -NewName "$Hostname"
    }
    else{
        Write-Log "File for Hostname not Found" 2
    }
    
    
}
Function Config_System_DisableUAC{
    #PARAMETERS
    $ActionName = "Disabling UAC"

    #COMMAND
    Write-Log $ActionName
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorUser" -Value "0"
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value "1"
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "PromptOnSecureDesktop" -Value "0"
}

Function Config_Privacy_AllowMicrophoneAccess{
    #PARAMETERS
    $ActionName = "Privacy Settings - Allowing Apps to access microphone"

    #COMMAND
    Write-Log $ActionName
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
    $RegName = "Value"
    $RegKey = "String"
    $RegValue = "Allow"

    If (!(Test-Path $RegPath)){New-Item -Path $RegPath -Force | Out-Null}
    Set-ItemProperty -Path $RegPath -Name $RegName -Type $RegKey -Value $RegValue
}
Function Config_System_DisableQuickBoot{
    Write-Host "Disabling QuickBoot"
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Type DWord -Value 0
}
Function Config_Applications_Defaults{}
Function Config_Power_PowerPlanHigh{
    Write-Log "Setting Powerplan to high power" 0
    try{
        Start-Process "powercfg" -ArgumentList '-setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c' -NoNewWindow -Wait
    }
    catch{
        Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 2
        #Write-Log ("{0} - $_.ExitCode" -f $MyInvocation.MyCommand) 2
    }
    finally{
        #Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 9
        #Write-Log ("{0} - $_.ExitCode" -f $MyInvocation.MyCommand) 9 
        #Write-Log ($_.Exception.Message) 9
        #Write-Log ($_.Exception.ItemName) 9
    }
    
}

Function Config_Power_DisableFastboot{
    Write-Log "Disabling Fastboot" 0
    try{
        Start-Process "powercfg" -ArgumentList '/hibernate off' -NoNewWindow -Wait
    }
    catch{
        Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 2
        #Write-Log ("{0} - $_.ExitCode" -f $MyInvocation.MyCommand) 2
    }
    finally{
        #Write-Log ("{0} - $_" -f $MyInvocation.MyCommand) 9
        #Write-Log ("{0} - $_.ExitCode" -f $MyInvocation.MyCommand) 9 
        #Write-Log ($_.Exception.Message) 9
        #Write-Log ($_.Exception.ItemName) 9
    }
    
}

###############################################
Function Config_Explorer_DisableThumbsdbOnNetwork {
    Write-Log "Disabling Thumbs.db on Network Drives..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "DisableThumbsDBOnNetworkFolders" -Type DWord -Value 1
}
Function Config_Explorer_AddPhotoViewerOpenWith {
    Write-Log "Adding PhotoViewer to open with..."
    If (!(Test-Path "HKCR:")) {
	    New-PSDrive -Name "HKCR" -PSProvider "Registry" -Root "HKEY_CLASSES_ROOT" | Out-Null
    }
    New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\command" -Force | Out-Null
    New-Item -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Force | Out-Null
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open" -Name "MuiVerb" -Type String -Value "@photoviewer.dll,-3043"
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\command" -Name "(Default)" -Type ExpandString -Value "%SystemRoot%\System32\rundll32.exe `"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll`", ImageView_Fullscreen %1"
    Set-ItemProperty -Path "HKCR:\Applications\photoviewer.dll\shell\open\DropTarget" -Name "Clsid" -Type String -Value "{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}"
}
Function Config_IE_DisableFirstrunWizard {
    Write-Log "Disabling First Run Wizard for Internet Explorer..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Type DWord -Value 1
}
Function Config_System_EnableClipboardHistory {
    Write-Log "Enabling Clipboard History..."
    $Path = "HKCU:\Software\Microsoft\Clipboard"
    If (!(Test-Path $Path)) {
	    New-Item -Path $Path -Force | Out-Null
    }
    Set-ItemProperty -Path $Path -Name "EnableClipboardHistory" -Type DWord -Value 1
}
Function Config_System_EnableStorageSense {
    Write-Log "Enabling Windows StorageSense..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy")) {
	    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" -Name "01" -Type DWord -Value 1
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" -Name "StoragePoliciesNotified" -Type DWord -Value 1
}
Function Config_System_DisableRebootOnBluescreen {
    Write-Log "Disabling Reboot on Bluescreen..."
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -Name "AutoReboot" -Type DWord -Value 0
}
Function Config_System_ShowShutdownOptionsLockscreen {
    Write-Log "Showing Shutdown Options on Lockscreen..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ShutdownWithoutLogon" -Type DWord -Value 1
}
Function Config_System_ShowFileOperationsDetails {
    Write-Log "Setting Explorer to show File Operation Details..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\OperationStatusManager")) {
	    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\OperationStatusManager" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\OperationStatusManager" -Name "EnthusiastMode" -Type DWord -Value 1
}
Function Config_System_DisableSearchUnknownExtensions {
    Write-Log "Disabling search for unknown extensions..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Name "NoUseStoreOpenWith" -Type DWord -Value 1
}
Function Config_System_EnableNumlockOnStartup {
    Write-Log "Enabling NUMLOCK on Startup..."
    If (!(Test-Path "HKU:")) {
	    New-PSDrive -Name "HKU" -PSProvider "Registry" -Root "HKEY_USERS" | Out-Null
    }
    Set-ItemProperty -Path "HKU:\.DEFAULT\Control Panel\Keyboard" -Name "InitialKeyboardIndicators" -Type DWord -Value 2147483650
    Add-Type -AssemblyName System.Windows.Forms
    If (!([System.Windows.Forms.Control]::IsKeyLocked('NumLock'))) {
	$wsh = New-Object -ComObject WScript.Shell
	$wsh.SendKeys('{NUMLOCK}')
    }
}
Function Config_Network_EnableLinkedConnections {
    Write-Log "Allowing Administrative Users to Access Network Shares..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLinkedConnections" -Type DWord -Value 1
}
Function Config_Network_SetUnknownNetworksPrivate {
    Write-Log "Setting Unknown Networks to Private..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\010103000F0000F0010000000F0000F0C967A3643C3AD745950DA7859209176EF5B87C875FA20DF21951640E807D7C24")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\010103000F0000F0010000000F0000F0C967A3643C3AD745950DA7859209176EF5B87C875FA20DF21951640E807D7C24" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\010103000F0000F0010000000F0000F0C967A3643C3AD745950DA7859209176EF5B87C875FA20DF21951640E807D7C24" -Name "Category" -Type DWord -Value 1
}
Function Config_Firewall_AllowPing {
    Write-Log "Allowing Response to Ping Requests..."
    netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol="icmpv4:8,any" dir=in action=allow | Out-Null
    netsh advfirewall firewall add rule name="ICMP Allow incoming V6 echo request" protocol="icmpv6:8,any" dir=in action=allow | Out-Null
}
Function Config_Power_DisableStandbyOnAC{
    powercfg.exe -x -standby-timeout-ac 0 | Out-Null
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
}
Function Config_Power_DisableStandbyOnBattery{
    powercfg.exe -x -standby-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
}
Function Config_Power_DisableMonitorTimeoutAC{
    powercfg.exe -x -monitor-timeout-ac 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
}
Function Config_Power_DisableMonitorTimeoutOnBattery{
    powercfg.exe -x -monitor-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
}
Function Config_Power_DisableAllStandbyOptions{
    powercfg.exe -x -monitor-timeout-ac 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -monitor-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -disk-timeout-ac 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -disk-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -standby-timeout-ac 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -standby-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -hibernate-timeout-ac 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    powercfg.exe -x -hibernate-timeout-dc 0
    if( -not $? ){$FunctionMame = $MyInvocation.MyCommand[0];$ErorMesgCMD = $Error[0].Exception.Message;Write-Log "$FunctionMame failed $ErorMesgCMD" 2}
    }
Function Config_Remote_AllowRdpWithBlankPassword {
    Write-Log "Allowing Remote Connections with blank password..."
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Type DWORD -Value "0"
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name "LimitBlankPasswordUse" -Type DWORD -Value "0"
}
Function Config_Network_AddPrivateNetworksIntranet{
    Write-Log "Adding multiple sites to intranet and trusted sites..."
    $TrustedSites = (127.0.0.1<#,"xxx.de","xxx.support"#>)

    #DETERMINE LOCAL FQDN 
    $RangeLocalDomain = "*.$env:userdnsdomain"

    #DETERMINE HOSTNAME DC / FILESERVER
    $HostnameFileserver = Get-SmbConnection | Select-Object -ExpandProperty ServerName | Sort-Object -Property @{Expression={$_.Trim()}} -Unique

    #DETERMINE EXTERNAL DOMAIN
    #$ExternalDomain = "mail.hilgefort.de"
    #$ExternalIP = Invoke-RestMethod 'http://ipinfo.io/json' | Select -exp ip
    #$ExternalDomain = ""

    #DEFINE COMMON INTRANET ADRESSSES
    $IntranetRanges = (
    "localhost",
    "192.168.*.*",
    "172.16.*.*",
    "127.0.0.*",
    "10.*.*.*")

    $count = 0

    foreach ($Site in $TrustedSites){
        if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$Site")){$null = New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$Site"}
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$Site" -Name "HTTP" -Value 2 -Type DWord
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\$Site" -Name "HTTPS" -Value 2 -Type DWord
        }

    foreach ($Range in $IntranetRanges){
        $count++
        if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count")){$null = New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count"}
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name '*' -Value 1 -Type DWord
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name ':Range' -Value "$Range" -Type Multistring 
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name 'file' -Value "1" -Type DWord
        }

    #SET INTRANET FQDN IF AVAILABLE
    if ($RangeLocalDomain){
        foreach ($Range in $RangeLocalDomain){
            $count++
            if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count")){$null = New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count"}
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name '*' -Value 1 -Type DWord
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name ':Range' -Value "$Range" -Type Multistring 
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name 'file' -Value "1" -Type DWord
        }
    }

    #SET FILESERVER IF AVAILABLE
    if ($HostnameFileserver){
        foreach ($FileHost in $HostnameFileserver){
            $count++
            if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count")){$null = New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count"}
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name '*' -Value 1 -Type DWord
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name ':Range' -Value "$FileHost" -Type Multistring 
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range$count" -Name 'file' -Value "1" -Type DWord
        }
    }
}
Function Config_WinDefender_DisableMsAccountWarning {
    Write-Log "Disabling Defender Warnings for not having a Microsoft Account..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows Security Health\State")) {
	New-Item -Path "HKCU:\Software\Microsoft\Windows Security Health\State" -Force | Out-Null
    }
    Set-ItemProperty "HKCU:\Software\Microsoft\Windows Security Health\State" -Name "AccountProtection_MicrosoftAccount_Disconnected" -Type DWord -Value 1
}
Function Config_WinDefender_DisableOneDriveWarning {
    Write-Log "Disabling Windows Defender Warning if OneDrive is not configured..."
     
    Write-Log "Disabling Windows Defender Cloud..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SpynetReporting" -Type DWord -Value 0
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SubmitSamplesConsent" -Type DWord -Value 2
    
    # DEFENDER EXPLOIT GUARD
    #if ($Global:WindowsVersion -ne "6.3") { throw "Exploit-Guard Configuration only supported on Windows 10, but Version is $WindowsVersion" }
    #if ($Global:WindowsRelease -lt 1709)  { throw "Exploit-Guard Configuration only supported on Windows 10 Release 1709 and higher, but Release is $WindowsRelease" }
    #If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows Defender\Exploit Guard")) {
	#    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows Defender\Exploit Guard" -Force | Out-Null
    #}
    #If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows Defender\Exploit Guard\Controlled Folder Access")) {
    #	New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows Defender\Exploit Guard\Controlled Folder Access" -Force | Out-Null
    #}
    #Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows Defender\Exploit Guard\Controlled Folder Access" -Name "GuardMyFolders" -Type DWord -Value 1
    
    #DISABLE DEFENDER REPORTING (SCAN RESULTS, SUMMARY AND ONEDRIVE WARNING)
    #Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableEnhancedNotifications" -Type DWord -Value 1

}
Function Config_WinUpdates_DisabledDriverUpdates {
    Write-Log "Disabling DriverUpdate via Windows Update..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Device Metadata")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Device Metadata" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Device Metadata" -Name "PreventDeviceMetadataFromNetwork" -Type DWord -Value 1
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontPromptForWindowsUpdate" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontSearchWindowsUpdate" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DriverUpdateWizardWuSearchEnabled" -Type DWord -Value 0
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate")) {
    	New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "ExcludeWUDriversInQualityUpdate" -Type DWord -Value 1
}
Function Config_WinUpdates_EnableMicrosoftUpdate {
    Write-Log "Allowing Windows Update to provide Updates for other Microsoft Products..."
    (New-Object -ComObject Microsoft.Update.ServiceManager).AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
}
Function Config_WinUpdates_DisableNightlyWakeUp {
    Write-Log "Disabling nightly wakeup for maintenance tasks..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUPowerManagement" -Type DWord -Value 0
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\Maintenance" -Name "WakeUp" -Type DWord -Value 0
}
Function Config_WinUpdates_EnableRestartSignOn {
    Write-Log "Enabling SingOn that is needed to complete Updates on Restart..."
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "DisableAutomaticRestartSignOn" -ErrorAction SilentlyContinue
}
Function Config_WinUpdates_SystemDriveAsTempDir {
    Write-Log "Setting the system drive as only allowed Temp Drive for MSI Installer..."
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    Set-ItemProperty -Path "HKCR:\Msi.Package\shell\Open\Command"      -Name '(Default)' -Type "ExpandString" -Value '"%SystemRoot%\System32\msiexec.exe" /i "%1" ROOTDRIVE=C:\ %*'
    Set-ItemProperty -Path "HKCR:\Msi.Package\shell\Repair\Command"    -Name '(Default)' -Type "ExpandString" -Value '"%SystemRoot%\System32\msiexec.exe" /f "%1" ROOTDRIVE=C:\ %*'
    Set-ItemProperty -Path "HKCR:\Msi.Package\shell\Uninstall\Command" -Name '(Default)' -Type "ExpandString" -Value '"%SystemRoot%\System32\msiexec.exe" /x "%1" ROOTDRIVE=C:\ %*'
    Set-ItemProperty -Path "HKCR:\Msi.Patch\shell\Open\Command"        -Name '(Default)' -Type "ExpandString" -Value '"%SystemRoot%\System32\msiexec.exe" /p "%1" ROOTDRIVE=C:\ %*'
    Remove-PSdrive -Name HKCR
}
Function Config_OneDrive_DisableAutorun {
    Write-Log "Disabling OneDrive..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive")) {
	New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Type DWord -Value 1
    
    #BUG FIXING, Die Eigenschaft OneDrive ist im Pfad HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run nicht vorhanden.
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -ErrorAction SilentlyContinue
    
    Start-process "taskkill" -ArgumentList "/f /im onedrive.exe" -ErrorAction SilentlyContinue
}

##########################
###      FUNCTIONS     ###
###        FIXES       ###
##########################
Function Fix_Powershell_IseFilenameComma{
    Set-ItemProperty HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Edit\Command '(default)' '"C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe" """%1"""'
}

##########################
###      FUNCTIONS     ###
###      HARDENING     ###
##########################
Function Hardening_Network_DisableWiFiSense {
    Write-Log "Disabling WiFi Sense..."
If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting")) {
New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" -Force | Out-Null
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting" -Name "Value" -Type DWord -Value 0
If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots")) {
New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots" -Force | Out-Null
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots" -Name "Value" -Type DWord -Value 0
If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config")) {
New-Item -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Force | Out-Null
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Name "AutoConnectAllowedOEM" -Type DWord -Value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" -Name "WiFISenseAllowed" -Type DWord -Value 0
}
Function Hardening_Store_DisableSuggestions {
    Write-Log "Disabling Application suggestions..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "ContentDeliveryAllowed" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "OemPreInstalledAppsEnabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "PreInstalledAppsEnabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "PreInstalledAppsEverEnabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SilentInstalledAppsEnabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-310093Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-314559Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338387Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338388Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338389Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338393Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-353694Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-353696Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-353698Enabled" -Type DWord -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SystemPaneSuggestionsEnabled" -Type DWord -Value 0
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent")) {
	    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Type DWord -Value 1
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace")) {
    	New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace" -Name "AllowSuggestedAppsInWindowsInkWorkspace" -Type DWord -Value 0
    # Empty placeholder tile collection in registry cache and restart Start Menu process to reload the cache
    # Seems not to work anymore in Win10 20H2, coz the path wasnt found
    If ([System.Environment]::OSVersion.Version.Build -ge 17134) {
        $key = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount\*windows.data.placeholdertilecollection\Current"
	    Set-ItemProperty -Path $key.PSPath -Name "Data" -Type Binary -Value $key.Data[0..15]
	    Stop-Process -Name "ShellExperienceHost" -Force -ErrorAction SilentlyContinue
    }
}
Function Hardening_System_DisableAutoplay {
    Write-Log "Disabling AutoPlay Function for removable Devices..."
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers" -Name "DisableAutoplay" -Type DWord -Value 1
}
Function Hardening_System_DisableAutorun {
    Write-Log "Disabling Autorun for optical Drives..."
    If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer")) {
	    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "NoDriveTypeAutoRun" -Type DWord -Value 255
}
Function Hardening_System_DisableTailoredExperience {
    Write-Log "Disabling Tailored Experiences..."
    If (!(Test-Path "HKCU:\Software\Policies\Microsoft\Windows\CloudContent")) {
	    New-Item -Path "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" -Name "DisableTailoredExperiencesWithDiagnosticData" -Type DWord -Value 1
}
Function Hardening_System_DisableAdvertisingID {
    Write-Log "Disabling Advertising ID..."
    Write-Log "Disabling Advertising ID..."
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo")) {
    	New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" -Name "DisabledByGroupPolicy" -Type DWord -Value 1
}
Function Hardening_System_DisallowLanguage {
    Write-Log "Disabling Website Access to Language List..."
    Set-ItemProperty -Path "HKCU:\Control Panel\International\User Profile" -Name "HttpAcceptLanguageOptOut" -Type DWord -Value 1
}
Function Hardening_System_DisableErrorReporting {
    Write-Log "Disabling Error Reporting..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" -Name "Disabled" -Type DWord -Value 1
    Disable-ScheduledTask -TaskName "Microsoft\Windows\Windows Error Reporting\QueueReporting" | Out-Null
}
Function Hardening_System_DisableUserExpAndTelemtry {
    Write-Log "Stopping and disabling Connected User Experiences and Telemetry Service..."
    Stop-Service "DiagTrack" -WarningAction SilentlyContinue
    Set-Service "DiagTrack" -StartupType Disabled
}
Function Hardening_System_DisableSharedExperience {
    Write-Log "Disabling Shared Experience..."
    If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\CDP")) {
	    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\CDP" | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\CDP" -Name "RomeSdkChannelUserAuthzPolicy" -Type DWord -Value 0
}
Function Hardening_System_DisableWindowsScriptHost {
    Write-Log "Disabling Windows Script Host..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings" -Name "Enabled" -Type DWord -Value 0
}
Function Hardening_Defender_EnableSmartScreen {
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Type String -Value "RequireAdmin" -ErrorAction SilentlyContinue
}
Function Hardening_Defender_EnableNotificationBlocks {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile" -Name "DisableNotifications" -Type DWORD -Value "0" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile"   -Name "DisableNotifications" -Type DWORD -Value "0" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile"   -Name "DisableNotifications" -Type DWORD -Value "0" -ErrorAction SilentlyContinue
}

Function Hardening_CMD_DisableQuickEdit {
    Set-ItemProperty -Path "HKCU:\Console" -Name "QuickEdit" -Type DWORD -Value "0" -ErrorAction SilentlyContinue
    }

##########################
###      FUNCTIONS     ###
###    INSTALLATIONS   ###
##########################
Function Install_RemoteControl_TightVNC{
    #PARAMETERS
    $AppName = "TightVNC"
    $DownloadUrl_64 = "https://www.tightvnc.com/download/2.8.27/tightvnc-2.8.27-gpl-setup-64bit.msi"
    $DownloadUrl_32 = "https://www.tightvnc.com/download/2.8.27/tightvnc-2.8.27-gpl-setup-32bit.msi"
    $Filename_64 = "tightvnc-2.8.27-gpl-setup-64bit.msi"
    $Filename_32 = "tightvnc-2.8.27-gpl-setup-32bit.msi"
    
    #DOWNLOAD
    Write-Log "Downloading $AppName..."
    Push-Location $global:TempFolder
    
    if ($global:bittype -eq "64-Bit") {
        #Invoke-WebRequest -Uri $DownloadUrl_64  -OutFile "$global:TempFolder\$Filename_64"
        $path = "$global:TempFolder\$Filename_64"
        #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
        $webClient = new-object System.Net.WebClient
        $webClient.DownloadFile( $DownloadUrl_64, $path )
    }
    if ($global:bittype -eq "32-Bit") {
        #Start-BitsTransfer -Source $DownloadUrl_32
        $path = "$global:TempFolder\$Filename_32"
        #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
        $webClient = new-object System.Net.WebClient
        $webClient.DownloadFile( $DownloadUrl_32, $path )
    }

    #EXTRACT
    #if ($global:bittype -eq "64-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_64" -DestinationPath $global:TempFolder -Force}
    #if ($global:bittype -eq "32-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_32" -DestinationPath $global:TempFolder -Force}

    #INSTALL MSI
    Write-Log "Installing $AppName..."
    $password = "YourVncPassword"
    $adminpassword = "YourVncAdminPassword"
    if ($global:bittype -eq "64-Bit") {Start-Process "msiexec.exe " -ArgumentList "/i $global:TempFolder\tightvnc-2.8.27-gpl-setup-64bit.msi /quiet /norestart SET_USEVNCAUTHENTICATION=1 VALUE_OF_USEVNCAUTHENTICATION=1 SET_PASSWORD=1 VALUE_OF_PASSWORD=$password SET_USECONTROLAUTHENTICATION=1 VALUE_OF_USECONTROLAUTHENTICATION=1 SET_CONTROLPASSWORD=1 VALUE_OF_CONTROLPASSWORD=$adminpassword SET_VIEWONLYPASSWORD=1 VALUE_OF_VIEWONLYPASSWORD==$password" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit") {Start-Process "msiexec.exe " -ArgumentList "/i $global:TempFolder\$Filename_32 /quiet /norestart SET_USEVNCAUTHENTICATION=1 VALUE_OF_USEVNCAUTHENTICATION=1 SET_PASSWORD=1 VALUE_OF_PASSWORD=$password SET_USECONTROLAUTHENTICATION=1 VALUE_OF_USECONTROLAUTHENTICATION=1 SET_CONTROLPASSWORD=1 VALUE_OF_CONTROLPASSWORD=$adminpassword SET_VIEWONLYPASSWORD=1 VALUE_OF_VIEWONLYPASSWORD==$password" -NoNewWindow -Wait}
    Pop-Location

    #CONFIG
    #FIREWALL RULE
    New-NetFirewallRule -DisplayName "TightVNC" -Direction Inbound –Protocol TCP –LocalPort 5900 -Action allow | Out-Null
   

    #CLEANUP
    Start-Sleep -s 2
    if ($global:bittype -eq "64-Bit") {Remove-Item "$global:TempFolder\$Filename_64"}
    if ($global:bittype -eq "32-Bit") {Remove-Item "$global:TempFolder\$Filename_32"}
    }
Function Install_RemoteControl_PCVisit{
    #PARAMETERS
    $Global:ApplicationName = "PCVisit"
    $Global:DownloadURL64 = 'https://nacl.pcvisit.com/fast_update/v1/hosted/jumplink?&langid=de-DE&func=download&productid=18&mode=release&os=osWin32&productRole=remoteHostSetup&fileending=zip&os=osWin32'
    $Global:DownloadURL32 = 'https://nacl.pcvisit.com/fast_update/v1/hosted/jumplink?&langid=de-DE&func=download&productid=18&mode=release&os=osWin32&productRole=remoteHostSetup&fileending=zip&os=osWin32'
    $Filename = "pcvisit_RemoteHost_Setup.zip"

    #GLOBAL
    Write-Log "Installing $Global:ApplicationName..."
    Set_GlobalTempDir

    #RESOLVE
    Resolve-DownloadURL

    #DOWNLOAD
    #if($Global:DowloadFinalURL){Get-FileWebclient $Global:DowloadFinalURL }#-includeStats}
    <#
    #Powershell 7 and newer
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    Write-Log "Downloading $ApplicationName..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -UseBasicParsing -Uri $Global:DowloadFinalURL -OutFile "$global:TempFolder\$FileName" #-SkipCertificateCheck 
    $ProgressPreference = 'Continue'
    #>
    Write-Log "Downloading $Global:ApplicationName..." 9
    $path = "$global:TempFolder\$Filename"

    
    $webClientPCV = new-object System.Net.WebClient
    $webClientPCV.DownloadFile( $Global:DowloadFinalURL, $path )
    #}



    #EXTRACT
    Extract-Archive 

    #INSTALL
    if($Global:DowloadFinalURL -AND $global:ExeFileName){
        Write-Log "Installing $ApplicationName..." 9
        if (Test-Path "$global:TempFolder\$global:ExeFileName"){
            $InstallFile = Join-Path "$global:TempFolder\" "$global:ExeFileName"
            Start-Process $InstallFile -ArgumentList "/S" -NoNewWindow -Wait
        }
    }
    Pop-Location

    #CONFIG
    #RENAME SHORTCUT
    #$SourceFilename = Get-ChildItem "$env:Public\Desktop" | Where-Object { $_.Name -match 'PCvisit' }
    #$FinalFilename = "Fernwartung"
    #Rename-Item -Path $SourceFilename.Fullname -NewName "$FinalFilename -PCVisit-.lnk"
        
    #CREATE DESKTOP SHORTCUT
    $ShortcutIconFile = 'C:\Program Files (x86)\pcvisit Software AG\pcvisit RemoteHost\client.exe'
    $ShortcutFile = 'C:\Program Files (x86)\pcvisit Software AG\pcvisit RemoteHost\PCVisit_client.exe'
    if (Test-Path $ShortcutFile){
        $ShortcutDestination = "$env:Public\Desktop\Fernwartung -PCVisit-.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutDestination)
        $Shortcut.TargetPath = $ShortcutFile
        $Shortcut.IconLocation = $ShortcutIconFile
        $Shortcut.Save()
    }

    #CLEANUP
    Clear-TempDir
    }   

Function Install_RemoteControl_Teamviewer{
    #PARAMETERS
    $Global:ApplicationName = "TeamViewer"
    $Global:DownloadURL64 = 'https://dl.teamviewer.com/download/version_15x/TeamViewer_Setup.exe'
    $Global:DownloadURL32 = 'https://dl.teamviewer.com/download/version_15x/TeamViewer_Setup.exe'
    
    #GLOBAL
    Set_GlobalTempDir

    #RESOLVE
    Resolve-DownloadURL

    #DOWNLOAD
    #if($Global:DowloadFinalURL){Get-FileWebclient $Global:DowloadFinalURL -includeStats}
    Write-Log "Downloading Teamviewer..."
    $path = "$global:TempFolder\TeamViewer_Setup.exe"

    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $Global:DowloadFinalURL, $path )

    #EXTRACT
    Extract-Archive 

    #INSTALL
    if($Global:DowloadFinalURL -AND $global:ExeFileName){
        Write-Log "Installing $ApplicationName..."
        if (Test-Path "$global:TempFolder\$global:ExeFileName"){
            $InstallFile = Join-Path "$global:TempFolder\" "$global:ExeFileName"
            Start-Process $InstallFile -ArgumentList "/S" -NoNewWindow -Wait
        }
    }
    Pop-Location

    #CONFIG    
    $SourceFilenameTV = Get-ChildItem "$env:Public\Desktop" | Where-Object { $_.Name -match 'Teamviewer' }
    $FinalFilenameTV = "Fernwartung -Teamviewer-"
    Rename-Item -Path $SourceFilenameTV.Fullname -NewName "$FinalFilenameTV.lnk"

    #CLEANUP
    Clear-TempDir
    }  

<#
Function Install_RemoteControl_Teamviewer{
    #PARAMETERS
    $AppName = "Teamviewer"
    $DownloadUrl_64 = "https://dl.teamviewer.com/download/version_15x/TeamViewer_Setup.exe"
    $DownloadUrl_32 = "https://dl.teamviewer.com/download/version_15x/TeamViewer_Setup.exe"
    $Filename_64 = "TeamViewer_Setup.exe"
    $Filename_32 = "TeamViewer_Setup.exe"
    
    #DOWNLOAD
    Write-Log "Downloading $AppName..."
    Push-Location $global:TempFolder
    if ($global:bittype -eq "64-Bit") {Start-BitsTransfer -Source $DownloadUrl_64}
    if ($global:bittype -eq "32-Bit") {Start-BitsTransfer -Source $DownloadUrl_32}
    Pop-Location

    #EXTRACT
    #if ($global:bittype -eq "64-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_64" -DestinationPath $global:TempFolder -Force}
    #if ($global:bittype -eq "32-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_32" -DestinationPath $global:TempFolder -Force}

    #INSTALL
    Write-Log "Installing $AppName..."
    if ($global:bittype -eq "64-Bit") {Start-Process "$global:TempFolder\$Filename_64" -ArgumentList "/S /norestart" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit") {Start-Process "$global:TempFolder\$Filename_32" -ArgumentList "/S /norestart" -NoNewWindow -Wait}

    #CONFIG
    #RENAME SHORTCUT
    $SourceFilenameTV = Get-ChildItem "$env:Public\Desktop" | Where-Object { $_.Name -match 'Teamviewer' }
    $FinalFilenameTV = "Fernwartung -Teamviewer-"
    Rename-Item -Path $SourceFilenameTV.Fullname -NewName "$FinalFilenameTV.lnk"
   
    #CLEANUP
    Start-Sleep -s 5
    if ($global:bittype -eq "64-Bit") {Remove-Item "$global:TempFolder\$Filename_64"}
    if ($global:bittype -eq "32-Bit") {Remove-Item "$global:TempFolder\$Filename_32"}
}
#>
Function Install_AntiVirus_EsetESMC{}
Function Install_Tool_TotalCommander{}
Function Install_Backup_StorageCraft{}
Function Install_Browser_Chrome{
    
    #DOWNLOAD PARAMETER
    $Appname = "Google Chrome"
    $DownloadURL = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B16A784A5-15EE-2AC6-9A5A-D78D267EEC4E%7D%26lang%3Dde%26browser%3D5%26usagestats%3D0%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"
    $Filename = "ChromeSetup.exe"
    
    #DOWNLOAD
    Write-Log "Installing $Appname..."
    #Invoke-WebRequest -Uri $DownloadURL  -OutFile "$global:TempFolder\$Filename" 

    $path = "$global:TempFolder\$Filename"

    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURL, $path )

    #INSTALL
    #Write-Log "Download $Appname..."
    Start-Process "$global:TempFolder\$Filename" -NoNewWindow -Wait

    #CLEANUP
    Start-Process "taskkill" -ArgumentList "/IM Chrome.exe /f" -ErrorAction SilentlyContinue
    Remove-Item "$global:TempFolder\$Filename"
}
Function Install_Runtimes_Java{
    
    #DOWNLOAD PARAMETER
    $Appname = "Java AdoptOpenJRE 64bit und 32bit, V11 LTS"
    Write-Log "Installing $Appname..."
    $DownloadURL = "https://ninite.com/adoptjava8-adoptjavax8/ninite.exe"
    $DownloadURLJava64 = "https://github.com/AdoptOpenJDK/openjdk11-binaries/releases/download/jdk-11.0.8%2B10/OpenJDK11U-jre_x64_windows_hotspot_11.0.8_10.msi"
    $DownloadURLJava32 = "https://github.com/AdoptOpenJDK/openjdk11-binaries/releases/download/jdk-11.0.8%2B10/OpenJDK11U-jre_x86-32_windows_hotspot_11.0.8_10.msi"
    $FilenameJava64 = "OpenJDK11U-jre_x64_windows_hotspot_11.0.8_10.msi"
    $FilenameJava32 = "OpenJDK11U-jre_x86-32_windows_hotspot_11.0.8_10.msi"
    
    #DOWNLOAD
    Write-Log "Downloading Java AdoptOpenJRE 64bit..."
    #Invoke-WebRequest -Uri $DownloadURLJava64  -OutFile "$global:TempFolder\$FilenameJava64" 
    $path = "$global:TempFolder\$FilenameJava64"

    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURLJava64, $path )

    Write-Log "Downloading Java AdoptOpenJRE 32bit..."
    #nvoke-WebRequest -Uri $DownloadURLJava32  -OutFile "$global:TempFolder\$FilenameJava32" 
    $path = "$global:TempFolder\$FilenameJava32"
    #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURLJava32, $path )

    #INSTALL
    Write-Log "Installing Java AdoptOpenJRE 64bit......"
    Start-Process "msiexec" -ArgumentList "/i $global:TempFolder\$FilenameJava64 INSTALLLEVEL=1 /quiet" -NoNewWindow -Wait
    Write-Log "Installing Java AdoptOpenJRE 32bit......"
    Start-Process "msiexec" -ArgumentList "/i $global:TempFolder\$FilenameJava32 INSTALLLEVEL=1 /quiet" -NoNewWindow -Wait
    

    #CLEANUP
    Remove-Item "$global:TempFolder\$FilenameJava64"
    Remove-Item "$global:TempFolder\$FilenameJava32"
}
Function Install_Runtimes_Silverlight{
    
    #DOWNLOAD PARAMETER
    $Appname = "Microsoft Silverlight"
    $DownloadURL = "https://download.microsoft.com/download/D/D/F/DDF23DF4-0186-495D-AA35-C93569204409/50918.00/Silverlight_x64.exe"
    $Filename = "Silverlight_x64.exe"
    
    #DOWNLOAD
    Write-Log "Downloading $Appname..."
    #Invoke-WebRequest -Uri $DownloadURL -OutFile "$global:TempFolder\$Filename" 
    $path = "$global:TempFolder\$Filename"
    #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURL, $path )

    #INSTALL
    Write-Log "Installing $Appname..."
    Start-Process "$global:TempFolder\$Filename" -ArgumentList "/q /doNotRequireDRMPrompt " -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$Filename"
}
Function Install_Runtimes_Air{
    
    #DOWNLOAD PARAMETER
    $Appname = "Adobe Air"
    $DownloadURL = "https://airdownload.adobe.com/air/win/download/32.0/AdobeAIRInstaller.exe"
    $Filename = "AobeAIRInstaller.exe"
    
    #DOWNLOAD
    Write-Log "Downloading $Appname..."
    #Invoke-WebRequest -Uri $DownloadURL  -OutFile "$global:TempFolder\$Filename" 
    $path = "$global:TempFolder\$Filename"
    #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURL, $path )

    #INSTALL
    Write-Log "Installing $Appname..."
    Start-Process "$global:TempFolder\$Filename" -ArgumentList "-silent" -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$Filename"
}
Function Install_Editor_NotepadPlusPlus{
    #DOWNLOAD PARAMETER
    $Appname = "Notepad++"
    $homeUrl = 'https://notepad-plus-plus.org'
    
    #DOWNLOAD
    Write-Log "Downloading $Appname..."

    #OPEN URL
    #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $res = Invoke-WebRequest -UseBasicParsing $homeUrl
    if ($res.StatusCode -ne 200) {throw ("status code to getDownloadUrl was not 200: "+$res.StatusCode)}
    #FIND LINK TO LATEST VERSION
    $tempUrl = ($res.Links | Where-Object {$_.outerHTML -like "*Current Version *"})[0].href
    if ($tempUrl.StartsWith("/")) { $tempUrl = "$homeUrl$tempUrl" }
    $res = Invoke-WebRequest -UseBasicParsing $tempUrl
    if ($res.StatusCode -ne 200) {throw ("status code to getDownloadUrl was not 200: "+$res.StatusCode)}
    #FIND FILE LINK WITH X64
    $dlUrl = ($res.Links | Where-Object {$_.href -like "*x64.exe"})[0].href
    if ($dlUrl.StartsWith("/")) { $dlUrl = "$homeUrl$dlUrl" }
    #SET DOWNLOAD TEMPDIR AND START DOWNLOAD
    $installerPath = Join-Path $env:TEMP (Split-Path $dlUrl -Leaf)
    
    #Invoke-WebRequest $dlUrl -OutFile $installerPath
    #$path = "$global:TempFolder\$Filename"
    #[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClientNPP = new-object System.Net.WebClient
    $webClientNPP.DownloadFile( $dlUrl, $installerPath )

    #INSTALL SILENTLY WITH ADMIN RIGHTS
    Write-Log "Installing $Appname..."
    Start-Process -FilePath $installerPath -Args "/S" -Verb RunAs -Wait

    #CLEANUP
    $DesktopLink = "C:\Users\Public\Desktop\Notepad++.lnk"
    if (Test-Path $DesktopLink){Remove-Item $DesktopLink}
    Remove-Item $installerPath
}
Function Uninstall_Tool_OfficeRemoval{
    #DOWNLOAD PARAMETER
    $Appname = "Office Removal"
    #DownloadURL = "https://aka.ms/SaRA-officeUninstallFromPC"
    $DownloadURL = "https://outlookdiagnostics.azureedge.net/sarasetup/SetupProd_OffScrub.exe"
    $Filename = "SetupProd_OffScrub.exe"
    
    #DOWNLOAD
    Write-Log "Downloading $Appname..."
    Invoke-WebRequest -Uri $DownloadURL  -OutFile "$global:TempFolder\$Filename" 

    #INSTALL
    Write-Log "Running $Appname..."
    Start-Process "$global:TempFolder\$Filename" -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$Filename"
}
###############################################
Function Install_Backup_VeeamAgentforWindows {}
Function Install_Browser_Edge {
    Write-Log "Installing Microsoft Edge Chrome Version"
    $EdgeDownloadURL = "https://go.microsoft.com/fwlink/?linkid=2108834&Channel=Stable&language=de"
    $EdgeFileName = "MicrosoftEdgeSetup.exe"
    
    #DOWNLOAD
    Start-BitsTransfer -Source $EdgeDownloadURL -Destination "$global:TempFolder\$EdgeFileName"
     
    #INSTALL
    Start-process "taskkill" -ArgumentList "/f /im MicrosoftEdge.exe" -ErrorAction SilentlyContinue
    Start-process "taskkill" -ArgumentList "/f /im MicrosoftEdgeCP.exe" -ErrorAction SilentlyContinue
    Start-Process "$global:TempFolder\$EdgeFileName" -ArgumentList "/silent /install" -NoNewWindow -Wait

    #CONFIG
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge" -Force | Out-Null
    }
    If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge")) {
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge\Main" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge\Main" -Name "PreventFirstRunPage" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name "DisableEdgeDesktopShortcutCreation" -Type DWord -Value 1

    #CLEANUP
    Start-Process "taskkill" -ArgumentList "/IM msedge.exe /f" -ErrorAction SilentlyContinue
    Remove-Item "C:\Users\Public\Desktop\Microsoft Edge.lnk" -ErrorAction SilentlyContinue
    Remove-Item "$Env:USERPROFILE\Desktop\Microsoft Edge.lnk" -ErrorAction SilentlyContinue
    Clear-TempDir
}
Function Install_Browser_Firefox {
    
    #DOWNLOAD PARAMETER
    $Appname = "Firefox"
    #$DownloadURL = "https://cdn.stubdownloader.services.mozilla.com/builds/firefox-stub/de/win/10ddd7ac7926034e1b52f24128f466a1fa31e70806b5ca8b508a29b7b8903635/Firefox%20Installer.exe"
    #$DownloadURL = "https://download.mozilla.org/?product=firefox-stub&os=win&lang=de&attribution_code=c291cmNlPXd3dy5nb29nbGUuY29tJm1lZGl1bT1yZWZlcnJhbCZjYW1wYWlnbj0obm90IHNldCkmY29udGVudD0obm90IHNldCkmZXhwZXJpbWVudD0obm90IHNldCkmdmFyaWF0aW9uPShub3Qgc2V0KSZ1YT1lZGdl&attribution_sig=c936ba59abbeac3b860ac738e6dcc23356514d26f2b2465988655705bf944cc9"
    $DownloadURL = "https://cdn.stubdownloader.services.mozilla.com/builds/firefox-stub/de/win/10ddd7ac7926034e1b52f24128f466a1fa31e70806b5ca8b508a29b7b8903635/Firefox%20Installer.exe"
    $Filename = "Firefox Installer.exe"
    
    #DOWNLOAD
    #Write-Log "Downloading $Appname..."
    #Invoke-WebRequest -Uri $DownloadURL  -OutFile "$global:TempFolder\$Filename" 
    #Start-BitsTransfer -Source $DownloadURL -Destination "$global:TempFolder\$Filename"
    Write-Log "Downloading $Appname..."
    $path = "$global:TempFolder\$Filename"
    [Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURL, $path )

    #INSTALL
    Write-Log "Installing $Appname..."
    Start-Process "$global:TempFolder\$Filename" -ArgumentList "-ms -ma" -Wait -NoNewWindow
    # /DesktopShortcut=false /TaskbarShortcut=false

    #CLEANUP
    Start-Sleep -Seconds 3
    Start-Process "taskkill" -ArgumentList "/IM Firefox.exe /f" -ErrorAction SilentlyContinue
    Remove-Item "$global:TempFolder\$Filename"
    #Remove-Item "C:\Users\Public\Desktop\Firefox.lnk"
    }
    
Function Install_Archive_7Zip {
    Write-Log "Installing 7ZIP..."
    
    #DOWNLOAD PARAMETER
    $DownloadURL = "https://www.7-zip.org/a/7z1900-x64.exe"
    $FileName = "7z1900-x64.exe"
    
    #DOWNLOAD
    #Start-BitsTransfer -Source $7ZipDownloadURL -Destination "$global:TempFolder\$7ZipFileName" 
    $path = "$global:TempFolder\$Filename"
    [Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $webClient = new-object System.Net.WebClient
    $webClient.DownloadFile( $DownloadURL, $path )

    #INSTALL
    Start-Process "$global:TempFolder\$FileName" -ArgumentList "/S" -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$FileName"
}

Function Install_Tool_Everything {
    #PARAMETERS, BASIC
    $Global:ApplicationName = "Everything"
    $urlContainingDownloadlink = "https://www.voidtools.com/downloads"
    $SearchTextx86 = "Download Installer"
    $SearchTextx64 = "Download Installer 64-bit"

    #GRAB DOWNLOADLINK FROM WEBSITE
    #LOAD MSHTM DLL
    #$path = 'C:\Program Files (x86)\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll'
    #$Path = Resolve-Path -Path $Path -ErrorAction:Stop
    #$fs = ([System.IO.FileInfo] (Get-Item -Path $Path)).OpenRead()
    #$buffer = New-Object Byte[] $fs.Length
    #$n = $fs.Read($buffer, 0, $fs.Length)
    #$fs.Close()
    #if ( $n -gt 0 ){[System.Reflection.Assembly]::Load($buffer) | Out-Null}

    #BROWSE WEBSITE TO DETECT DOWNLOAD URL
    Write-Log "Detecting Downloadlinks for $global:ApplicationName..." 0
    $ie = New-Object -Com InternetExplorer.Application 
    $ie.Visible = $false
    $ie.Silent = $true
    #$ie.Navigate($urlContainingDownloadlink) 
    $ie.Navigate($urlContainingDownloadlink)
    while ($ie.ReadyState -ne 4) {Start-Sleep -m 1000};
    #while($ie.Busy) { Start-Sleep -Milliseconds 100 }

    #DEBUG
    #$ie.Document | Get-Member

    #FILTER ELEMENTS FOR DOWNLOAD LINKS
    $Global:DownloadURL32 = ($ie.Document.getElementsByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx86}).href
    $Global:DownloadURL64 = ($ie.Document.getElementsByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx64}).href
    #$Global:DownloadURL32 = ($ie.Document.IHTMLDocument3_getElementsByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx86}).href
    #$Global:DownloadURL64 = ($ie.Document.IHTMLDocument3_getElementsByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx64}).href
    #$Global:DownloadURL32 = ($ie.Document.IHTML3_getElementByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx86}).href
    #$Global:DownloadURL64 = ($ie.Document.IHTML3_getElementByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx64}).href
    #$Global:DownloadURL32 = ($ie.Document.documentElement.getElementByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx86}).href
    #$Global:DownloadURL64 = ($ie.Document.documentElement.getElementByTagName("*") | Where-Object {$_.className -like "*button*" -AND $_.outerText -like $SearchTextx64}).href
    Write-Log "Download URL x86: $Global:DownloadURL32" 9
    Write-Log "Download URL x64: $Global:DownloadURL64" 9
    $ie.Quit()

    #PARAMETER
    #$Global:DownloadURL64 = 'https://www.voidtools.com/Everything-1.4.1.1005.x64-Setup.exe'
    #$Global:DownloadURL32 = 'https://www.voidtools.com/Everything-1.4.1.1005.x86-Setup.exe'
    $EverySettings = "$Env:APPDATA\Everything\Everything.ini"
    $EveryMainSettings = "C:\Program Files\Everything\Everything.ini"

    #RESOLVE
    Resolve-DownloadURL

    #DOWNLOAD
    Write-Log "Downloading $global:ApplicationName..." 0
    if($Global:DowloadFinalURL){Get-FileWebclient $Global:DowloadFinalURL -includeStats}


    #EXTRACT
    Write-Log "Extracing Archive $global:ApplicationName..." 0
    Extract-Archive 

    #INSTALL
    Write-Log "Installing $global:ApplicationName..." 0
    if($Global:DowloadFinalURL -AND $global:ExeFileName){
            if (Test-Path "$global:TempFolder\$global:ExeFileName"){
            $InstallFile = Join-Path "$global:TempFolder\" "$global:ExeFileName"
            Start-Process $InstallFile -ArgumentList "/S" -NoNewWindow -Wait
        }
    }
    Pop-Location

    #CONFIG
    Write-Log "Configuring $global:ApplicationName..." 0
    if(Test-Path "$global:TempFolder\$global:ExeFileName"){
    Start-Process "C:\Program Files\Everything\Everything.exe" -ArgumentList "-install-service"
    (Get-Content $EveryMainSettings).replace('run_as_admin=1', 'run_as_admin=0') | Set-Content $EveryMainSettings

    If (!(Test-Path $EverySettings)) {New-Item -Path "$EverySettings" -Force | Out-Null} 
    (Get-Content $EverySettings).replace('alternate_row_color=0', 'alternate_row_color=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('show_mouseover=0', 'show_mouseover=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('hide_empty_search_results=0', 'hide_empty_search_results=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('language=1031', 'language=0') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('show_size_in_statusbar=0', 'show_size_in_statusbar=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('double_click_path=0', 'double_click_path=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('show_number_of_results_with_selection=0', 'show_number_of_results_with_selection=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('exclude_hidden_files_and_folders=0', 'exclude_hidden_files_and_folders=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('exclude_system_files_and_folders=0', 'exclude_system_files_and_folders=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('filters_visible=0', 'filters_visible=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('preview_visible=0', 'preview_visible=1') | Set-Content $EverySettings
    (Get-Content $EverySettings).replace('exclude_folders=', 'exclude_folders="C:\\Windows","C:\\Windows.old","C:\\Windows10Upgrade","C:\\Program Files (x86)","C:\\Program Files"') | Set-Content $EverySettings
    #minimize_to_tray=0
    #check_for_updates_on_startup=1
    }

    #CLEANUP
    Clear-TempDir
}

Function Install_Tool_JDAST {
    
    #PARAMETERS, BASIC
    $Global:ApplicationName = "JDAST"
    # "https://web.archive.org/web/20180907232222if_/http://www.gmwsoftware.co.uk/files/JDast_installer.exe"
    $Global:DownloadURL32 = "https://softpedia-secure-download.com/dl/a81f9dd6d691c226ead4ce29b9325365/6037d9e1/100184855/software/network/JDast_installer.exe"
    $Global:DownloadURL64 = "https://softpedia-secure-download.com/dl/a81f9dd6d691c226ead4ce29b9325365/6037d9e1/100184855/software/network/JDast_installer.exe"
    $ConfigLocation = "$env:appdata\jdast\main.ini"
    $DownloadServerList = "$env:appdata\jdast\locmenu.ini"
    $ActiveServerList = "$env:appdata\jdast\mult_locmenu.txt"

    #RESOLVE
    Resolve-DownloadURL

    #DOWNLOAD
    if($Global:DowloadFinalURL){Get-FileWebclient $Global:DowloadFinalURL -includeStats}


    #EXTRACT
    Extract-Archive 

    #INSTALL
    Write-Log "Installing $global:ApplicationName..." 0
    if($Global:DowloadFinalURL -AND $global:ExeFileName){
            if (Test-Path "$global:TempFolder\$global:ExeFileName"){
            $InstallFile = Join-Path "$global:TempFolder\" "$global:ExeFileName"
            Start-Process $InstallFile -ArgumentList "/S" -NoNewWindow -Wait
        }
    }
    Pop-Location

    #CONFIG
    Write-Log "Configuring $global:ApplicationName..." 0
    #(Get-Content $ConfigLocation).replace('language=2', 'language=0') | Set-Content $ConfigLocation
    add-Content "$ConfigLocation" '[main]'
    add-Content "$ConfigLocation" 'language=0'
    add-Content "$ConfigLocation" 'autoupdate=4'
    add-Content "$ConfigLocation" 'autoupdate_beta=4'
    add-Content "$ConfigLocation" 'OTS_Ping_loc=google.de'
    add-Content "$ConfigLocation" 'ots_fedc_wan=4'
    add-Content "$ConfigLocation" ''
    ###SETUP UPLOAD FTP SERVER###
    #KABEL DEUTSCHLAND
    #add-Content "$ConfigLocation" 'ftpserver=ftp.rzg.mpg.de'
    #add-Content "$ConfigLocation" 'ftpdefaultON=4'
    #add-Content "$ConfigLocation" 'ftpusername='
    #add-Content "$ConfigLocation" 'ftppassword=0'
    #add-Content "$ConfigLocation" 'ftpuplocation=/pub/test/'

    #add-Content "$ConfigLocation" 'ftpserver=ftp.dialogika.de'
    #add-Content "$ConfigLocation" 'ftpdefaultON=4'
    #add-Content "$ConfigLocation" 'ftpusername='
    #add-Content "$ConfigLocation" 'ftppassword=0'
    #add-Content "$ConfigLocation" 'ftpuplocation=/upload'

    #TELEKOM
    add-Content "$ConfigLocation" 'ftpserver=hgd-speedtest-1.tele2.net'
    add-Content "$ConfigLocation" 'ftpdefaultON=4'
    add-Content "$ConfigLocation" 'ftpusername='
    add-Content "$ConfigLocation" 'ftppassword=0'
    add-Content "$ConfigLocation" 'ftpuplocation=/upload'
    add-Content "$ConfigLocation" 'UploadMultiThreads=5'

    #1&1 
    #add-Content "$ConfigLocation" 'ftpserver=212.93.6.58'
    #add-Content "$ConfigLocation" 'ftpdefaultON=4'
    #add-Content "$ConfigLocation" 'ftpusername=versatel'
    #add-Content "$ConfigLocation" 'ftppassword=versatel'
    #add-Content "$ConfigLocation" 'ftpuplocation='

    ###SETUP APPEARANCE###
    add-Content "$ConfigLocation" '[graphview]'
    add-Content "$ConfigLocation" 'graphlayout=3'

    #MODIFY DOWNLOAD SERVER LIST, DELETING ALL BAD SERVERS
    #ADD SERVERS, DEPENDS AN ISP / LOCATION

    #ALL SERVERS IN CONFIG
    <#
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/10mb.bin 10'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://82.15.207.2/speedtest/download/1MB?sc_1234900057_415,1'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'ftp://ftp.free.fr/mirrors/download.linuxtag.org/knoppix/contrib/grubboot.zip,1.414'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://download.thinkbroadband.com/5MB.zip,5'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'ftp://ftp.free.fr/mirrors/download.linuxtag.org/knoppix/qemu-0.8.1/qemu.exe,7.22'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/10mb.bin,10'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://download.thinkbroadband.com/10MB.zip,10'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://download.thinkbroadband.com/20MB.zip,20'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/25mb.bin,25'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://download.thinkbroadband.com/50MB.zip,50'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/50mb.bin,50'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://download.thinkbroadband.com/100MB.zip,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/100mb.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://tokyo1.linode.com/100MB-tokyo.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://london1.linode.com/100MB-london.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://newark1.linode.com/100MB-newark.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://atlanta1.linode.com/100MB-atlanta.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://dallas1.linode.com/100MB-dallas.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://fremont1.linode.com/100MB-fremont.bin,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://cachefly.cachefly.net/100mb.test,100'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/250mb.bin,250'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/500mb.bin,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://95.215.63.211/1000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://test.unet.nl/1000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://www.pro-noc.nl/1000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://chi01-spd.neosurge.com/1GBtest.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://ca01-spd.neosurge.com/1GBtest.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://www.bbned.nl/scripts/speedtest/download/file1000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.tweak.nl/1000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://limestonenetworks.com/test500.zip,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.dal01.softlayer.com/downloads/test500.zip,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.sea01.softlayer.com/downloads/test500.zip,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://speedtest.wdc01.softlayer.com/downloads/test500.zip,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://noc.cdp.pl/get/500mb.cdp,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://mirror.intrapower.net.au/500MB.dat,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://noc.gts.pl/500mb.gts,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://www.singlehop.com/speedtest/500megabytefile,500'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://67.159.44.209/1GBtest.zip,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://208.85.242.69/1g.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://74.63.66.114/1GBtest.zip,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://76.73.0.4/1GBtest.zip,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://mirrors.nfsi.pt/speedtest/10000mb.bin,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://test.north.kz/downloads/1000mb.test.txt,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://vinax.net/1GBtest.zip,1000'} | Set-Content "$DownloadServerList"
    Get-Content "$DownloadServerList" | Where-Object {$_ -notmatch 'http://lg.denver.fdcservers.net/1GBtest.zip,1000'} | Set-Content "$DownloadServerList"
    #>

    #REMOVE SERVERS (BAD PERFORMANCE), DEPENDS AN ISP / LOCATION
    Get-Content "$ActiveServerList" | Where-Object {$_ -notmatch 'http://cachefly.cachefly.net/100mb.test,100'} | Set-Content "$DownloadServerList"


    #CLEANUP
    Clear-TempDir
}
Function Install_Tool_DesktopRestore {}
Function Install_Tool_NetTime {
    Write-Log "Installing and Configuring NetTime..."

    #PARAMETERS
    $NettimeDownloadUrlDe64 = "http://www.timesynctool.com/NetTimeSetup-314.exe"
    $NettimeRegistrySettings = "HKLM:\Software\Wow6432Node\Subjective Software\NetTime"

    #DOWNLOAD
    Start-BitsTransfer -Source $NettimeDownloadUrlDe64 -Destination "$global:TempFolder\nettime.exe"
 
    #CONFIG
    New-Item $NettimeRegistrySettings -Force | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "AlwaysProvideTime" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "AutomaticUpdateChecks" -Value "1" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "DaysBetweenUpdateChecks" -Value "7" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "DemoteOnErrorCount" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Hostname" -Value "ptbtime1.ptb.de" -PropertyType String | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Hostname1" -Value "ptbtime2.ptb.de" -PropertyType String | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Hostname2" -Value "ptbtime3.ptb.de" -PropertyType String | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Hostname3" -Value "" -PropertyType String | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LargeAdjustmentAction" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LargeAdjustmentThreshold" -Value "1" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LargeAdjustmentThresholdUnits" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LastUpdateCheck" -Value "1506704632" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LogLevel" -Value "1" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LostSync" -Value "24" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "LostSyncUnits" -Value "3" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Port" -Value "123" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Port1" -Value "123" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Port2" -Value "123" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Port3" -Value "123" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Protocol" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Protocol1" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Protocol2" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Protocol3" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Retry" -Value "3" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "RetryUnits" -Value "2" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "Server" -Value "0" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "SyncFreq" -Value "10" -PropertyType DWord | Out-Null
    New-ItemProperty -Path $NettimeRegistrySettings -Name "SyncFreqUnits" -Value "2" -PropertyType DWord | Out-Null
    Pop-Location

    #INSTALL
    Start-Process -FilePath "$global:TempFolder\nettime.exe" -ArgumentList "/SILENT /NORESTART"

    #CLEANUP
    Start-Sleep -s 5
    Remove-Item "$global:TempFolder\nettime.exe"
}
Function Install_AdobeReader {
    #PARAMETERS
    $AppName = "Adobe Reader"
    $DownloadUrl_64 = "https://admdownload.adobe.com/bin/live/readerdc_de_a_install.exe"
    $DownloadUrl_32 = ""
    $Filename_64 = "readerdc_de_a_install.exe"
    $Filename_32 = ""

    #DOWNLOAD
    Write-Log "Downloading $AppName..."
    Push-Location $global:TempFolder
    
    if ($global:bittype -eq "64-Bit") {Invoke-WebRequest -Uri $DownloadUrl_64  -OutFile "$global:TempFolder\$Filename_64"}
    if ($global:bittype -eq "32-Bit") {Start-BitsTransfer -Source $DownloadUrl_32}

    #EXTRACT
    #if ($global:bittype -eq "64-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_64" -DestinationPath $global:TempFolder -Force}
    #if ($global:bittype -eq "32-Bit") {Expand-Archive -LiteralPath "$global:TempFolder\$Filename_32" -DestinationPath $global:TempFolder -Force}

    #INSTALL
    Write-Log "Installing $AppName..."
    if ($global:bittype -eq "64-Bit") {Start-Process "$global:TempFolder\$Filename_64" -ArgumentList "/sAll" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit") {Start-Process "$global:TempFolder\$Filename_32" -ArgumentList "/sAll" -NoNewWindow -Wait}
    Pop-Location

    #CONFIG
    #SET AS DEFAULT
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithList" /v a /t REG_SZ /d AcroRd32.exe /f

    #SURPRESS TOUR ON EVERY START
    $strKey = "Software\Adobe\Acrobat Reader\DC\FTEDialog"
    $CurLoc = Get-Location
    $HKU = Get-PSDrive HKU -ea silentlycontinue
    if (!$HKU ) {
        New-PSDrive -Name HKU -PsProvider Registry HKEY_USERS | out-null
        Set-Location HKU:
    }
 
    # select all desired user profiles, exlude *_classes & .DEFAULT
    $regProfiles = Get-ChildItem -Path HKU: | Where-Object { ($_.PSChildName.Length -gt 8) -and ($_.PSChildName -notlike "*.DEFAULT") }

    # loop through all selected profiles & delete registry
    ForEach ($profile in $regProfiles ) {
        If(Test-Path -Path $profile\$strKey){
            Remove-Item -Path $profile\$strKey -recurse
            New-ItemProperty -Path $profile\$strKey -Name "iFTEVersion" -Value "10" -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -path $profile\$strKey -Name "iLastCardShown" -Value "0" -PropertyType DWORD -Force | Out-Null
            }
    }
 
    # return to initial location at the end of the execution
    Set-Location $CurLoc
    Remove-PSDrive -Name HKU

    #CLEANUP
    Start-Sleep -s 2
    if ($global:bittype -eq "64-Bit") {Remove-Item "$global:TempFolder\$Filename_64"}
    if ($global:bittype -eq "32-Bit") {Remove-Item "$global:TempFolder\$Filename_32"}
    Remove-Item "C:\Users\Public\Desktop\Acrobat Reader DC.lnk"

    }
Function Install_RemoteControl_SolarwindsAgent {
    Write-Log "Installing Solarwinds Agent..."

    #DOWNLOAD PARAMETER
    #$SWADownloadURL = "https://xxx.de/dms/FileDownload?customerID=142&softwareID=101"
    $SWAFileName = "WindowsAgentSetup.exe"

    #DOWNLOAD
    Start-BitsTransfer -Source $SWADownloadURL -Destination "$global:TempFolder\$SWAFileName" 

    #INSTALL
    Start-Process "taskkill" -ArgumentList "/IM agent.exe /f" -ErrorAction SilentlyContinue
    Start-Process "$global:TempFolder\$SWAFileName" -ArgumentList "-ai" -NoNewWindow -Wait

    #CLEANUP
    Remove-Item "$global:TempFolder\$SWAFileName"
}
Function Install_KeePass {}

Function Install_VPN_FortiClient {
    Write-Log "Downloading FortiClient..."
    #$FortiClientDownloadUrl_626_64 = "https://www.xxx.de/Download/FortiClientVPNSetup_6.2.6.0951_x64.exe"
    #$FortiClientDownloadUrl_626_32 = "https://www.xxx.de/Download/FortiClientVPNSetup_6.2.6.0951.exe"
    $FortiClientVersion = "6.2.6"
    #$FortiClientConfig = "https://www.xxx.de/Download/FortiClientSettings_609.xml"
    $FortiClientSettings = "C:\Temp\settings.xml"

    #DOWNLOAD
    Push-Location $global:TempFolder
    if ($global:bittype -eq "64-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Start-BitsTransfer -Source $FortiClientDownloadUrl_626_64}
    if ($global:bittype -eq "32-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Start-BitsTransfer -Source $FortiClientDownloadUrl_626_32}
    Pop-Location

    #EXTRACT
    if ($global:bittype -eq "64-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {}
    if ($global:bittype -eq "32-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {}

    #INSTALL
    Write-Log "Installing FortiClient..."
    if ($global:bittype -eq "64-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Start-Process "$global:TempFolder\FortiClientVPNSetup_6.2.6.0951_x64.exe" -ArgumentList "/passive /norestart" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Start-Process "$global:TempFolder\FortiClientVPNSetup_6.2.6.0951.exe" -ArgumentList "/passive /norestart" -NoNewWindow -Wait}

    #CONFIG
    #GET CUSTOMER CONFIG
    Write-Log "Configuring FortiClient..."
    $ScriptDir = Split-Path $psise.CurrentFile.FullPath
    Push-Location $ScriptDir
    Get-Content 0_FORTICLIENT-SETTINGS.txt | Where-Object {$_ -notmatch '#'} | Foreach-Object{
       $var = $_.Split('=')
       Set-Variable -Name $var[0] -Value $var[1]
    }
    Write-Host "Setting VPN Connection"
    $FortiClientCustomer
    $FortiClientGateway
    $FortiClientPort
    $FortiClientUser
    Pop-Location

    #SET CONFIG
    $FCwebclient = New-Object System.Net.WebClient
    $FCwebclient.DownloadFile($FortiClientConfig,$FortiClientSettings)

    #REPLACE DEFAULT CONFIG WITH CUSTOM SETTINGS
    (Get-Content $FortiClientSettings).replace('ConnectionName', $FortiClientCustomer) | Set-Content $FortiClientSettings
    (Get-Content $FortiClientSettings).replace('ConnectionGateway', $FortiClientGateway) | Set-Content $FortiClientSettings
    (Get-Content $FortiClientSettings).replace('34463', $FortiClientPort) | Set-Content $FortiClientSettings
    (Get-Content $FortiClientSettings).replace('Enc 109b72b0696694cdafc70e8e619742d9bb7abad165408a071625edcee0911b2230c9b5e727725f155275', $FortiClientUser) | Set-Content $FortiClientSettings

    #IMPORT MODIFIED SETTINGS
    if ($global:bittype -eq "64-Bit") {Start-Process "C:\Program Files\Fortinet\FortiClient\FCConfig.exe"       -ArgumentList "-f $FortiClientSettings -m vpn -o import" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit") {Start-Process "C:\Program Files (x86)\Fortinet\FortiClient\FCConfig.exe" -ArgumentList "-f $FortiClientSettings -m vpn -o import" -NoNewWindow -Wait}

    #CLEANUP
    Start-Sleep -s 5
    if ($global:bittype -eq "64-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Remove-Item "$global:TempFolder\FortiClientVPNSetup_6.2.6.0951_x64.exe"}
    if ($global:bittype -eq "32-Bit" -AND ($FortiClientVersion -eq "6.2.6")) {Remove-Item "$global:TempFolder\FortiClientVPNSetup_6.2.6.0951.exe"}
    Remove-Item $FortiClientSettings

}
Function Install_UblockFF {
    #$url = "https://addons.mozilla.org/de/firefox/addon/ublock-origin/"
}
Function Install_UblockChrome {}
Function Install_UblockEdge {
    #$URL = "https://microsoftedge.microsoft.com/addons/detail/odfafepnkmbhccpbejgmiehpchacaeak"
}
Function Install_PowershellV7 {
    Write-Log "Downloading Powershell V7..."
    $PowershellDownloadUrl_64 = "https://github.com/PowerShell/PowerShell/releases/download/v7.0.2/PowerShell-7.0.2-win-x64.msi"
    $PowershellDownloadUrl_32 = ""

    #DOWNLOAD
    Push-Location $global:TempFolder
    if ($global:bittype -eq "64-Bit") {Start-BitsTransfer -Source $PowershellDownloadUrl_64}
    if ($global:bittype -eq "32-Bit") {Start-BitsTransfer -Source $PowershellDownloadUrl_32}
    Pop-Location

    #EXTRACT
    if ($global:bittype -eq "64-Bit") {}
    if ($global:bittype -eq "32-Bit") {}

    #INSTALL
    Write-Log "Installing Powershell v7..."
    if ($global:bittype -eq "64-Bit") {Start-Process "$global:TempFolder\PowerShell-7.0.2-win-x64.msi"} #-ArgumentList "/passive /norestart" -NoNewWindow -Wait}
    if ($global:bittype -eq "32-Bit") {Start-Process "$global:TempFolder\PowerShell-7.0.2-win-x32.msi"} #-ArgumentList "/passive /norestart" -NoNewWindow -Wait}

    #CLEANUP
    Start-Sleep -s 5
    if ($global:bittype -eq "64-Bit") {Remove-Item "$global:TempFolder\PowerShell-7.0.2-win-x64.msi"}
    if ($global:bittype -eq "32-Bit") {Remove-Item "$global:TempFolder\PowerShell-7.0.2-win-x32.msi"}
    }
Function Install_VSCode {
    Write-Log "Downloading Visual StudioCode..."
    $VSCodeDownloadUrl_64 = "https://aka.ms/win32-x64-user-stable"
    $VSCodeDownloadUrl_32 = ""

    #DOWNLOAD
    Push-Location $global:TempFolder
    if ($global:bittype -eq "64-Bit") {Start-BitsTransfer -Source $VSCodeDownloadUrl_64}
    if ($global:bittype -eq "32-Bit") {Start-BitsTransfer -Source $VSCodeDownloadUrl_32}
    Pop-Location

    #EXTRACT
    if ($global:bittype -eq "64-Bit") {}
    if ($global:bittype -eq "32-Bit") {}

    #INSTALL
    Write-Log "Installing VisualStudio Code..."
    #if ($global:bittype -eq "64-Bit") {Start-Process "$global:TempFolder\PowerShell-7.0.2-win-x64.msi" #-ArgumentList "/passive /norestart" -NoNewWindow -Wait}
    #if ($global:bittype -eq "32-Bit") {Start-Process "$global:TempFolder\PowerShell-7.0.2-win-x32.msi" #-ArgumentList "/passive /norestart" -NoNewWindow -Wait}

    #CLEANUP
    #Start-Sleep -s 5
    #if ($global:bittype -eq "64-Bit") {Remove-Item "$global:TempFolder\PowerShell-7.0.2-win-x64.msi"}
    #if ($global:bittype -eq "32-Bit") {Remove-Item "$global:TempFolder\PowerShell-7.0.2-win-x32.msi"}
}


##########################
###      FUNCTIONS     ###
###       UPDATES      ###
##########################
#V1903, Build 18362.239
Function Update_StoreApps_All {
    Write-Log "Updating Store Apps"

    $namespaceName = "root\cimv2\mdm\dmmap"
    $className = "MDM_EnterpriseModernAppManagement_AppManagement01"
    $wmiObj = Get-WmiObject -Namespace $namespaceName -Class $className
    $result = $wmiObj.UpdateScanMethod()
}
Function Update_Windows_CurrentlyAvailable {
    Write-Log "Updating Windows via UsoClient"

    UsoClient ScanInstallWait

    Write-Log "Updating Windows via WUUpdates"

    Install-WUUpdates -Updates (Start-WUScan)
}
Function Update_Windows_WsusOffline {
    Write-Log "Installing and Running WSUS Offline..."

    #Parameter
    #$InstallDir_WsusOffline = "\\192.168.178.22\WindowsUpdate"

    #FOR LOCAL USE
    $WSUSDownloadUrl = "https://download.wsusoffline.net/wsusoffline119.zip"
    $WSUSTempDir = "C:\Windows\Temp"
    $WSUSInstallDir = "C:"

    #DOWNLOADS WSUSOFFLINE
    If(!(test-path $WSUSTempDir)){New-Item -ItemType Directory -Force -Path $WSUSTempDir}
    Start-BitsTransfer -Source $WSUSDownloadUrl -Destination "$WSUSTempDir\wsusoffline119.zip"
    
    #EXTRACT
    If(!(test-path $WSUSInstallDir)){New-Item -ItemType Directory -Force -Path $WSUSInstallDir}
    Expand-Archive -LiteralPath "$WSUSTempDir\wsusoffline119.zip" -DestinationPath $WSUSInstallDir
    
    #Enable Windows Script Host
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings" -Name "Enabled" -Type DWord -Value 1

    #DOWNLOADS WIN10 64 UPDATE FILES, 10GB
    Start-Process "$WSUSInstallDir\wsusoffline\cmd\DownloadUpdates.cmd" -ArgumentList "w100-x64 glb /includedotnet /includemsse /includewddefs /verify" -Wait
    #INSTALLS UPDATES AND REBOOTS
    Start-Process "$WSUSInstallDir\wsusoffline\client\cmd\doupdate.cmd" -ArgumentList "/all /autoreboot" -Wait

    #CLEANUP
    timeout /t 3
    Remove-Item $WSUSTempDir\wsusoffline119.zip  -ErrorAction SilentlyContinue

    #Disable Windows ScriptHost again
    #Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings" -Name "Enabled" -Type DWord -Value 0
}
Function Update_Windows_PSWindowsUpdate {
    Write-Log "Installing PSWindowsUpdate Plugin and Installing Updates
    ..."
    If(-not(Get-PackageProvider Nuget -Force -ErrorAction silentlycontinue)){
    Set-PSRepository PSGallery -InstallationPolicy Trusted
    }

    If(-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)){
    Install-Module PSWindowsUpdate -Confirm:$False -Force
    }
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -AutoReboot -Verbose #| Out-File "c:\$(get-date -f yyyy-MM-dd)-WindowsUpdate.log" -force 
}

##########################
###      FUNCTIONS     ###
###       GENERIC      ###
##########################
Function Generic_Explorer_Restart {
    Write-Log "Restarting Windows Explorer..."
    Get-Process explorer | Stop-Process -Force -ErrorAction SilentlyContinue
}
Function Generic_Useraccount_CreateAdmin {
    Write-Log "Creating Local Admin Account..."

    #IMPORT CREDENTIALS FROM CONFIG
    $ScriptDir = Split-Path $psise.CurrentFile.FullPath
    Push-Location $ScriptDir
    Get-Content 0_LOCAL-Credentials.txt | Where-Object {$_ -notmatch '^#.*'} | Foreach-Object{
        $var = $_.Split('=')
    Set-Variable -Name $var[0] -Value $var[1]
    }

    #CREATE LOCAL ADMIN ACCOUNT (NO WINDOWS LIVE)
    NET USER $Local_Username $Local_Password /ADD
    NET LOCALGROUP "Administratoren" "$Local_Username" /add

    #DELETE EXISTING LENOVO AND CUSTOM ACCOUNTS 
    $ListOfUsersToDelete = (
    "aa", 
    "lenovo"
    )

    foreach ($User in $ListOfUsersToDelete){
        if (Get-LocalUser $User -ErrorAction SilentlyContinue){
            NET USER $User /DELETE /YES
        }
        else{
            Write-Log -ForegroundColor green "$User not found on host" 2
            }
    }
}

Function Generic_Useraccount_DefaultPassword {
    #CHANGE PASSWORD IF NO DOMAIN USER
    if(!$env:USERDNSDOMAIN){net user $env:USERNAME "Start1234"}
    else{net user $env:USERNAME "Start1234" /DOMAIN}
}
Function Generic_Useraccount_HideUsersLoginscreen {
    $UsersToHide = (
        "admin",
        "WINRM"
        )
    ForEach ($UserAccount in $UsersToHide) {Write-Log "Hiding User: $UserAccount from Loginscreen" 9}
    If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList")) {
	    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" -Force | Out-Null
    }
    ForEach ($UserAccount in $UsersToHide) {
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" -Name "$UserAccount" -Type DWord -Value 0
    }
}

##########################
###      FUNCTIONS     ###
###       DOMAIN       ###
##########################
Function Domain_DisallowLocalLogin {
    Write-Log "Denying Domain Users to Log onto this machine ..."
    #unsolved
}
Function Domain_JoinComputer {
    Write-Log "Starting Domain Join Process..."
    Write-Output -ForegroundColor "green" "please enter domain administrator password:"
    Add-Computer -DomainName $ADFQDN -Credential (Get-Credential -Credential Administrator) -Force
}

##########################
###      FUNCTIONS     ###
###         VPN        ###
##########################
Function VPN_Win10SetupClient {
    $DownloadUrlFAvmClient = "https://avm.de/fileadmin/user_upload/DE/Service/VPN/FRITZ_VPN64_German_win10_Installation.zip"
    $DownloadUrlFAvmConfigTool = "https://avm.de/fileadmin/user_upload/DE/Service/VPN/FRITZ_Box-Fernzugang_einrichten_Installation.zip"

    $VPNConnectionName = "FritzBox"
    #Add-VpnConnection -Name $VPNConnectionName -ServerAddress "xxx.de" -TunnelType IKEv2 -EncryptionLevel Required -AuthenticationMethod EAP -SplitTunneling -AllUserConnection
    Set-VpnConnectionIPsecConfiguration -ConnectionName $VPNConnectionName -AuthenticationTransformConstants GCMAES256 -CipherTransformConstants GCMAES256 -EncryptionMethod AES256 -IntegrityCheckMethod SHA256 -DHGroup Group14 -PfsGroup PFS2048 -PassThru
    Add-VpnConnectionRoute -Connectio
    nName $VPNConnectionName -DestinationPrefix 192.168.178.0/24 -PassThru
    }

##########################
###      FUNCTIONS     ###
###    RUN SELECTION   ###
##########################
#$Selection | ForEach-Object { Invoke-Expression $_ }
if($script:FormProperlyClosed){
    $script:ExecutionList | ForEach-Object { Invoke-Expression $_ }
    #Write-Host -NoNewLine "Press any key to continue...";
    #$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown");
}

Pause