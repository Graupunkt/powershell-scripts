Import-Module PSWindowsUpdate

<#
#KNOWN ERRORS
ERROR #1 - Die Datei "C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate\2.2.0.2\PSWindowsUpdate.psm1" kann nicht geladen werden, da die Ausführung von Skripts auf diesem System deaktiviert ist.
SOLUTION #1 - Install-Module -Name PSWindowsUpdate -Force
#>

#Select specific groups
#$AllWindows = Get-ADComputer -Filter "OperatingSystem -like 'Win*'" -credential $adcred -Properties *
#$Server = (Get-ADComputer -Filter {operatingsystem -like '*server*'} -credential $adcred -Properties *).Name
#$Clients = Get-ADComputer -Filter {operatingsystem -notlike '*server*'} -credential $adcred -Properties *

#IF ERRORS ON POWERSHELL REMTOE EXECUTION
#COMMAND TO TEST CONNECTIFITY = Test-WSMan "localhost"
#run commands on destination pcs with admin permissions, to allow remote commands to be executed via powershell. Warnung trustedhosts is set to any, you might wanna replace these with your local network address
<#
    winrm quickconfig -quiet
    Set-Service winrm -StartupType Automatic 
    Start-Service winrm
    Enable-PSRemoting -Force
    Set-Item wsman:\localhost\client\trustedhosts * -Force
    winrm s winrm/config/client '@{TrustedHosts="*"}'
    Restart-Service WinRM
#>

<#
MSG Unable to Download, Die Liste der verfügbaren anbieter ...
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Install PSWIndowsUpdate on current executing PC
# Allow Repository
If(-not(Get-PackageProvider Nuget -Force -ErrorAction silentlycontinue)){
    Write-Host -ForegroundColor Green "Settings PSGallery as Trusted Repository"
    Set-PSRepository PSGallery -InstallationPolicy Trusted
}

# Install Module PSWindows Update
If(-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)){
    Write-Host -ForegroundColor Green "Installing Module PSWindowsUpdate"
    Install-Module PSWindowsUpdate -Confirm:$False -Force
}

#Install or Update PSWindowsUpdate on Target-PCs
Update-WUModule -ComputerName $servers -Local -Confirm:$false
#>



function Install-WUonServers{
        <#
    .SYNOPSIS
    This script will automatically install all avaialable windows updates on a device and will automatically reboot if needed, after reboot, windows updates will continue to run until no more updates are available.
    .PARAMETER URL
    User the Computer parameter to specify the Computer to remotely install windows updates on.




    #>

    [CmdletBinding()]
    param (
    [parameter(Mandatory=$true,Position=1)]
    [string[]]$computer
    )
    Set-Variable -Name 'ConfirmPreference' -Value 'None' -Scope Global

    ForEach ($pc in $computer){
    #$computer | ForEach-Object -Parallel {
        #install pswindows updates module on remote machine
        if((Test-WSMan $pc -ErrorAction SilentlyContinue).wsmid -match "wsmanidentity\.xsd"){
            Invoke-Command -ComputerName $pc -Scriptblock{
                #ENABLE TEMPORARILY TLS 1.2 IN RARE CASES SOME SERVER 2016 HAVE THIS DISABLED, THIS IS NECESSARY TO DOWNLOAD NUGET THESE DAYS
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12                

                #CHECK IF NUGET AND PSWINDOWS UPDATE IS INSTALELD; IF NOT INSTALL THEM
                $ProgressPreference = "SilentlyContinue"
                if((Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue -Force).version -lt 2.8.5.201){
                    Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Updating PreRequest Nuget on $env:computername"
                    Install-PackageProvider -Name NuGet -RequiredVersion "2.8.5.201" -Force
                }
                If(-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)){
                    Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Installing PreRequest PSWindowsUpdate on $env:computername"
                    Install-Module PSWindowsUpdate -Confirm:$False -Force -Repository PSGallery
                    }
                $ProgressPreference = "Continue"
                Import-Module PSWindowsUpdate -force | Out-Null

                #CHECK IF PSWindows Update was installed sucessfully
                If(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue){
                    Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] $env:computername has prerequests installed"
                }else{
                    Write-Host -ForegroundColor Red "[$(get-date -Format "HH:mm:ss")] $env:computername could not install modul PSWIndowsUpdate"
                }
            }
            
        }else{
            Write-Host -ForegroundColor Red "[$(get-date -Format "HH:mm:ss")] $env:computername has WinRM not properly configured"
        }

        Do{
            #Reset Timeouts
            $connectiontimeout = 0
            $updatetimeout = 0
            $PreviousCount = 0
            $idlecount = 0
            $LogFileCount = 0

            #starts up a remote powershell session to the computer
            do{
                $session = New-PSSession -ComputerName $pc
                Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Connecting to $pc"
                sleep -seconds 10
                $connectiontimeout++
            } until ($session.state -match "Opened" -or $connectiontimeout -ge 10)
            #retrieves a list of available updates
            Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Checking for new updates available on $pc"
            $updates = invoke-command -session $session -scriptblock {Get-WindowsUpdate -NotCategory Drivers -verbose}
           
            #Write-Host -ForegroundColor White "[$(get-date -Format "HH:mm:ss")] Skipped Updates:"
            #foreach($update in $updates){$update}
            Write-Host -ForegroundColor White "[$(get-date -Format "HH:mm:ss")] Detected Updates (without drivers from pre check):"
            foreach($update in $updates){$update}

            #counts how many updates are available
            $global:updatenumber = ($updates.kb).count
            #if there are available updates proceed with installing the updates and then reboot the remote machine
            if ($updates -ne $null){
                #remote command to install windows updates, creates a scheduled task on remote computer
                invoke-command -ComputerName $pc -ScriptBlock { 
                    $path = "C:\logs"
                    If(!(test-path -PathType container $path))
                    {
                          New-Item -ItemType Directory -Path $path | Out-Null
                    }
                    Invoke-WUjob -ComputerName localhost -Script "Import-Module PSWindowsUpdate; Install-WindowsUpdate -NotCategory Drivers -AcceptAll -Verbose | Out-File C:\Windows\PSWindowsUpdate.log -Width 300" -Confirm:$false -RunNow
                 }
                #Show update status until the amount of installed updates equals the same as the amount of updates available
                Start-Sleep -Seconds 30
                Write-Host -ForegroundColor Gray "[$(get-date -Format "HH:mm:ss")] Processing updates (timeout 20mins until reboot)"
                do {
                    $logfile = "\\$pc\c$\Windows\PSWindowsUpdate.log"
                    if(!(test-path $logfile)){
                        Write-Host -ForegroundColor Red "Logfile is not accessable on target computer, canceling update "
                        break
                        }
                
                    #$updatestatus = Get-Content \\$pc\c$\Windows\PSWindowsUpdate.log
                    
                    $updatestatus = Get-Content \\$pc\c$\Windows\PSWindowsUpdate.log | Foreach {$_.TrimEnd()} | ? {$_.trim() -ne "" }
                    $LogFileCount = $updatestatus.Count
                    if($LogFileCount -ne $PreviousCount){
                        $CountDifference = $LogFileCount - $PreviousCount #Get the amount of new lines to be display
                        Write-Host " " #break nonewline from idle line
                        #Output multiple lines in the correct color to the user
                        $NewLines = $updatestatus[$PreviousCount..$LogFileCount] # Get the new lines since last update
                        foreach($line in $NewLines){
                            switch -Wildcard ($line){ # colorize output on content
                                "*Accepted*"     {$LineColor = "Gray" ;Break}
                                "*Downloaded*"   {$LineColor = "Gray" ;Break}
                                "*Installed*"    {$LineColor = "Green";Break}
                                "*Failed*"       {$LineColor = "Red"  ;Break}
                                "*ComputerName*" {$LineColor = "White";Break}
                                default          {$LineColor = "Gray"}
                            }
                            Write-Host -ForegroundColor $LineColor "[$(get-date -Format "HH:mm:ss")] $line"
                        }
                        $PreviousCount = $LogFileCount #store the current count for comparison in the next loop
                        $idlecount = 0
                    }else{
                        if($idlecount -eq 0){
                            #Write-Host "[$(get-date -Format "HH:mm:ss")] Updates Total: $updatenumber, Installed: $installednumber, Failed: $Failednumber"
                            Write-Host -ForegroundColor Gray -NoNewLine "no changes detected"
                        }
                        Write-Host -ForegroundColor Gray -NoNewLine "."
                        $idlecount++
                    }

                    Start-Sleep -Seconds 10
                    $ErrorActionPreference = ‘SilentlyContinue’
                    $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count
                    $global:Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count
                    $ErrorActionPreference = ‘Continue’
                    $updatetimeout++
                }until (($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 720)
                
                #restarts the remote computer and waits till it starts up again
                Write-Host -ForegroundColor Green "[[$(get-date -Format "HH:mm:ss")] Restarting remote computer (timeout 5mins) and waiting to process further more updates"
                #removes schedule task from computer
                invoke-command -computername $pc -ScriptBlock {
                    Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false
                    Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Updating GUI for Windows Update on $env:computername"
                    Start-Process "C:\windows\system32\usoclient.exe" -ArgumentList "StartInteractiveScan" -Wait
                }
                # rename update log
                $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                #Rename is blocked, because its still in use by the update process
                Rename-Item \\$pc\c$\Windows\PSWindowsUpdate.log -NewName "PSWindowsUpdate-$date.log" -ErrorAction SilentlyContinue
                #Update Frontend / GUI to reflect current Updates to User and notify WSUS about changes
                Start-Process "c:\windows\system32\usoclient" -ArgumentList "startscan" -Wait
                #Restarts the remote PC and waits until its fully started again. Timeout 5mins, Check occurs every 10secs
                Write-Host -ForegroundColor Gray "[$(get-date -Format "HH:mm:ss")] Updates Total: $updatenumber, Installed: $installednumber, Failed: $Failednumber"
                if($updatenumber -eq $Failednumber){Write-Host -ForegroundColor RED "[$(get-date -Format "HH:mm:ss")] Failed to install all remaining [$Failednumber] Updates, reboot aborted";return}
                Restart-Computer -Wait -For "WinRM" -ComputerName $pc -Protocol WSMan -Force -Timeout 300 -Delay 10
            }
        }until($updates -eq $null)

        Write-Host -ForegroundColor Green "[$(get-date -Format "HH:mm:ss")] Windows is now up to date on $pc"
    }
}

$servers = @(
     "NameOfYourServer1",
     "NameOfYourServer2"
)
Install-WUonServers -Computer $servers
Write-Host 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
