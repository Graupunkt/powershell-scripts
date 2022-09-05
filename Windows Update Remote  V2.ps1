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

    ForEach ($pc in $computer){
        #install pswindows updates module on remote machine
        if((Test-WSMan $pc -ErrorAction SilentlyContinue).wsmid -match "wsmanidentity\.xsd"){
            Invoke-Command -ComputerName $pc -Scriptblock{
                #CHECK IF NUGET AND PSWINDOWS UPDATE IS INSTALELD; IF NOT INSTALL THEM
                if((Get-PackageProvider -Name NuGet).version -lt 2.8.5.201){Install-PackageProvider -Name NuGet -RequiredVersion "2.8.5.201" -Force}
                If(-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)){Install-Module PSWindowsUpdate -Confirm:$False -Force -Repository PSGallery -WA 0}
                Import-Module PSWindowsUpdate -force

                #CHECK IF PSWindows Update was installed sucessfully
                If(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue){
                    Write-Host -ForegroundColor Green "$env:computername has PSWindowsUpdate sucessfully installed"
                }else{
                    Write-Host -ForegroundColor Red "$env:computername has Powershell Modul not installed"
                }
            }
        }else{
            Write-Host -ForegroundColor Red "$env:computername has WinRM not properly configured"
        }

        Do{
            #Reset Timeouts
            $connectiontimeout = 0
            $updatetimeout = 0
            #starts up a remote powershell session to the computer
            do{
                $session = New-PSSession -ComputerName $pc
                Write-Host -ForegroundColor Green "reconnecting remotely to $pc"
                sleep -seconds 10
                $connectiontimeout++
            } until ($session.state -match "Opened" -or $connectiontimeout -ge 10)
            #retrieves a list of available updates
            Write-Host -ForegroundColor Green "Checking for new updates available on $pc"
            $updates = invoke-command -session $session -scriptblock {Get-wulist -verbose}
            #counts how many updates are available
            $updatenumber = ($updates.kb).count
            #if there are available updates proceed with installing the updates and then reboot the remote machine
            if ($updates -ne $null){
                #remote command to install windows updates, creates a scheduled task on remote computer
                invoke-command -ComputerName $pc -ScriptBlock { Invoke-WUjob -ComputerName localhost -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\PSWindowsUpdate.log" -Confirm:$false -RunNow}
                #Show update status until the amount of installed updates equals the same as the amount of updates available
                Start-Sleep -Seconds 30
                do {$updatestatus = Get-Content \\$pc\c$\PSWindowsUpdate.log
                    Write-Host -ForegroundColor Green "Currently processing the following update: (waiting 20mins for timeout)"
                    Get-Content \\$pc\c$\PSWindowsUpdate.log | select-object -last 1
                    Start-Sleep -Seconds 10
                    $ErrorActionPreference = ‘SilentlyContinue’
                    $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count
                    $Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count
                    $ErrorActionPreference = ‘Continue’
                    $updatetimeout++
                }until ( ($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 720)
                #restarts the remote computer and waits till it starts up again
                Write-Host -ForegroundColor Green "restarting remote computer and waiting 5mins until continueing to process further more updates"
                #removes schedule task from computer
                invoke-command -computername $pc -ScriptBlock {Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false}
                # rename update log
                $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
                #Rename is blocked, because its still in use by the update process
                Rename-Item \\$pc\c$\PSWindowsUpdate.log -NewName "WindowsUpdate-$date.log" -ErrorAction SilentlyContinue
                Restart-Computer -Wait -For "WinRM" -ComputerName $pc -Force
            }
        }until($updates -eq $null)

        Write-Host -ForegroundColor Green "Windows is now up to date on $pc"
    }
}

$servers = @("ubiquiti.tepe.local","ubiquiti.tepe.local")
Install-WUonServers -Computer $servers
