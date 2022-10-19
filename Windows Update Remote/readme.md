# Windows Update Remote
![screenshot of the script performing updates]("Windows Update Script - Screenshot2.JPG")

allows to execute updates on multiple computers, only one computer is processed at any time
uses/installs Powershell modul PSWindowsUpdate to check and perform updates
detects updates on the regarding computer and installs them
server get rebooted, once all currently available updates got installed
after the reboot the script waits for the computer to return, to check for new updates
this continues until all updates are installed or all updates have the status failed
run this script with active directory administrator permissions, otherwise it might fail
some domains or computers might be not prepared for powershell remoting, check the first comments to prepare them
i can not give any support or warranty. If an updates breaks your computer, you have to deal with that yourself.
