
function Open-RemoteRegistry ($computerName = "127.0.0.1") {
    #add needed assemblies
    Add-Type -AssemblyName Microsoft.VisualBasic
    Add-Type -AssemblyName System.Windows.Forms

    #start regedit
    #C:\windows\SysWOW64\regedit.exe
    Start-Process regedit -Verb runAs

    #wait for it to start
    Start-Sleep -Seconds 2

    #activate regedit window
    [Microsoft.VisualBasic.Interaction]::AppActivate("Regedit.exe")

    #send Alt F C
    [System.Windows.Forms.SendKeys]::SendWait("%AC")

    #wait for dialog
    Start-Sleep -Seconds 1

    #insert computer name and send enter
    [System.Windows.Forms.SendKeys]::SendWait("$computerName{ENTER}")

    #profit
}

Clear-Host
$equip = $args[0]

Open-RemoteRegistry ($equip)
