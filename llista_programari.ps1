
#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$ComputerName = $args[0]
#$ComputerName = "localhost"

ForEach ($Computer in $ComputerName) {
#[System.Windows.Forms.Messagebox]::Show($Computer)	


$lmKeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        $lmReg = [Microsoft.Win32.RegistryHive]::LocalMachine
        $cuKeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        $cuReg = [Microsoft.Win32.RegistryHive]::CurrentUser
 $masterKeys = @()
        $remoteCURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($cuReg,$Computer)
        $remoteLMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($lmReg,$Computer)
        foreach ($key in $lmKeys) {
            $regKey = $remoteLMRegKey.OpenSubkey($key)
            foreach ($subName in $regKey.GetSubkeyNames()) {
                foreach($sub in $regKey.OpenSubkey($subName)) {
                    $masterKeys += (New-Object PSObject -Property @{
                        "Name" = $sub.GetValue("displayname")
                        "Version" = $sub.GetValue("DisplayVersion")

                    })
                }
            }
        }
        foreach ($key in $cuKeys) {
            $regKey = $remoteCURegKey.OpenSubkey($key)
            if ($regKey -ne $null) {
                foreach ($subName in $regKey.getsubkeynames()) {
                    foreach ($sub in $regKey.opensubkey($subName)) {
                        $masterKeys += (New-Object PSObject -Property @{
                            "Name" = $sub.GetValue("displayname")
                            "Version" = $sub.GetValue("DisplayVersion")
                        })
                    }
                }
            }
        }
        $woFilter = {$null -ne $_.name -AND $_.SystemComponent -ne "1" -AND $null -eq $_.ParentKeyName}
        $props = 'Name','Version'
        $masterKeys = ($masterKeys | Where-Object $woFilter | Select-Object $props | Sort-Object Name)
        $masterKeys
        
        

}