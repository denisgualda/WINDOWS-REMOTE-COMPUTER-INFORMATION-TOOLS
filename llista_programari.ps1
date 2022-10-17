$computers = $args[0]

$array = @()

foreach($pc in $computers){

    $computername=$pc
    Write-Host "nom de lordinador: $computername"



    $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 



    $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computername) 



    $regkey=$reg.OpenSubKey($UninstallKey) 

    #Retrieve an array of string that contain all the subkey names

    $subkeys=$regkey.GetSubKeyNames() 

    #Open each Subkey and use GetValue Method to return the required values for each

    foreach($key in $subkeys){

        $thisKey=$UninstallKey+"\\"+$key 

        $thisSubKey=$reg.OpenSubKey($thisKey) 

        $obj = New-Object PSObject

        #$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computername

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))


        $array += $obj

    } 

}

$array | Where-Object { $_.DisplayName } | Select-Object  DisplayName, DisplayVersion | Format-Table -auto