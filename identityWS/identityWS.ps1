# 
#
#
Clear

'START'

$ou_array = 'OU=vl20 Vladivostok Derevenskaya14,OU=Office,OU=Ws,DC=severotorg,DC=local'

$ws_count = 0

foreach ($array_item in $ou_array)
{
    $pc_list = Get-ADComputer -Filter * -SearchBase $array_item
    
    foreach ($item in $pc_list)
    {
        $ws_count +=1

        $osStaus = 'OnLine'

        $pc_status = Test-Connection -computername $item.Name -quiet

        $pcName = $item.Name.ToString()

        if ($pc_status -eq $True) 
        {
            $srv_count.ToString() +' - '+ $pcName + ' - ' +$osStaus
            
            $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $pcName -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace
            '  C size: ' + '{0:N2} GB' -f ($disk.Size / 1gb)
            '  C free: ' + '{0:N2} GB' -f ($disk.FreeSpace / 1gb)
            ' '

            $usersPath = '\\'+$pcName+'\c$\users'
            gci -force $usersPath -ErrorAction SilentlyContinue | ? { $_ -is [io.directoryinfo] } | % {
                $len = 0
                gci -recurse -force $_.fullname -ErrorAction SilentlyContinue | % { $len += $_.length }
                '  ' + $_.fullname, ': {0:N2} GB' -f ($len / 1Gb)
            }
        }
        else
        {
            $osStaus = 'OFFLine'
            $srv_count.ToString() +' - '+ $pcName + ' - ' + $osStaus

            $osStaus
        }


        
        '------------------'
    }
}

'END'
