# Скрипт ищет на сервере процессы, по которым его можно идентифицировать
#
#
Clear-Host
# ------------------------------------------------------------------------------------------------

'START'

$ou_array = 'OU=Domain Controllers,DC=severotorg,DC=local',
            'OU=Office,OU=Srv,DC=severotorg,DC=local',
            'OU=Shop,OU=Srv,DC=severotorg,DC=local'        

$servises = 'Transport',
            'UkmService'

$srv_count = 0

foreach ($array_item in $ou_array)
{
    $pc_list = Get-ADComputer -Filter * -SearchBase $array_item
    
    foreach ($item in $pc_list)
    {
        $srv_count +=1

        $pc_status = Test-Connection -computername $item.Name -quiet

        if ($pc_status -eq $True) 
        {
            $srv_count.ToString() +' - '+ $item.Name + ' - OnLine'
            
            foreach ($sub_item in $servises)
            {
                $servise = Get-Service -computername $item.Name | Where-Object {$_.Name -like $sub_item}

                if ($servise.Name -ne '')
                {
                    '  >'+$servise.Name+'<'
                }
            }
        }
        else
        {
            $srv_count.ToString() +' - '+ $item.Name + ' - OffLine'
        }
        
        '------------------'
    }
}

'END'