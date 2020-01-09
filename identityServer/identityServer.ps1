# Скрипт ищет на сервере процессы, по которым его можно идентифицировать
#
#
Clear

$log = $PSScriptRoot +'\log.html'

# ------------------------------------------------------------------------------------------------
Function toLog ($str)
{
    Add-Content -Path $log -Value $str
}

# ------------------------------------------------------------------------------------------------

'START'

$ou_array = 'OU=Domain Controllers,DC=severotorg,DC=local',
            'OU=Srv,DC=severotorg,DC=local'

$servises = 'UkmService',
            'Transport',
            'Zabbix Agent'

$srv_count = 0

$colorOn  ='#99FF99'
$colorOff ='#FF9999'

Remove-Item -Path $log

toLog '<html>'
toLog '<body>'
toLog '<table border="1" cellpadding="5" cellspacing="0">'

toLog '<tr>'
toLog '  <th>Сервер</td>'
toLog '  <th>Статус</td>'
#
toLog '  <th>УКМ</td>'
toLog '  <th>УТМ</td>'
toLog '  <th>Zabbix</td>'
#
toLog '  <th>Материнская плата</td>'
toLog '  <th>ОС</td>'
toLog '  <th>Версия ОС</td>'
toLog '  <th>Дата Установки</td>'
toLog '  <th>Последний запуск</td>'
toLog '  <th>RAM</td>'
toLog '</tr>'

foreach ($array_item in $ou_array)
{
    $pc_list = Get-ADComputer -Filter * -SearchBase $array_item
    
    foreach ($item in $pc_list)
    {
        $srv_count +=1

        $pc_status = Test-Connection -computername $item.Name -quiet

        $pcName           = $item.Name.ToString()

        toLog '<tr>'
        toLog ('  <td>'+$pcName+'</td>')

        if ($pc_status -eq $True) 
        {
            $osInf = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $item.Name 

            #$item | format-list -property *
            #$osInf | format-list -property *
            
            
            $osStaus          = 'ONLine'
            $osCaption        = $osInf.Caption.ToString()
            $osVersion        = $osInf.Version.ToString()
            $osInstallDate    = $osInf.InstallDate.ToString()  
            $osLastBootUpTime = $osInf.LastBootUpTime.ToString()
            $osRAM            = ([math]::Round($osInf.TotalVisibleMemorySize/1048576, 2)).ToString()

            foreach ($sub_item in $servises)
            {
                $servise = Get-Service -computername $item.Name | Where-Object {$_.Name -like $sub_item}

                if ($servise.Name -eq 'UkmService')
                {
                    $servUKM = '+'
                }
                if ($servise.Name -eq 'Transport')
                {
                    $servUTM = '+'
                }
                if ($servise.Name -eq 'Zabbix Agent')
                {
                    $servZabbix = '+'
                }
            }

            $srv_count.ToString() +' - '+ $pcName + ' - ' +	$osStaus

            toLog ('  <td bgcolor="'+$colorOn+'">'+$osStaus+'</td>')
            #
            toLog ('  <td>'+$servUKM +'</td>')
            toLog ('  <td>'+$servUTM+'</td>')

            $clr = $colorOn
            if ($servZabbix -ne '+')
            {
                $clr = $colorOff
            }
            
            toLog ('  <td bgcolor="'+$clr+'">'+$servZabbix+'</td>')
            #
            $mbInf  = Get-WmiObject -Class Win32_BaseBoard -ComputerName $pcName
            $mbName = $mbInf.Manufacturer + ' ' +$mbInf.Product + ' ' + $mbInf.Model
            toLog ('  <td>'+$mbName+'</td>')
            #
            toLog ('  <td>'+$osCaption+'</td>')
            toLog ('  <td>'+$osVersion +'</td>')
            toLog ('  <td>'+$osInstallDate +'</td>')
            toLog ('  <td>'+$osLastBootUpTime+'</td>')
            toLog ('  <td>'+$osRAM+'</td>')
            
            $osStaus          = ''
            $osCaption        = ''
            $osVersion        = ''
            $osInstallDate    = ''  
            $osLastBootUpTime = ''
            $osRAM            = ''
            #
            $servUKM    = ''
            $servUTM    = ''
            $servZabbix = ''

        }
        else
        {
            $osStaus = 'OFFLine'
            $srv_count.ToString() +' - '+ $pcName + ' - ' + $osStaus

            toLog ('  <td bgcolor="'+$colorOff+'">'+$osStaus+'</td>')

        }

        toLog '</tr>'
        
        '------------------'
    }
}

'END'

toLog '</table>'
toLog '</body>'
toLog '</html>'
