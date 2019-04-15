# Ставит зелёные галки напроти ключей в usbredirector клиентская часть
#
# usbrdrltsh.exe -help

# ------------------------------------------------------------------------------------------------------------------

function findValue ([string]$str, [string]$findKey)
{
    #$str
    $result = ''
    $str_len = $str.Length
    if ($str_len -gt 0)
    {
        $pos1 = -1
        For ($i=0; $i -lt $str_len; $i++) 
        {
            if ($str[$i] -eq ' ' -and $str[$i+2] -eq ' ' -and  $str[$i+3] -ne ' ')
            {
                $result = $str.Substring($pos1+1, $i-$pos1-1)
                $pos1 = $i+2

                $pos = $result.indexOf($findKey)
                if ($pos -gt -1)
                {
                    $result = $result.Substring($findKey.Length+2)
                    return $result
                }
                else
                {
                    $result = ''
                }
            }
        }
        
        $result = $str.Substring($pos1+1, $str_len-$pos1-1) # последняя пара ключ:значение в строке
        
        $pos = $result.indexOf($findKey)
        if ($pos -gt -1)
        {
            $result = $result.Substring($findKey.Length+2)
            return $result
        }
        else
        {
            $result = ''
        }
    }
    return $result
}

#

function GetList()
{
    return .\usbrdrltsh.exe -list
}

#

function ConnectKey ($server_u, $vid_dev, $pid_dev)
{
    return .\usbrdrltsh.exe -connect -server $server_u -vid $vid_dev -pid $pid_dev
}

# 

function ConnectAllKeys()
{
    $res = GetList

    $srv_flg = 0

    ForEach($str in $res)
    {    
        # Получаем имя сервера
        $tmp_str = 'USB server at' 
        $pos = $str.indexOf($tmp_str)
        if ($pos -gt -1)
        {
            $server_u = $str.Substring($pos + $tmp_str.Length + 1)

            $srv_flg = 1

            'Server name: ' +$server_u
        }

        # Получаем имя устройства
        $pos = $str.indexOf($tmp_str)
        if ($str[9] -eq ':')
        {
            $dev_name = $str.Substring(11)
            'Device name: '+$dev_name
        }

        # Параметры устройства Vid/Pid/Serial
        $tmp_str = 'Vid:' 
        $pos = $str.indexOf($tmp_str)
        if ($pos -gt -1)
        {
            #$str
            $vid_dev    = findValue $str 'Vid'
            $pid_dev    = findValue $str 'Pid'
            $serial_dev = findValue $str 'Serial'
            'Vid: '    + $vid_dev
            'Pid: '    + $pid_dev
            'Serial: ' + $serial_dev
        }
    
        # Mode/Status
        $tmp_str = 'Mode:' 
        $pos = $str.indexOf($tmp_str)
        if ($pos -gt -1)
        {
            if ($srv_flg -eq 1) # Параметры сервера 
            {
                $Mode_srv   = findValue $str 'Mode'
                $Status_srv = findValue $str 'Status'
                'Mode: '   + $Mode_srv
                'Status: ' + $Status_srv
                $srv_flg = 0
            }
            else # Параметры устройства
            {
                #$str
                $Mode_dev   = findValue $str 'Mode'
                $Status_dev = findValue $str 'Status'
                'Mode: '   + $Mode_dev
                'Status: ' + $Status_dev

                if ($Status_dev -eq 'available for connection')
                {
                    $connect_device = ConnectKey $server_u $vid_dev $pid_dev
                    $connect_device
                }
            }
            '---------------------------'
        }
    }
}

# ------------------------------------------------------------------------------------------------------------------

Clear-Host
$command_param = ''

if ($args.Length -gt 0)
{
    $command_param = $args[0]
}

$usbrcl  = 'C:\Program Files\USB Redirector Client'
$usbrcn  = 'usbrdrltsh.exe'
$usbrcfn = $usbrcl+'\'+$usbrcn

if (Test-Path ($usbrcfn))
{
    Set-Location $usbrcl    
    
    if ($command_param -eq '-conall')
    {
        ConnectAllKeys
    }
}
