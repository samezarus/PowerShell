# sameza
#
# В оснастке Крипто ПРО удалить все считыватели кроме "Реестр" что бы максимально сократить количество окон выскакивающих при добавлении
#
# Ветка на форуме https://www.cryptopro.ru/forum2/default.aspx?g=posts&t=17557

Clear

$certsLocation = 'cert:\CurrentUser\My'
#$certsFolder   = $PSScriptRoot
$certsFolder   = 'c:\temp'
$mypwd         = ConvertTo-SecureString -String 'qwerty' -Force -AsPlainText

$zabbixSender = 'C:\Zabbix Agent\bin\win64\zabbix_sender.exe'
$zabbixServer = '192.168.10.209'
$senderHost   = 'vl20-cert'
$trapName     = 'certs_expiring'


#################################################################################################################################
#################################################################################################################################
#################################################################################################################################
function clear_string($s)
{
    # Очишает строку от спецсимволов

    $result = $s.replace("`0", '')
    $result = $result.replace("`a", '')
    $result = $result.replace("`b", '')
    $result = $result.replace("`r", '')
    $result = $result.replace("`t", '')
    $result = $result.replace("`v", '')
    $result = $result.replace("`f", '')

    return $result
}
#################################################################################################################################
function get_current_user_certs()
{
    return Get-ChildItem -Path $certsLocation -Recurse #
    
    #foreach ($item in $certsList)
    #{
        #$item | Get-Member -memberType *property
    #}
}
#################################################################################################################################
function export_certs_from_current_user($certs, $certsFolder)
{
    foreach ($item in $certs)
    {
        Export-PfxCertificate -Cert $item -FilePath ($certsFolder+'\'+$item.SerialNumber+'.pfx') -Password $mypwd
    }
}
#################################################################################################################################
function find_cert_by_serial_number($certs, $serialNumber)
{
    $result = $false
    
    foreach ($item in $certs)
    {
        if ($item.SerialNumber -eq $serialNumber)
        {
            $result = $true
            return $result
            break
        }
    }

    return $result
}
#################################################################################################################################
function import_certs_to_current_user($certs, $certsFolder)
{
    $certsFileList =  Get-ChildItem $certsFolder

    foreach ($item in $certsFileList)
    {
        $indexPfx = $item.Name.IndexOf('.pfx')
        if ($indexPfx -gt -1)
        {
            $certFile = $certsFolder+'\'+$item.Name
            $serialNumber = $item.Name.split('.')[0]
            
            $fr = find_cert_by_serial_number -certs $certs -serialNumber $serialNumber

            if ($fr -ne $true)
            {
                Import-PfxCertificate -FilePath $certFile -CertStoreLocation $certsLocation -Password $mypwd -Exportable
            }
        }
    }
}
#################################################################################################################################
function remove_all_certs($certs)
{
    # Удаляем открытые ключи
    foreach ($item in $certs)
    {

        $item | Remove-Item -Force -Recurse 
    }

    # Удаляем приватные ключи
    #   Получение имени текущей учетной записи
    $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    #   Получение SID текущего пользователя
    $objUser = New-Object System.Security.Principal.NTAccount($userName)
    $userSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])

    $cryptoProUsers = 'HKLM:\SOFTWARE\Wow6432Node\Crypto Pro\Settings\USERS\'
    $userKeysPath = $cryptoProUsers + $userSID + "\Keys"

    $userKeys = Get-ChildItem -Path $userKeysPath

    foreach ($item in $userKeys)
    {
        #$item.Name
        reg.exe delete $item.Name /f
    }
}
#################################################################################################################################
function get_subject_params($subject)
{
    # Возвращает элементы как минимум со следующими параметрами:
    # ОГРН
    # СНИЛС
    # ИНН
    # E      - Эл. почта
    # CN     - Название организации
    # T      - Должность
    # SN     - Фамилия
    # G      - Имя, Отчество
    # S      - Край/Область
    # L      - Город
    # STREET - Улица  

    $s = $subject + ','
    $s = clear_string -s $s

    $dict = @{}

    $a = 0
    For ($i=0; $i -le $s.Length +1; $i++)
    {
        if ($s[$i] -eq '=')
        {
            $key = $s.Substring($a, $i -$a)
            $b = $i +1
        }
        
        if ($s[$i] -eq ',')
        {
            $value = $s.Substring($b, $i -$b)
            if ($value[0] -eq '"')
            {
                if ($s[$i -1] -ne '"')
                {
                    continue
                }
            }   
            $a = $i+2 

            $dict.Add($key, $value)
        }
    }

    return $dict
}
#################################################################################################################################
function get_subjects_params($certs)
{
    $paramsList = New-Object System.Collections.ArrayList

    foreach ($item in $certs)
    {
        $params = get_subject_params -subject $item.Subject
        $x = $paramsList.Add($params)
    }
    
    return $paramsList
}
#################################################################################################################################
function get_expirings_days($certs)
{
    $dateMask = 'yyyy.MM.dd'

    $paramsList = New-Object System.Collections.ArrayList
   
    foreach ($item in $certs)
    {   
        $eDate = $item.NotAfter | Get-Date -Format $dateMask  # Действует до
        $exp = New-TimeSpan -End $eDate
        $expDays = $exp.Days

        $params = get_subject_params -subject $item.Subject
        $INN = $params['ИНН']
        $CN  = $params['CN']
        $SN  = $params['SN']
        $G   = $params['G']

        $dict = @{}

        $dict.Add('expDays', $expDays) 
        $dict.Add('INN'    , $INN)
        $dict.Add('CN'     , $CN)
        $dict.Add('SN'     , $SN)
        $dict.Add('G'      , $G)

        $x = $paramsList.Add($dict)
        #'--------------'
    }

    return $paramsList
}
#################################################################################################################################
function ConvertTo-Encoding ([string]$From, [string]$To)
{  
    # https://social.technet.microsoft.com/Forums/ru-RU/9f84de0e-f68c-446a-bc07-f079d98e7b7b/powershell-108710861087108810721074108010901100?forum=scrlangru
    Begin
    {  
        $encFrom = [System.Text.Encoding]::GetEncoding($from)  
        $encTo = [System.Text.Encoding]::GetEncoding($to)  
    }  
    Process
    {  
        $bytes = $encTo.GetBytes($_)  
        $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)  
        $encTo.GetString($bytes)  
    }  
} 
#################################################################################################################################
function send_to_zabbix ($tName, $msg)
{
    $msg = $msg | ConvertTo-Encoding 'windows-1251' 'utf-8'
    $msg = clear_string -s $msg
    $msg = $msg.Replace('"', '')
    &$zabbixSender -z $zabbixServer -s $senderHost -k $tName -o $msg
}
#################################################################################################################################
function find_expiring($certs, $interval, $trapName)
{
    # если $trapName пустая строка, то не посылаем в zabbix
    $expiringsList = get_expirings_days -certs $certs
    foreach ($item in $expiringsList)
    {
        $expDays = $item['expDays']
        if (($expDays -le $interval) -and ($expDays -gt -1))
        {
            $msg = 'Организация: '+$item['CN'] + '; ИНН: ' + $item['INN'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] + '; Заканчивается через: ' + $expDays
            
            if (($trapName -ne '') -and ($trapName -ne $null))
            {
                send_to_zabbix -tName $trapName -msg $msg
            }

            $msg
            '------------------------------------------------'
        }
    }
}
#################################################################################################################################
#################################################################################################################################
#################################################################################################################################

# --- Получаем список сертификатов текущего пользователя
$certs = get_current_user_certs

if ($args.Count -gt 0)
{
    switch -Exact($args[0])
    {
        '-export'
        {
            # --- Экспортируем все сертификаты текущего пользователя в папку
            export_certs_from_current_user -certs $certs -certsFolder $certsFolder
        }

        '-import'
        {
            # --- Импортируем все сертификаты текущему пользователю из папки
            import_certs_to_current_user -certs $certs -certsFolder $certsFolder
        }

        '-clear'
        {
            # --- Удаляем все сертификаты у текущего пользователя
            remove_all_certs -certs $certs
        }

        '-expiring'
        {
            # --- Выводим сертификаты у которых истекает срок действия через $interval ($args[1])
            if ($args.Count -gt 1)
            {
                $interval = $args[1]
                $tName    = $args[2]
            }
            else
            {
                $interval = 7
                $tName    = ''
            }

            find_expiring -certs $certs -interval $interval -trapName $tName
        }
    }
}

# --- Вывод параметров из Subject
#$paramsList = get_subjects_params -certs $certs
#foreach ($item in $paramsList)
#{
#    $item
#    '--------'
#}

# --- Вывод количества дней действия сертификатов
#$expiringsList = get_expiring_days -certs $certs
#foreach ($item in $expiringsList)
#{
#    $item
#    '--------'
#}

# --- Отправка сообщения/значения трапу в zabbix 
#send_to_zabbix -trapName $trapName -msg 'тест'

#find_expiring -certs $certs -interval $interval -trapName $trapName
#find_expiring -certs $certs -interval 14 -trapName ''
