clear

$certsFolder = $PSScriptRoot+'\certs' # Каталог выгрузки сертификатов
#
$sortFolder = $PSScriptRoot+'\sort_certs' # Каталог для отсортированных сертификатов
#
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name # Получение имени текущей учетной записи
#
# Получение SID текущего пользователя
$objUser = New-Object System.Security.Principal.NTAccount($userName)
$userSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
#
# Получение пути к ветке с ключами  текущего пользователя
$cryptoProUsers = 'HKLM:\SOFTWARE\Wow6432Node\Crypto Pro\Settings\USERS\'
$userKeysPath   = $cryptoProUsers + $userSID + '\Keys'
#
$certsLocation = 'cert:\CurrentUser\My' # Путь к хранилищу "Личное"
#
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
function time_stamp()
{
    $result = (Get-Date -format 'yyyyMMddHHmmssfff').ToString()
    return $result
}
#################################################################################################################################
function create_folder($folderName)
{
    $result = $false

    if (($folderName -ne '') -and ($folderName -ne $null))
    {
        $x      = New-Item -Path $folderName -ItemType 'directory' -Force
        $result = $true
    }

    return $result
}
#################################################################################################################################
function export_certs()
{
    # Экспорт сертификатов текущего пользователя из ветки реестра в каталог $certsFolder

    $x = create_folder -folderName $certsFolder

    $userKeys = Get-ChildItem -Path $userKeysPath

    foreach ($item in $userKeys)
    {
	    $keyPath = $item.Name
	
	    $ts = time_stamp
        $exportFilePath = $certsFolder+'\' + $ts + '.reg'
	    reg.exe export $keyPath $exportFilePath "/y"
    }
}
#################################################################################################################################
function get_certs_list()
{
    # Получение списка сертификата из хранилища "Личное"

    return Get-ChildItem -Path $certsLocation -Recurse
}
#################################################################################################################################
function remove_all_certs()
{
    # Удаляем все публичные и приватные ключи

    $certs = get_certs_list

    # Удаляем открытые ключи
    foreach ($item in $certs)
    {
        $item | Remove-Item -Force -Recurse 
    }

    # Удаляем приватные ключи

    $userKeys = Get-ChildItem -Path $userKeysPath

    foreach ($item in $userKeys)
    {
        #$item.Name
        reg.exe delete $item.Name /f
    }
}
#################################################################################################################################
function import_cert_from_reg_file($regFile)
{
    $keyFilePath = $regFile.FullName
    $keyFile     = Get-Content $keyFilePath
    $findTerm    = '[HKEY_LOCAL_MACHINE'

    foreach ($line in $keyFile)
    {
        $findIndex = $line.IndexOf($findTerm)
        if ($findIndex -gt -1)
        {
            $temp   = $line.split('\')
            $oldSID = $temp[6]
            
            $orgStr = $temp[8]
            $orgStr = $orgStr.split(']')
            $orgStr = $orgStr[0]

            break        
        }
    }

    if ($oldSID -ne '')
    {        
        # Меняем сид на сид текущего пользователя 
        if ($oldSID -ne $userSID.value) 
        {
            $keyFile | ForEach-Object {$_ -replace $oldSID, $userSID.value} | Set-Content $keyFilePath
        }

        # Устанавливаем сертификат закртой части ключа в реестр
        reg import $keyFilePath

        # Устанавливаем сертификат открытой части ключа в реестр
        $ts = time_stamp
        $cont = '\\.\REGISTRY\' + $orgStr
        &'C:\Program Files\Crypto Pro\CSP\csptest.exe' -property -cinstall -cont $cont
    }
}
#################################################################################################################################
function import_certs_from_reg_files()
{
    # Импортирование всех сертификатов текущему пользователю из каталога $certsFolder

    $regFilesList = Get-Childitem -Path $certsFolder
    
    foreach ($item in $regFilesList)
    {
        import_cert_from_reg_file -regFile $item
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
        $encTo   = [System.Text.Encoding]::GetEncoding($to)  
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
function find_expiring($interval, $trapName)
{
    # если $trapName пустая строка, то не посылаем в zabbix

    $certs = get_certs_list

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
function certs_sort()
{
    $serts = get_certs_list #

    $regFilesList = Get-Childitem -Path $certsFolder #

    Remove-Item $sortFolder -Recurse -Force
    create_folder -folderName $sortFolder

    foreach ($regFile in $regFilesList)
    {
        remove_all_certs
        
        import_cert_from_reg_file -regFile $regFile

        $sert = get_certs_list # в списке будит только один сертификат, который импортировали выше

        $subject = get_subject_params -subject $sert[0].Subject

        $fio = $subject['SN'] +' '+ $subject['G']

        $fioFolder = $sortFolder+'\'+$fio

        create_folder -folderName $fioFolder

        Copy-Item -Path $regFile.FullName -Destination $fioFolder
    }

    remove_all_certs

    import_certs_from_reg_files
}
#################################################################################################################################
#################################################################################################################################
#################################################################################################################################

if ($args.Count -gt 0)
{
    switch -Exact($args[0])
    {
        '-export'
        {
            # --- Экспортируем все сертификаты текущего пользователя в каталог
            export_certs
        }

        '-import'
        {
            # --- Импортируем все сертификаты текущему пользователю из каталога
            import_certs_from_reg_files
        }

        '-clear'
        {
            # --- Удаляем все сертификаты у текущего пользователя
            remove_all_certs
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

            find_expiring -interval $interval -trapName $tName
        }

        '-sort'
        {
            # --- Сортируем сертификаты по владельцу
            
            certs_sort
        }
    }
}

# --- 1
#export_certs

# --- 2
#remove_all_certs

# --- 3
#import_certs_from_reg_files

# --- 4
#certs_sort
