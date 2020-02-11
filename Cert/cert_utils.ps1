clear

$exportFolder = $PSScriptRoot+'\export_certs' # Каталог выгрузки сертификатов
$importFolder = $PSScriptRoot+'\import_certs' # Каталог загрузки сертификатов
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
    $result = $result.replace("`f", '')
    $result = $result.replace('"', '')

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

    $userKeys = Get-ChildItem -Path $userKeysPath

    foreach ($item in $userKeys)
    {
	    $keyPath = $item.Name
	
	    $ts = time_stamp
        $exportFilePath = $exportFolder+'\' + $ts + '.reg'
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
function remove_by_fio($fio)
{
    # отложено !!!
    # Удаление сертификатов по фамилии владельца

    $serts = get_certs_list #

    foreach ($sert in $serts)
    {
        $sert.Subject
        '-------------------------------'
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

    $regFilesList = Get-Childitem -Path $importFolder
    
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
    # При первом запуске !!! Если не было экспорта, то его надо обязательно сделать !!!!!!!!
    
    $serts = get_certs_list #

    $x = Remove-Item $exportFolder -Recurse -Force

    export_certs

    $regFilesList = Get-Childitem -Path $exportFolder # Работаем с экспортироваными ключами данным скриптом

    if ((Test-Path $exportFolder) -eq $true)
    {
        $x = Remove-Item $sortFolder -Recurse -Force
        $x = create_folder -folderName $sortFolder

        foreach ($regFile in $regFilesList)
        {
            remove_all_certs
        
            import_cert_from_reg_file -regFile $regFile

            $sert = get_certs_list # в списке будит только один сертификат, который импортировали выше

            $subject = get_subject_params -subject $sert[0].Subject

            $fio = $subject['SN'] +' '+ $subject['G']

            $fioFolder = $sortFolder + '\' + $fio + '\' + $subject['CN']

            $x = create_folder -folderName $fioFolder

            $eDate = $sert.NotAfter | Get-Date -Format 'yyyy.MM.dd'

            $newName = $fioFolder+'\'+$subject['CN'] + ' ' + $subject['ИНН'] + ' '+$eDate+'.reg'
            $newName = (clear_string -s $newName)
            
            $x = Copy-Item -Path $regFile.FullName -Destination $newName
        }

        remove_all_certs

        import_certs_from_reg_files
    }
}
#################################################################################################################################
# -------------------------------------------------------------------------------------------------------------------------------
#################################################################################################################################
function get_launcher_fio_dir()
{
    $result = $PSScriptRoot

    Set-Location $PSScriptRoot

    $result = '..\'+$PSScriptRoot

    Set-Location -Path '..\'

    $result = Get-Location

    Set-Location $PSScriptRoot

    $result = $result.Path + '\Launcher\'

    return $result
}
#################################################################################################################################
function find_in_file($fileName, $searchString, $findFl)
{
    # $findFl - флаг алгоритма поиска 0 - жёсткое соответствие строки, 1 - содержится в строке
    #
    $result = $false
    
    $stringList = Get-Content $fileName
    foreach ($item in $stringList)
    {
        if($findFl -eq 0)
        {
            if($item -eq $searchString)
            {
                $result = $true
                break
            }
        }
        #
        if($findFl -eq 1)
        {
            if($item.indexOf($searchString) -ne -1)
            {
                $result = $true
                break
            }
        }
    }

    return $result
}
#################################################################################################################################
function launcher_conf()
{
    # Конфигурирование лаунчера

    $certs = get_certs_list
    
    $launcherFioDir = get_launcher_fio_dir

    # Удаляем все файлы
    $filesList = Get-Childitem -Path $launcherFioDir
    foreach ($item in $filesList)
    {
        Remove-Item -path $item.FullName -Recurse -Force
    }

    
    foreach ($item in $certs)
    {
        $params = get_subject_params -subject $item.Subject
        $fio    = $params['SN'] +' ' +$params['G']
        
        # --- Создаём структуру каталогов с ФИО владельцев сертификатов и эксмортируем в них сертификаты
        $fioDir = $cert_sort +$fio
        
        #New-Item -Path $fioDir -ItemType "directory" -Force
        #export_cert_from_current_user -cert $item -certsFolder $fioDir

        # --- Создаём структуру каталогов по ФИО
        $fioDir = $launcherFioDir + 'fio\' + $fio
        $x = New-Item -Path $fioDir -ItemType "directory" -Force
        
        $orgName = $params['CN'].replace('"', '') # Наименование организации
        $orgFile = $fioDir + '\' + $orgName + '.txt'
        $x = New-Item -Path $orgFile -ItemType File -Force
        
        # --- Создаём/дополняем ФИО в файл с ip-адресам серверов 

        $serversFile = $PSScriptRoot + '\servers.txt'

        if ((Test-Path $serversFile) -ne $true)
        {
            $x =  New-Item -Path $serversFile -ItemType File -Force
        }

        $fio = $fio +'='
        
        $findFl = find_in_file -fileName $serversFile -searchString $fio -findFl 1
        if ($findFl -eq $false)
        { 
            $fio | out-file -filepath $serversFile -Append   
        }
        
    }
}
#################################################################################################################################
#################################################################################################################################
#################################################################################################################################

$x = create_folder -folderName $exportFolder
$x = create_folder -folderName $importFolder
$x = create_folder -folderName $sortFolder


if ($args.Count -gt 0)
{
    switch -Exact($args[0])
    {
        '-export'
        {
            # --- Экспортируем все сертификаты текущего пользователя в каталог $exportFolder
            export_certs
        }

        '-import'
        {
            # --- Импортируем все сертификаты текущему пользователю из каталога $importFolder
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
            # --- !!!!!!!!!!!!!!!!!!!! Очень осторожно с этим параметром
            # --- Сортируем сертификаты по владельцу в каталог $sortFolder
            
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

#launcher_conf
