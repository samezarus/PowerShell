<#

telegram https://t.me/sameza

Данный скрипт предназначен для рабрты с ЭЦП в связке с Крипто-Про

Скрипт работает через передаваемые ему параметры:

    -export 
        Экспорт сертификатов текущего пользователя в папку $exportFolder

    -exportOpenKeys
        Экспортируем все открытые части сертифиеатов и их отпечатки
        Под каждую организацию создаётся своя папка с именем организации и ИНН
        Внутри папки создаются папки по дате выпуска сертификата, и внутри неё создаётся сертификат и файл со штампом сертификата
    
    -import
        Импорт сертификатов текущему пользователю из папки $importFolder
    
    -clear 
        Удаление всех личных сертификатов у текущего пользователя
    
    -expiring 
        Вывод сертификатов у которых истекает срок действия через $interval
    
    -certs_sort 
        !!! ОЧЕНЬ ОСТОРОЖНО !!! Сортирует сертификаты по ФИО

    -rar_import
        Добавляет сертификаты из папки importFolder в архив 

    -orgcsv
        Формирует CSV-файл по всем организациям

    -setspass
        Установка пароля из глобальной переменной $pass

--------------------------------------------------------------------------------------------------

Порядок работы со скриптом:
    
    Ипорт сертификатов пользователю:
        1. Создать каталог для данного скрипта "C:\temp"
        2. Закидываем в этот каталог данный скрипт
        3. Создать "C:\temp\import_certs", скопировать в него ранее экспортированные *.reg файлы
        4. Запускаем скрипт для импорта "C:\temp\cert_utils.ps1 -import"

    Экспорт личных сертификатов пользователя:
                
        "C:\temp\cert_utils.ps1 -export"

    Сортировка сертификатов:
        Результат сортировки:
            <ФИО>(каталог)
                <Название организации>(каталог)
                    <Название организации> <ИНН организации> <Дата окончания сертификата>.reg(Файл)


                
        "C:\temp\cert_utils.ps1 -certs_sort"
#>
Add-Type -AssemblyName System.Web


clear


function Get-ScriptDirectory {
    if ($psise) {
        Split-Path $psise.CurrentFile.FullPath
    }
    else {
        $global:PSScriptRoot
    }
}

$scriptDir = Get-ScriptDirectory

$exportFolder    = $scriptDir+'\export_certs'            # Каталог выгрузки сертификатов (*.reg - открытая и закрытая часть)
$exportCerFolder = $exportFolder+'\open_keys'               # Каталог выгрузки открытых ключей
$importFolder    = $scriptDir+'\import_certs'            # Каталог загрузки сертификатов
$importFolderW7  = '\\certs\c$\BuhgSoft\Certs\import_certs' # Каталог загрузки сертификатов
$sortFolder      = $scriptDir+'\sort_certs'              # Каталог для отсортированных сертификатов
$arcFolder       = $scriptDir+'\arc_certs'               # Каталг для архивирования импортируемых сертификатов
#
$serversFile = $scriptDir + '\servers.txt' # Файл со списком сопоставления серверов к ФИО руководителя
$launcherFolder = 'C:\BuhgSoft\Launcher\fio2'
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
$senderHost   = 'certs'
$trapName     = 'certs_expiring'
$trapAuto     = 'CertProcessing' # для автодескавери
#
$zabbixServer = 'http://zabbix4/zabbix/zabbix_sender/index.php?'
$zabbixHost   = 'certs'
$trapAuto     = 'CertProcessing' # для автодескавери
#
$rar = 'C:\Program Files\WinRAR\Rar.exe'
#
$pass = 'Remi1535'


function clear_string($s)
{
    # Очишает строку от спецсимволов

    $result = ''

    if (($s -ne $null) -and ($s -ne ''))
    {
        $result = $s.replace("`0", '')
        $result = $result.replace("`a", '')
        $result = $result.replace("`b", '')
        $result = $result.replace("`r", '')
        $result = $result.replace("`t", '')
        $result = $result.replace("`v", '')
        $result = $result.replace("`f", '')
        $result = $result.replace("`f", '')
        $result = $result.replace('"', '')
    }

    return $result
}


function time_stamp()
{
    $result = (Get-Date -format 'yyyyMMddHHmmssfff').ToString()
    return $result
}


function imports_to_rar()
{
    if (Test-Path $rar)
    {
        cd $arcFolder
        $fn = time_stamp + '.rar'

        $x = &$rar a -r $fn $importFolder
    }
}


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


function export_certs()
{
    # Экспорт сертификатов текущего пользователя из ветки реестра в каталог $exportFolder с именем в виде таймстемпа

    $userKeys = Get-ChildItem -Path $userKeysPath

    foreach ($item in $userKeys)
    {  
        $keyPath = $item.Name

        $lasPosSlash = $keyPath.LastIndexOf('\')
        $keyName = $keyPath.Substring($lasPosSlash +1)
	
        $exportFilePath = $exportFolder+'\' + $keyName + '.reg'
	    reg.exe export $keyPath $exportFilePath "/y"
    }
}


function export_open_keys()
{
    $certs = get_certs_list

    foreach ($item in $certs)
    {
        $orgInf  = get_subject_params -subject $item.Subject
        $orgName = $orgInf['CN']
        $orgINN  = $orgInf['ИНН']

        $fldName = $exportCerFolder +'\'+ $orgName + ' - ' + $orgINN
        $x = create_folder -folderName $fldName
        
        $dateMask = 'dd.MM.yyyy'
        $bDate = $item.NotBefore | Get-Date -Format $dateMask  # Действует с
        
        $subFldName = $fldName + '\' + $bDate
        $x = create_folder -folderName $subFldName

        $x = Export-Certificate -Cert $item -FilePath ($subFldName + '\'+$orgName+'.cer')
        
        $thumbprintFile = $subFldName + '\thumbprint.txt'
        $x = New-Item $thumbprintFile -Force
        $x = Set-Content $thumbprintFile $item.Thumbprint  
    }
}


function get_certs_list()
{
    # Получение списка сертификата из хранилища "Личное"

    return Get-ChildItem -Path $certsLocation -Recurse
}


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
        $cont = '\\.\REGISTRY\' + $orgStr
        &'C:\Program Files\Crypto Pro\CSP\csptest.exe' -property -cinstall -cont $cont
    }
}


function import_certs_from_reg_files()
{
    # Импортирование всех сертификатов текущему пользователю из каталога $certsFolder

    $regFilesList = Get-Childitem -Path $importFolder
    
    foreach ($item in $regFilesList)
    {
        import_cert_from_reg_file -regFile $item
    }
}


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


            if ($key -ne 'STREET') # Проблема с обработкой параметра STREET
            {
                $dict.Add($key, $value)
            }
        }
    }

    return $dict
}


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


            if ($key -ne 'STREET') # Проблема с обработкой параметра STREET
            {
                $dict.Add($key, $value)
            }
        }
    }

    return $dict
}


function get_expiring_days($cert)
{
    # Получение остатка дней (но так же и сопутствующую информацию), через который истечёт сертификат
    
    $dateMask = 'yyyy.MM.dd'
   
    $eDate = $cert.NotAfter | Get-Date -Format $dateMask  # Действует до
    $exp = New-TimeSpan -End $eDate
    $expDays = $exp.Days

    $params = get_subject_params -subject $cert.Subject
    $INN = $params['ИНН']
    $SNILS = $params['СНИЛС']
    $CN  = $params['CN']
    $SN  = $params['SN']
    $G   = $params['G']


    if($INN -eq "")
    {
       $INN = "123" 
    }


    $dict = @{}

    $dict.Add('expDays', $expDays) 
    $dict.Add('INN'    , $INN)
    $dict.Add('CN'     , $CN)
    $dict.Add('SN'     , $SN)
    $dict.Add('G'      , $G)
    $dict.Add('cert'   , $cert)
    $dict.Add('SNILS'  , $SNILS)

    return $dict
}


function get_cert_parent($cert)
{
    # Получаем организацию выпустившую сертификат

    #Write-Host ($cert | Format-List | Out-String)
    $params = get_subject_params $cert.Issuer
    return $params['CN']
}


function get_expirings_days($certs)
{
    # Выборка всех сертификатов(хранящихся в реестре) с остатками дней действия

    $paramsList = New-Object System.Collections.ArrayList
   
    foreach ($item in $certs)
    {   
        $x = get_expiring_days -cert $item
        $x = $paramsList.Add($x)
    }

    return $paramsList
}


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


function send_to_zabbix ($tName, $msg)
{
    if (Test-Path $zabbixSender)
    {
        $msg = $msg | ConvertTo-Encoding 'windows-1251' 'utf-8'
        $msg = clear_string -s $msg
        $msg = $msg.Replace('"', '')
        &$zabbixSender -z $zabbixServer -s $senderHost -k $tName -o $msg
    }
}


function send_to_zabbix_web ($tName, $msg)
{    
    $url = $zabbixServer + 'server=' + $zabbixHost + '&key=' + $tName + '&value=' + $msg
    $x = Invoke-WebRequest -UseBasicParsing $url
}


function find_expiring($interval, $trapName)
{
    # если $trapName пустая строка, то не посылаем в zabbix

    $certs = get_certs_list

    $expiringsList = get_expirings_days -certs $certs
    foreach ($item in $expiringsList)
    {
        $expDays = $item['expDays']
        if (($expDays -le $interval) -and ($expDays -gt 0))
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


function find_expiring2($interval, $trapName)
{
    # Суть функции с индексом 2, в том что может случится ситуация, при которой у одной организации
    # будут два сертификата, один старый с неистёкшим остатком дней, и новый перевыпущенный 
    #
    # если $trapName пустая строка, то не посылаем в zabbix

    $certs = get_certs_list

    $expiringsList    = get_expirings_days -certs $certs
    $subExpiringsList = get_expirings_days -certs $certs

    foreach ($item in $expiringsList)
    {
        if ($item['expDays'] -gt 0) # Если остаток действия дней сертификата больше 0
        {
            $expDays = $item['expDays']  

            foreach($item2 in $subExpiringsList)
            {
                if ($item['SNILS'] -eq $item2['SNILS'] )
                {
                    if ($item2['expDays'] -gt $expDays)
                    {
                        $expDays = $item2['expDays']  
                    }
                }
            }
            

            if ($expDays -le $interval)
            {
                foreach($si in $item['cert'])
                {
                    "    "+$si
                }
                
                $certParent = get_cert_parent -cert $item['cert']
                #$msg = 'Организация: '+$item['CN'] + '; ИНН: ' + $item['INN'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] + '; Заканчивается через: ' + $expDays
                $msg = 'Организация: '+$item['CN'] + '; ИНН: ' + $item['INN'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] +'; Выпустил: ' + $certParent + '; Заканчивается через: ' + $expDays            
                #$msg = 'Организация: '+$item['CN'] + '; СНИЛС: ' + $item['SNILS'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] +'; Выпустил: ' + $certParent + '; Заканчивается через: ' + $expDays            

                if (($trapName -ne '') -and ($trapName -ne $null))
                {
                    #send_to_zabbix -tName $trapName -msg $msg
                }

                $fio = $item['SN']+ ' ' + $item['G']

                # Создаём/обновляем элемент автообнаружения
                create_zabbix_auto_detect_item -inn $item['INN'] -cn $item['CN'] -fio $fio

                # Генерируем алерт для данного сертификата
                $dinamTrap = 'cert['+$item['INN']+']'

                send_to_zabbix_web -tName $dinamTrap -msg $expDays


                $msg
                '------------------------------------------------'
            }
        }
    }
}


function get_ext_sert($cert_item)
{
    # Функция получает параметры сертификата из поля Subject
    # Форматирует и ищет ИНН в сертификатах нового типа
    # Дополняет параметры свойства вне Subject
    
    $subject = $cert_item.Subject

    $new_str = ""
    $char = ""
    $wsc = 0
    
    #  Очистка строки от запятых в тексте, что бы оставить только разделяющие запятые
    for ($i=0; $i -le $subject.Length +1; $i++)
    {
        $char = $subject[$i]

        if($char -eq '"')
        {
            $wsc += 1
        }
        
        if($char -eq ',')
        {
            if(($wsc % 2) -ne 0)
            {
                $char = '.'
            }
        }

        $new_str += $char
    }

    # Избавляемся от ковычек и от пробелов в разделителях
    $new_str = $new_str.Replace('"', '')
    $new_str = $new_str.Replace(", ", ',')
    
    
    # Получение списка key=value строк
    $kv_list = $new_str.Split(",")
    
    # Формирование результирующего словаря
    $dict = @{}
    $dict.Add('cert_object', $cert_item) # Добавляем объект текущего сертификата

    $dateMask = 'yyyy.MM.dd'
    $eDate = $cert_item.NotAfter | Get-Date -Format $dateMask  # Действует до
    $exp = New-TimeSpan -End $eDate
    $dict.Add('expiring_days', $exp.Days) # Количество дней действия сертификата

    foreach($kv_item in $kv_list)
    {
        $kv = $kv_item.Split('=')
        $key = $kv[0]
        $value = $kv[1]
        $dict.Add($key, $value) 
    }

    # Очистка ИНН от нолей и находжение ИНН в новых сертификатах из параметра OID.1.2.840.113549.1.9.2
    if("OID.1.2.840.113549.1.9.2" -in $dict.Keys) # Переходной или новый сертификат
    {
        $p = $dict["OID.1.2.840.113549.1.9.2"]
        $l = $p.Split('-')
        $dict["ИНН"] = $l[0]
    }
    else # Старый сертификат
    {
        $inn = $dict["ИНН"]
        $p = -1
        for ($i=0; $i -le $inn.length +1; $i++)
        {
            if($inn[$i] -eq '0')
            {
                $p = $i 
            }
            else
            {
                $dict["ИНН"] = $inn.substring($p +1, $inn.length -$p -1)
                break
            }
        }
    }

    return $dict
}


function get_ext_serts_list()
{
    $list = @()
    
    $certs = get_certs_list
    foreach($c in $certs)
    {
        $ext = get_ext_sert $c
        $list += $ext
    }

    return $list
}


function find_expiring3($interval, $trapName)
{
    $expiringsList    = get_ext_serts_list
    $subExpiringsList = get_ext_serts_list

    foreach ($item in $expiringsList)
    {
        #$item["CN"] + " - " + $item["ИНН"] + " - " + $item["expiring_days"]
        
        if ($item['expiring_days'] -gt 0) # Если остаток действия дней сертификата больше 0
        {
            $expDays = $item['expiring_days']  

            foreach($item2 in $subExpiringsList)
            {
                if (($item['ОГРН'] -eq $item2['ОГРН']) -and ($item['ОГРНИП'] -eq $item2['ОГРНИП']))
                {
                    if ($item2['expiring_days'] -gt $expDays)
                    {
                        $expDays = $item2['expiring_days']  
                    }
                }
            }
            

            if($expDays -le $interval)
            {
                foreach($si in $item['cert'])
                {
                    "    "+$si
                }

                $ogrn = ""

                if("ОГРН" -in $item.Keys)
                {
                    $ogrn = $item["ОГРН"]
                }

                if("ОГРНИП" -in $item.Keys)
                {
                    $ogrn = $item["ОГРНИП"]
                }

                
                $certParent = get_cert_parent -cert $item['cert_object']
                #$msg = 'Организация: '+$item['CN'] + '; ИНН: ' + $item['INN'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] + '; Заканчивается через: ' + $expDays
                #$msg = 'Организация: '+$item['CN'] + '; ИНН: ' + $item['INN'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] +'; Выпустил: ' + $certParent + '; Заканчивается через: ' + $expDays            
                #$msg = 'Организация: '+$item['CN'] + '; ОГРН/ОГРНИП: ' + $item['ИНН'] + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] +'; Выпустил: ' + $certParent + '; Заканчивается через: ' + $expDays            
                $msg = 'Организация: '+$item['CN'] + '; ОГРН/ОГРНИП: ' + $ogrn + '; Руководитель: '+ $item['SN']+ ' ' + $item['G'] +'; Выпустил: ' + $certParent + '; Заканчивается через: ' + $expDays            

                if (($trapName -ne '') -and ($trapName -ne $null))
                {
                    #send_to_zabbix -tName $trapName -msg $msg
                }

                $fio = $item['SN']+ ' ' + $item['G']

                # Создаём/обновляем элемент автообнаружения
                #create_zabbix_auto_detect_item -inn $item['ИНН'] -cn $item['CN'] -fio $fio

                # Генерируем алерт для данного сертификата
                #$dinamTrap = 'cert['+$item['ИНН']+']'

                #send_to_zabbix_web -tName $dinamTrap -msg $expDays


                $msg
                '------------------------------------------------'
            }
        }
    }
}


function sort_certs()
{   
    # При первом запуске !!! Если не было экспорта, то его надо обязательно сделать !!!!!!!!
    
    $serts = get_certs_list #

    $x = Remove-Item ($exportFolder+'\*.*') -Recurse -Force

    export_certs

    $regFilesList = Get-Childitem -Path $exportFolder # Работаем с экспортироваными ключами данным скриптом

    if ((Test-Path $exportFolder) -eq $true)
    {
        #$x = Remove-Item $sortFolder -Recurse -Force

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


function get_org_name($keyRegPath)
{
    $result = ''
    
    $findTerm = '[HKEY_LOCAL_MACHINE'

    $result = $keyRegPath.split('\')
    $result = $result[8]

    return $result
}


function set_pass_to_certs($pass)
{
    $serts = Get-ChildItem -Path $userKeysPath -Recurse
   
    foreach ($item in $serts)
    {
        $orgName = get_org_name $item.Name
         if ($orgName -ne '')
         {
            $cont = '\\.\REGISTRY\' + $orgName
            &'C:\Program Files\Crypto Pro\CSP\csptest.exe' -passwd -cont $cont -change $pass -pin "" 
         }
    }
}


function ip_detect($subject)
{
    # Функция определяет является ли сертефикат выданным на ИП или нет

    $result = $false

    $p = $subject.indexof("ОГРНИП")
    if ($p -ne -1)
    {
        $result = $true
    }

    return $result
}


function create_zabbix_auto_detect_item($inn, $cn, $fio)
{
    if ($fio -eq $cn)
    {
        #$cn = 'ИП ' + $cn
        $fio = '' 
    }
    else
    {
        $fio = ' [' + $fio + ']' 
    }

    $s = '{"data":[{"{#MX}":"' + $inn + '","{#MXNAME}":"' + $cn + $fio + '"}]}'

    $msg = [System.Web.HTTPUtility]::UrlEncode($s)

    $x = send_to_zabbix_web -tName $trapAuto -msg $msg
}


function auto_detect_items_to_zabbix()
{
    foreach ($item in get_certs_list)
    {
        $res = get_subject_params -subject $item.subject

        $inn = $res['ИНН']
        $cn = $res['CN']
        $fio = $res['SN'] + ' ' + $res['G']

        $x = create_zabbix_auto_detect_item -inn $inn -cn $cn -fio $fio
    }
}

# -------------------------------------------------------------------------------------------------------------------------------


function launcher_conf()
{
    $serts = get_certs_list
    $fiosArr = @() # массив
    
    foreach ($item in $serts)
    {
        $subject = get_subject_params -subject $item.Subject
        $fio = $subject['SN'] +' '+ $subject['G']
        
        if (!($fiosArr -contains $fio))
        {
            $fiosArr += $fio
        }
    }
    
    # Удаляем все файлы
    $filesList = Get-Childitem -Path $launcherFolder
    foreach ($item in $filesList)
    {
        Remove-Item -path $item.FullName -Recurse -Force
    }

    foreach ($item in $fiosArr)
    {
        $fioFile = $launcherFolder + '\'+$item+'.txt'
        $x = New-Item -Path $fioFile -ItemType File -Force
        #'server=' | out-file -filepath $fioFile -Append 
        $item +':'
        $orgsArr = @() # массив

        foreach ($item2 in $serts)
        {
            $subject = get_subject_params -subject $item2.Subject
            $fio = $subject['SN'] +' '+ $subject['G']
            
            $orgName = $subject['CN']
            if ($item -eq $fio)
            {
                if (!($orgsArr -contains $orgName))
                {
                    $orgsArr += $orgName
                }
            }
        }
        $orgsArr | out-file -filepath $fioFile -Append
        #$orgsArr 
        #'-------------------'
    }
}


function orgs_to_csv()
{
    $serts = get_certs_list
    $fiosArr = @() # массив
    
    foreach ($item in $serts)
    {
        $subject = get_subject_params -subject $item.Subject
        $fio = $subject['SN'] +' '+ $subject['G']
        
        if (!($fiosArr -contains $fio))
        {
            $fiosArr += $fio
        }
    }

    $orgsFile = $scriptDir + '\orgs.csv'
    Remove-Item $orgsFile -Force
    $x = New-Item -Path $orgsFile -ItemType File -Force

    foreach ($item in $fiosArr)
    {
        $orgsArr = @() # массив

        foreach ($item2 in $serts)
        {
            $subject = get_subject_params -subject $item2.Subject
            $fio = $subject['SN'] +' '+ $subject['G']
            
            $orgName = $subject['CN']
            if ($item -eq $fio)
            {
                if (!($orgsArr -contains $orgName))
                {
                    $s = $item + '; '+ $orgName + '; ' + $subject['E'] +'; '+ (get_cert_parent $item2)
                    $s | out-file -FilePath $orgsFile -Append

                    $orgsArr += $orgName
                }
            }
        }
    }
}

#################################################################################################################################


clear

$x = create_folder -folderName $exportFolder
$x = create_folder -folderName $exportCerFolder
$x = create_folder -folderName $importFolder
$x = create_folder -folderName $sortFolder
$x = create_folder -folderName $arcFolder

if ($args.Count -gt 0)
{
    switch -Exact($args[0])
    {
        '-export'
        {
            export_certs
        }

        '-exportOpenKeys'
        {
            export_open_keys
        }

        '-import'
        {
            import_certs_from_reg_files
        }

        '-import_w7'
        {
            $importFolder = $importFolderW7
            import_certs_from_reg_files
        }

        '-clear'
        {
            remove_all_certs
        }

        '-expiring'
        {
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

            #find_expiring2 -interval $interval -trapName $tName
            find_expiring3 -interval $interval -trapName $tName
        }

        '-sort'
        {
            # --- !!!!!!!!!!!!!!!!!!!! Очень осторожно с этим параметром
            # --- Сортируем сертификаты по владельцу в каталог $sortFolder
            
            sort_certs
        }

        '-conf'
        {
            launcher_conf
        }

        '-rar_import'
        {   
            imports_to_rar
        }

        '-orgcsv'
        {
            orgs_to_csv
        }

        '-setspass'
        {
            set_pass_to_certs -pass $pass
        }
    }
}

# --- 1
#export_certs

# --- 2
#export_open_keys

# --- 3
#remove_all_certs

# --- 4
#import_certs_from_reg_files

# --- 5
#certs_sort

#launcher_conf

#find_expiring2 60

#find_expiring2 14 'certs_expiring'

#export_open_keys

#get_cert_parent ''

#imports_to_rar

#orgs_to_csv

#set_pass_to_certs $pass

#auto_detect_items_to_zabbix

#find_expiring3 60
