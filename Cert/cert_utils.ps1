# sameza
#
# В оснастке Крипто ПРО удалить все считыватели кроме "Реестр" что бы максимально сократить количество окон выскакивающих при добавлении
#
#

Clear

$certsLocation = 'cert:\CurrentUser\My'
#$certsFolder   = $PSScriptRoot
$certsFolder   = 'c:\temp'
$mypwd         = ConvertTo-SecureString -String 'qwerty' -Force -AsPlainText

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
            # --- Экспортируем все сертификаты текущего пользователя
            export_certs_from_current_user -certs $certs -certsFolder $certsFolder
        }

        '-import'
        {
            # --- Импортируе все сертификаты текущему пользователю
            import_certs_to_current_user -certs $certs -certsFolder $certsFolder
        }

        '-claer'
        {
            # --- Удаляем все сертификаты
            remove_all_certs -certs $certs
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
