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
#################################################################################################################################
#################################################################################################################################

# Получаем список сертификатов текущего пользователя
$certs = get_current_user_certs

# Экспортируем все сертификаты текущего пользователя
#export_certs_from_current_user -certs $certs -certsFolder $certsFolder

# Импортируе все сертификаты текущему пользователю
#import_certs_to_current_user -certs $certs -certsFolder $certsFolder

# Удаляем все сертификаты
#remove_all_certs -certs $certs
