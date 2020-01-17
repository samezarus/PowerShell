# sameza
# 
# Удаляем все сертификаты (закрытые части) Крипто-Про у текущего пользователя
#

clear

$exportPath = $PSScriptRoot+'\sbis_keys_global'

# Получение имени текущей учетной записи
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Получение SID текущего пользователя
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
