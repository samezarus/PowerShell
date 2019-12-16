# sameza
# 1. Запустить из под учётки где "варятся" и актуализируются все сертификаты
#
#

clear

$exportPath = $PSScriptRoot+'\sbis_keys_global'

# Получение имени текущей учетной записи
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Получение SID текущего пользователя
$objUser = New-Object System.Security.Principal.NTAccount($userName)
$userSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])

#$userName
#$userSID.value

$cryptoProUsers = 'HKLM:\SOFTWARE\Wow6432Node\Crypto Pro\Settings\USERS\'
$userKeysPath = $cryptoProUsers + $userSID + "\Keys"

$userKeys = Get-ChildItem -Path $userKeysPath

foreach ($item in $userKeys)
{
	$keyPath = $item.Name
	
	$temp = $keyPath.split('\')
	
	$org = $temp[8]
	
	$exportFilePath = $exportPath+'\' + $org + '.reg'
	reg.exe export $keyPath $exportFilePath "/y"
}