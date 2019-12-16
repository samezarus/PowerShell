# sameza
#
# Скрипт выгружает каждый ключ KONTUR в отдедьный файл.
# Запустить из под учётки где "варятся" и актуализируются все сертификаты

clear

$exportPath = $PSScriptRoot+'\kontur_keys'
#$exportPath = 'C:\Users\sert_user\Desktop\csp'+'\kontur_keys'
$exportPath

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
	
	$temp = $temp[8]
	
	$temp = $temp.split('@')
	
	$keyN = $temp[0]
	
	$temp = $temp[1]
	
	$temp = $temp.split('-')
	
	$createDate = $temp[2] + '-' + $temp[1] + '-' + $temp[0]
	$org        = $temp[3]
	
#	$keyPath
#	$keyN
#	$org 
#	$createDate
#	'-----------'
	
	$exportFilePath = $exportPath+'\' + $org + '_' + $createDate + '_' + $keyN + '.reg'
	reg.exe export $keyPath $exportFilePath "/y"
}


