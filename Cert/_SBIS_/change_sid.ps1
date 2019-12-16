# sameza
#
# Скрипт копирует в реестр ключи из папки $importPath
# 

clear

#[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("windows-1251")
#[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("windows-1251")
#[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("cp866")

$importPath = $PSScriptRoot+'\sbis_keys'
#$importPath = 'C:\Users\sert_user\Desktop\csp'+'\sbis_keys'

# Получение имени текущей учетной записи
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Получение SID текущего пользователя
$objUser = New-Object System.Security.Principal.NTAccount($userName)
$userSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])

#$userName
#$userSID.value

$cryptoProUsers = 'HKLM:\SOFTWARE\Wow6432Node\Crypto Pro\Settings\USERS\'
$userPath       = $cryptoProUsers + $userSID

# --- Изменяем SID-ы в импортируемых сертификатах? 
$keysFileList = Get-Childitem -Path $importPath

foreach ($item in $keysFileList)
{
    $keyFilePath = $item.FullName
    $keyFile     = Get-Content $keyFilePath
    $findTerm    = '[HKEY_LOCAL_MACHINE'

    foreach ($line in $keyFile)
    {
        $findIndex = $line.IndexOf($findTerm)
        if ($findIndex -gt -1)
        {
            $oldSID = $line.split('\')
            $oldSID = $oldSID[6]
            break        
        }
    }
    #
    if ($oldSID -ne '')
    {
        #Get-Content $keyFilePath | ForEach-Object {$_ -replace $oldSID, $userSID.value} | Set-Content $keyFilePath
        $keyFile | ForEach-Object {$_ -replace $oldSID, $userSID.value} | Set-Content $keyFilePath
        #$oldSID
        #$userSID.value
    }
}