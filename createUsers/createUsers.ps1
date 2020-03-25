# Скрипт создаёт пользователей в системе, которые задаются файлами с расширением ".ui"
# Имя пользователя - это имя файла
# В теле файла содержится одна строка вида <пароль>=<права>
#
# Для "get-localuser" нужно установить: Framework (WMF) v5.1 - http://www.catalog.update.microsoft.com/Search.aspx?q=3191564 - WMF_v5.1_windows2012_r2.msu

clear

$uiPath = $PSScriptRoot+'\ui'

$uiList = Get-Childitem -Path $uiPath

$localUsers = get-localuser

foreach ($item in $uiList)
{
    $uiFilePath = $uiPath+'\'+$item.Name
    
    $temp = $item.Name.split('.')
    $u    = $temp[0] # Имя пользователя
    
    $temp = (get-content $uiFilePath -totalcount 1)
    
    $temp = $temp.split('=')
    $p    = ConvertTo-SecureString $temp[0] -AsPlainText -Force # Пароль пользователя
    $g    = $temp[1]                                            # Права пользователя

    $fl = $False

    foreach ($item2 in $localUsers)
    {
        if ($u -eq $item2)
        {
            $fl = $True
            break
        }
    }

    if ($fl -eq $False)
    {
        New-LocalUser $u -Password $p -FullName $u -Description ''

        if ($g -eq 'user')
        {
            Add-LocalGroupMember -Group "Пользователи" -Member $u
            Add-LocalGroupMember -Group "Пользователи удаленного рабочего стола" -Member $u
        }
        #
        if ($g -eq 'admin')
        {
            Add-LocalGroupMember -Group "Администраторы" -Member $u
            Add-LocalGroupMember -Group "Пользователи удаленного рабочего стола" -Member $u
        }
    }

    $x = Remove-Item $uiFilePath -Force
} 

# ----------------------------------------------------------------------------------------------------------------------------------------------

$ul = Get-LocalUser

foreach ($item in $ul)
{
    if ($item.Enabled -eq $True)
    {
        $item.Name
        Set-LocalUser -Name $item.Name –PasswordNeverExpires $True
        '----------------------'
    }
}