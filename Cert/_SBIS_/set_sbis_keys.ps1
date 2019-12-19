# sameza

clear

# ���������
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("windows-1251")

$importPath = $PSScriptRoot+'\sbis_keys'

# ��������� ����� ������� ������� ������
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# ��������� SID �������� ������������
$objUser = New-Object System.Security.Principal.NTAccount($userName)
$userSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier])

#$userName
#$userSID.value

$cryptoProUsers = 'HKLM:\SOFTWARE\Wow6432Node\Crypto Pro\Settings\USERS\'
$userPath       = $cryptoProUsers + $userSID

# --- �������� SID-� � ������������� ������������? 
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
            $temp   = $line.split('\')
            $oldSID = $temp[6]
            
            $orgStr = $temp[8]
            $orgStr = $orgStr.split(']')
            $orgStr = $orgStr[0]

            break        
        }
    }
    #
    if (($oldSID -ne '') -and ($oldSID -ne $userSID.value))
    {
        # ������ ��� �� ��� �������� ������������ 
        $keyFile | ForEach-Object {$_ -replace $oldSID, $userSID.value} | Set-Content $keyFilePath

        # ������������� ���������� ������� ����� ����� � ������
        reg import $keyFilePath
        
        # ������������� ���������� �������� ����� ����� � ������
        $cont = '\\.\REGISTRY\' + $orgStr
        &'C:\Program Files\Crypto Pro\CSP\csptest.exe' -property -cinstall -cont $cont
    }
}