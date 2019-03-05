Clear-Host
# ------------------------------------------------------------------------------------------------

#dir "C:\windows\system32\mstsc.exe" -rec | select fullname,@{n="Version";e={$_.versioninfo.fileversion}}
#dir "C:\windows\system32\mstsc.exe" -rec | %{$_.versioninfo.fileversion}

#$filell = dir "C:\windows\system32\mstsc.exe"
#$filell.versioninfo.fileversion

Clear-Host
# ------------------------------------------------------------------------------------------------

'START'

$pc_count = 0

$find_file = '\C$\Windows\system32\mstsc.exe'

$pc_list = Get-ADComputer -Filter * -SearchBase 'OU=vl82 Vladivostok Yumasheva 7b, OU=Shop, OU=Ws,DC=severotorg,DC=local'

foreach ($item in $pc_list)
{
    $pc_count +=1
    $pc_count.ToString()
    
    $pc_status = Test-Connection -computername $item.Name -quiet

    $fle = '\\'+$item.Name+$find_file

    if ($pc_status -eq $True) 
    {
        $item.Name+' в сети'
        #
        if (Test-Path  ($fle))
        {
            '  файл "'+$fle+'" найден' 
            $filelll = dir $fle
            '    '+$filelll.versioninfo.fileversion

        }
        else
        {
            '  файл "'+$fle+'" не найден' 
        }
    }
    else
    {
        $item.Name+' не в сети'
    }

    '------------------------'
}
    
'END'