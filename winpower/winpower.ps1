# for Ippon

Clear-Host

# zabbix conf file
# 	Server=zabbix
# 	ServerActive=zabbix
# 	EnableRemoteCommands=1
# 	UserParameter=winpower[*],PowerShell.exe -nologo C:\scripts\winpower.ps1 $1

# zabbix conf
# key: winpower[<param>] exmpl: winpower[LOAD]



#$args[0]

#$fname = 'C:\Users\sameza\Desktop\winpower\UPSDATA.CSV'
$fname = 'C:\Program Files (x86)\MonitorSoftware\UPSDATA.CSV'

$c  = 0
$sc = @(Get-Content $fname).Length

$val   = ''
$valNo = 0

foreach($line in Get-Content $fname) 
{
    $c++
    if($sc -eq $c)
    {
        #$line
        For ($i=0; $i -le $line.Length; $i++) 
        {
            
            if (($line[$i]-eq ',') -or ($i -eq $line.Length))
            {
                #$val
                $valNo++

                if (($args[0] -eq 'DATE') -and ($valNo -eq 1))
                {
                    $val
                }
                if (($args[0] -eq 'TIME') -and ($valNo -eq 2))
                {
                    $val
                }
                if (($args[0] -eq 'PORT') -and ($valNo -eq 3))
                {
                    $val
                }
                if (($args[0] -eq 'MODEL') -and ($valNo -eq 4))
                {
                    $val
                }
                if (($args[0] -eq 'PROTOCOL') -and ($valNo -eq 5))
                {
                    $val
                }
                if (($args[0] -eq 'IN_V_R') -and ($valNo -eq 6)) # IN-V(R)
                {
                    $val
                }
                if (($args[0] -eq 'IN-V(S)') -and ($valNo -eq 7))
                {
                    $val
                }
                if (($args[0] -eq 'IN-V(T)') -and ($valNo -eq 8))
                {
                    $val
                }
                if (($args[0] -eq 'OUT_V_R') -and ($valNo -eq 9)) # OUT-V(R)
                {
                    $val
                }
                if (($args[0] -eq 'OUT-V(S)') -and ($valNo -eq 10))
                {
                    $val
                }
                if (($args[0] -eq 'OUT-V(T)') -and ($valNo -eq 11))
                {
                    $val
                }
                if (($args[0] -eq 'BATT-V(+)') -and ($valNo -eq 12))
                {
                    $val
                }
                if (($args[0] -eq 'IN-F') -and ($valNo -eq 13))
                {
                    $val
                }
                if (($args[0] -eq 'LOAD') -and ($valNo -eq 14))
                {
                    $val
                }
                if (($args[0] -eq 'TEMP') -and ($valNo -eq 15))
                {
                    $val
                }
                if (($args[0] -eq 'OUT-F') -and ($valNo -eq 16))
                {
                    $val
                }
                if (($args[0] -eq 'BATT-V(-)') -and ($valNo -eq 17))
                {
                    $val
                }
                if (($args[0] -eq 'BPS-V(R)') -and ($valNo -eq 18))
                {
                    $val
                }
                if (($args[0] -eq 'BPS-V(S)') -and ($valNo -eq 19))
                {
                    $val
                }
                if (($args[0] -eq 'BPS-V(T)') -and ($valNo -eq 20))
                {
                    $val
                }
                if (($args[0] -eq 'BPS-F') -and ($valNo -eq 21))
                {
                    $val
                }
                
                $val = ''
            }
            else
            {
                $val += $line[$i]
            }
        }
    }
}

