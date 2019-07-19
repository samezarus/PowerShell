Clear

# zabbix conf file
# 	Server=zabbix
# 	ServerActive=zabbix
# 	EnableRemoteCommands=1
# 	UserParameter=IntelRaidConsole[*,*],PowerShell.exe -nologo C:\Script\IntelRaidConsole\IntelRaidConsole.ps1 $1 $2

# zabbix conf
# key: IntelRaidConsole[<param>,<param>] exmpl: IntelRaidConsole[0,physicalDevicesDisks]

# $args[0] - индекс рейд контроллера
# $args[1] - ключ для получения данных ($terms)

function getValue ($str)
{
     $result = ''

     $index = $line.IndexOf(': ')   
     if ($index -gt -1)
     {
        $result = $str.Substring($index+2, $str.Length - $index-3)
     }

     return $result
}

$terms = @()
$params = @{}

$terms += 'virtualDrivesCount'
$terms += 'virtualDrivesDegraded'
$terms += 'virtualDrivesOffline'
$terms += 'physicalDevicesCount'         # Количество слотов под диски
$terms += 'physicalDevicesDisks'         # Количество вставленных дисков
$terms += 'physicalDevicesCriticalDisks'
$terms += 'physicalDevicesFailedDisks'

$raidIndex = '0'
if ($args.Count -eq 2)
{
    $raidIndex = $args[0]
}

$fileLog   = $PSScriptRoot+'\raid.log'
$app       = $PSScriptRoot+'\CmdTool2_64.exe' # Поставить галку - Выполнять от имени администратора
$param     = '-AdpAllinfo -a'+$raidIndex+' -AppLogFile ' + $fileLog

if ((Test-Path $fileLog) -eq $True)
{
    Remove-Item -Path $fileLog -Force
}

$out = cmd /c $app $param

#Start-Sleep -Seconds 1


if ((Test-Path $fileLog) -eq $True)
{
    $lineIndex = 0
    $findLineIndex = 0
    foreach($line in Get-Content $fileLog) 
    {
        $index = $line.IndexOf('Device Present')
        if ($index -gt -1)
        {
            $findLineIndex = $lineIndex +1
        }

        if ($findLineIndex -gt 0)
        {
            if ($lineIndex -eq $findLineIndex +1)
            {
                #$virtualDrivesCount = getValue $line
                $params[$terms[0]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +2)
            {
                #$virtualDrivesDegraded = getValue $line
                $params[$terms[1]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +3)
            {
                #$virtualDrivesOffline = getValue $line
                $params[$terms[2]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +4)
            {
                #$physicalDevicesCount = getValue $line
                $params[$terms[3]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +5)
            {
                #$physicalDevicesDisks = getValue $line
                $params[$terms[4]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +6)
            {
                #$physicalDevicesCriticalDisks = getValue $line
                $params[$terms[5]] = getValue $line
            }
            if ($lineIndex -eq $findLineIndex +7)
            {
                #$physicalDevicesFailedDisks = getValue $line
                $params[$terms[6]] = getValue $line
                break
            }
        }
        $lineIndex++
    }

    if ($args.Count -eq 2)
    {
        $params[$args[1]]
    }
    else
    {
        $pIndex = 0
        For ($i=0; $i -le $terms.Count -1; $i++)
        {
            $params[$terms[$pIndex]]
            $pIndex++
        }
    } 
}
else
{
    '-1'
}
