# for TrippLite ups

# srvice 'C:\Program Files (x86)\TrippLite\PowerAlert\engine\panms.exe' (PowerAlert Agent) 
# create log 'C:\Program Files (x86)\TrippLite\PowerAlert\data\padlog1.dat'

Clear-Host

# zabbix conf file
# 	Server=zabbix
# 	ServerActive=zabbix
# 	EnableRemoteCommands=1
# 	UserParameter=tripplite[*],PowerShell.exe -nologo C:\scripts\tripplite.ps1 $1

# zabbix conf
# key: tripplite[<param>] exmpl: tripplite[device_mode]

$terms = @()

#$terms+='date'
#$terms+='time'
$terms+='location'
$terms+='region'
$terms+='device_name'
$terms+='device_id'
$terms+='date_installed'
$terms+='serial_number'
$terms+='low_battery_warning'
$terms+='battery_age'
$terms+='load_banks_controllable'
$terms+='load_baks_total'
$terms+='device_mode' # 'Utility'/'On Battery'
$terms+='input_voltage'
$terms+='output_voltage'
$terms+='battery_voltage'
$terms+='battery_charge_remaining'
$terms+='battery_minutes_remaining'
$terms+='input_frequency'
$terms+='output_frequency'
$terms+='output_current'
$terms+='output_power'
$terms+='output_load'
$terms+='output_source'
$terms+='tap_state'
$terms+='general_fault_alarm'
$terms+='battery_charge'
$terms+='watchdog_status'
$terms+='watchdog_time'
$terms+='audible_alarm_status'
$terms+='battery_condition'
$terms+='self_test_date'
$terms+='self_test_status'
$terms+='va_rating'
$terms+='nominal_battery_vltage'
$terms+='low_transfer_voltage'
$terms+='high_transfer_voltage'
$terms+='device_type'
$terms+='firmware_version'
$terms+='auto_restart_on_shutdown'
$terms+='auto_restart_on_delayed_wakeup'
$terms+='auto_restart_on_low_voltage'
$terms+='auto_restart_on_overload'
$terms+='auto_restart_on_overtemp'
$terms+='14_day_self_test'
$terms+='minimum_input_voltage'
$terms+='maximum_input_voltage'
$terms+='power_on_delay'
$terms+='battery_age_alarm_threshhold'
$terms+='communication_protocol'
$terms+='communication_port'
$terms+='operating_system'

#$fname = 'C:\Users\sameza\Desktop\TrippLite\padlog1.dat'
$fname = 'C:\Program Files (x86)\TrippLite\PowerAlert\data\padlog1.dat'

$sPos = 0
$ePos = 0

foreach($line in Get-Content $fname) 
{
    $sPos = $line.LastIndexOf(', , , ,')
    $ePos = $line.LastIndexOf('),')
}

if (($sPos -gt 0) -and  ($ePos -gt 0))
{
    $params = @{} # ассоциативный массив ключ=значение

    $str = $line.Substring($sPos, $ePos -$sPos +1)

    $params['date'] = $line.Substring($sPos -24, 8)
    $params['time'] = $line.Substring($sPos -12, 6)

    $str+= ','

    $sPos = 0
    $pIndex = 0
    

    For ($i=0; $i -le $str.Length; $i++)
    {
        if ($str[$i] -eq ',')
        {
            $param = $str.Substring($sPos, $i -$sPos)
            $param = $param.TrimEnd('VA')
            $param = $param.TrimEnd('V')
            $param = $param.TrimEnd('A')
            $param = $param.TrimEnd('%')
            $param = $param.TrimEnd('W')
            $param = $param.TrimEnd('Years')
            $param = $param.TrimEnd('Hz')
            #$param = $param.TrimEnd('Second') # ??? TrimEnd

            $params[$terms[$pIndex]] = $param.Trim(' ')
            #
            $sPos = $i +1
            $pIndex++
        }
    }

    
    #$params
    #$testArg0 = 'input_voltage'
    #$params[$testArg0]

    $params[$args[0]]
}



