# https://sergeyvasin.net/2013/04/09/use-powershell-to-find-certificates-that-are-about-to-expire/

clear

function get_value_by_key($key, $parsString)
{
    $result = ''
    
    if (($key -ne '') -and ($parsString -ne ''))
    {
        $index = $parsString.IndexOf($key)
        if ($index -gt -1)
        {
             $str = $parsString.Substring($index +$key.Length, $parsString.Length -$index -$key.Length -1)

             $ndx = $str.IndexOf(',')
             if ($ndx -eq -1)
             {
                $ndx = $str.Length
             } 

             if ($key -eq 'STREET=')
             {
                $str    = $str.Substring(1, $str.Length -1) 
                $ndx    = $str.IndexOf('"')
                $result = $str.Substring(0, $ndx)   
             }
             else
             {        
                 if ($ndx -gt -1)
                 {
                    $result = $str.Substring(0, $ndx)   
                 }
             }
        }
    }

    return $result
}
#
function print_cert_inf($interval)
{
    # $interval - количество дней оставшееся до истечения сертификата (будут отображены между $interval и 1) (0 - параметр не учитывается)
    # Не отобраджает сертификаты с истёкшим периодом
    
    $expiring = ''
    
    if ($interval -eq 0)
    {
       $sertsList = Get-ChildItem -Path cert: -Recurse
    }
    else
    {
        $sertsList = Get-ChildItem -Path cert: -Recurse -ExpiringInDays $interval
        $expiring = ', Истекает через '+$interval + ' дня/дней'
    } 

    foreach ($item in $sertsList)
    {
        #$item | Get-Member -memberType *property

        $inf = $item.Subject | out-string

        if ($inf -ne '')
        {
            $indexG = $inf.IndexOf('G=')
            if ($indexG -gt -1)
            {
                #$item | Get-Member -memberType *property
                #$inf            
                $ogrn       = get_value_by_key 'ОГРН=' $inf   #
                $snils      = get_value_by_key 'СНИЛС=' $inf  #
                $inn        = get_value_by_key 'ИНН=' $inf    #
                $email      = get_value_by_key 'E=' $inf      # Эл. почта
                $org        = get_value_by_key 'CN=' $inf     # Название организации
                $status     = get_value_by_key 'T=' $inf      # Должность
                $familyName = get_value_by_key 'SN=' $inf     # Фамилия
                $name       = get_value_by_key 'G=' $inf      # Имя, Отчество
                $state      = get_value_by_key 'S=' $inf      # Край/Область
                $city       = get_value_by_key 'L=' $inf      # Город
                $street     = get_value_by_key 'STREET=' $inf # Улица        
            
                #$ogrn
                #$snils
                #$inn
                #$email
                #$org
                #$status
                #$familyName
                #$name
                #$state
                #$city
                #$street

                $tmp = $org + ' ИНН: ' +$inn + ', Руководитель: ' + $familyName + ' ' + $name + $expiring
                $tmp | out-file -filepath 'cerLogs.txt' -Append  
                '------------' | out-file -filepath 'cerLogs.txt' -Append 
                $tmp
                '------------'
            }
        }
    }
}
########################################################################################################################################
########################################################################################################################################
########################################################################################################################################

$interval = 30

if ($args.Count -gt 0)
{
    $interval = $args[0]
}

print_cert_inf $interval



