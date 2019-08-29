# Мануал https://kontur-ofd-api.readthedocs.io/ru/latest/index.html

# Площадки (https://kontur-ofd-api.readthedocs.io/ru/latest/Endpoints.html)
# тестовая https://ofd-project.kontur.ru:11002/
# боевая   https://ofd-api.kontur.ru/

Clear

$ofd_api_key  = 'xxxxxxxxxx' # ключ интегратора
$cookieDomain = '.kontur.ru'

$mail = 'xxxxxxxxxx'
$pass = 'xxxxxxxxxx'

Function getSID ($str)
{
    $result = ''

    $pos = $str.IndexOf(':"')

    if ($pos -ne -1)
    {
        $result = $str.Substring($pos+2, $str.Length -$pos -4)
    }

    return $result
}

########################################################################################################

# --- Аутентификация/авторизация и получение SID-a

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# кука авторизации
$authCookie        = New-Object System.Net.Cookie 
$authCookie.Name   = 'ofd_api_key'
$authCookie.Value  = $ofd_api_key 
$authCookie.Domain = $cookieDomain
$WebSession.Cookies.Add($authCookie)

$URL = 'https://api.kontur.ru/auth/authenticate-by-pass?login='+$mail # урл авторизации

$Request = Invoke-WebRequest -Uri $URL -WebSession $WebSession -Body $pass -Method 'POST'

if ($Request.StatusCode -eq '200')
{
    $sid = getSID $Request.Content
    #'SID: ' + $sid

    # ----------------------------------------------------

    # кука с SID-ом
    $SIDCookie        = New-Object System.Net.Cookie 
    $SIDCookie.Name   = 'auth.sid'
    $SIDCookie.Value  = $sid
    $SIDCookie.Domain = $cookieDomain
    $WebSession.Cookies.Add($SIDCookie)

    $URL = 'https://ofd-project.kontur.ru:11002/v2/organizations'

    $Request = Invoke-WebRequest -Uri $URL -WebSession $WebSession -Method 'GET'

    #$Request | format-list -property *
    $Request.Content
}
