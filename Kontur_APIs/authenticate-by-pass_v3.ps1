# Мануал https://kontur-ofd-api.readthedocs.io/ru/latest/index.html

# Площадки (https://kontur-ofd-api.readthedocs.io/ru/latest/Endpoints.html)
# тестовая https://ofd-project.kontur.ru:11002/
# боевая   https://ofd-api.kontur.ru/

Clear

class TOrganization
{
    [string]$id             # id организации в Контур
    [string]$inn            # ИНН организации
    [string]$kpp            # КПП организации
    [string]$shortName      # Краткое наименование организации
    [string]$fullName       # Полное наименование организации
    [string]$isBlocked      # Заблокирована ли организация или нет
    [string]$creationDate   # 
    [string]$hasEvotorOffer #

    [void]print_items()
    {
        Write-Host 'id:             ' $this.id
        Write-Host 'inn:            ' $this.inn
        Write-Host 'kpp:            ' $this.kpp
        Write-Host 'shortName:      ' $this.shortName
        Write-Host 'fullName:       ' $this.fullName
        Write-Host 'isBlocked:      ' $this.isBlocked
        Write-Host 'creationDate:   ' $this.creationDate
        Write-Host 'hasEvotorOffer: ' $this.hasEvotorOffer
    }
}

class TOrganizations
{
    $Items

    TOrganizations()
    {
        $this.Items = New-Object System.Collections.ArrayList
    }

    [void]print_items()
    {
        foreach ($item in $this.Items)
        {
            $item.print_items()
        }
    }
}

class TKkt
{
    [string]$regNumber      #
    [string]$name           #
    [string]$serialNumber   #
    [string]$organizationId #
    [string]$modelName      #
    $salesPointPeriods      #
    $fnEntity               #
    $closeDate              #
}

class TKkts
{
    $Items

    TKkts()
    {
        $this.Items = New-Object System.Collections.ArrayList
    }

    [void]print_items()
    {
        foreach ($item in $this.Items)
        {
            $item.print_items()
        }
    }
}

class TKontur
{
    [string]$ofd_api_key  # ключ интегратора
    [string]$SID          #
    [string]$cookieDomain #
    [string]$endPoint     # площадка
    [string]$mail         # емаил авторизации
    [string]$pass         # пароль авторизации
    $WebSession

    [string]$authStatusCode # код статуса авторизации

    $OrganizationsList #
    $OrganizationsCount
    ###########################################

    TKontur()
    {
        $this.WebSession        = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $this.SID               = 'null'
        $this.OrganizationsList = New-Object TOrganizations
        $this.OrganizationsCount = 0
    }

    [string]get_SID([string]$str)
    {
        [string]$result = ''
        $pos = $str.IndexOf(':"')
        if ($pos -ne -1)
        {
            $result = $str.Substring($pos+2, $str.Length -$pos -4)
        }
        return $result
    }

    [void]authenticate_by_pass()  
    {
        # Аутентификация/авторизация и получение SID-a (https://kontur-ofd-api.readthedocs.io/ru/latest/Auth/authenticate-by-pass.html)
        $URL     = 'https://api.kontur.ru/auth/authenticate-by-pass?login='+$this.mail # урл авторизации
        $Request = Invoke-WebRequest -Uri $URL -Body $this.pass -Method 'POST'
        
        $this.authStatusCode = [string]$Request.StatusCode
        if ($this.authStatusCode -eq '200')
        {            
            $this.SID = $this.get_SID($Request.Content)

            # кука авторизации
            $authCookie        = New-Object System.Net.Cookie 
            $authCookie.Name   = 'ofd_api_key'
            $authCookie.Value  = $this.ofd_api_key 
            $authCookie.Domain = $this.cookieDomain
            $this.WebSession.Cookies.Add($authCookie)

            # кука с SID-ом
            $SIDCookie        = New-Object System.Net.Cookie 
            $SIDCookie.Name   = 'auth.sid'
            $SIDCookie.Value  = $this.SID
            $SIDCookie.Domain = $this.cookieDomain
            $this.WebSession.Cookies.Add($SIDCookie)
        }
    }

    [void]get_organizations()
    {   
        # метод organizations (https://kontur-ofd-api.readthedocs.io/ru/latest/http/organizations.html) 
        if ($this.authStatusCode -eq '200')
        {
            #Write-Host '123'
            
            $URL     = $this.endPoint + 'v1/organizations'
            $Request = Invoke-WebRequest -Uri $URL -WebSession $this.WebSession -Method 'GET'

            if ($Request.StatusCode -eq '200')
            {
                $Organization = New-Object TOrganization
                
                $json = $Request.Content | ConvertFrom-Json

                foreach ($item in $json) 
                {
                    $Organization.id             = $item.id
                    $Organization.inn            = $item.inn
                    $Organization.kpp            = $item.kpp
                    $Organization.shortName      = $item.shortName
                    $Organization.fullName       = $item.fullName
                    $Organization.isBlocked      = $item.isBlocked
                    $Organization.creationDate   = $item.creationDate
                    $Organization.hasEvotorOffer = $item.hasEvotorOffer

                    $this.OrganizationsList.Items.Add($Organization)
                }

                $this.OrganizationsCount = $this.OrganizationsList.Items.Count
            }
        }
    }

    [void]get_cashboxes()
    {
        # метод cashboxes (https://kontur-ofd-api.readthedocs.io/ru/latest/http/cashboxes.html)
        if ($this.OrganizationsCount -gt 0)
        {
            foreach ($item in $this.OrganizationsList.Items) 
            {
                Write-Host $item.shortName

                $URL     = $this.endPoint + 'v1/organizations/' + $item.id + '/cashboxes'
                $Request = Invoke-WebRequest -Uri $URL -WebSession $this.WebSession -Method 'GET'

                #Write-Host $Request.Content

                if ($Request.StatusCode -eq '200')
                {
                    $json = $Request.Content  | ConvertFrom-Json
                    foreach ($item in $json)
                    {
                         Write-Host $item.regNumber
                    }
                }
                
            }
        }
    }
}

$Kontur = New-Object TKontur

$Kontur.ofd_api_key  = '' # ключ для доступа личного кабинета через браузер
$Kontur.cookieDomain = '.kontur.ru'
$Kontur.endPoint     = 'https://ofd-api.kontur.ru/'
$Kontur.mail         = ''
$Kontur.pass         = ''

$Kontur.authenticate_by_pass()
#$Kontur.SID
$Kontur.get_organizations()
#$Kontur.OrganizationsList.print_items()
$Kontur.get_cashboxes()
