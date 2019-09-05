# Мануал https://kontur-ofd-api.readthedocs.io/ru/latest/index.html

# Площадки (https://kontur-ofd-api.readthedocs.io/ru/latest/Endpoints.html)
# тестовая https://ofd-project.kontur.ru:11002/
# боевая   https://ofd-api.kontur.ru/

Clear

class TSalesPoint
{
    [string]$organizationId # id Организации за которой закреплена данная точка продаж
    [string]$id             # id Точки продажи
    [string]$name           # Имя Точки продажи

    [void]print_items()
    {
        Write-Host 'organizationId: ' $this.organizationId
        Write-Host 'id:             ' $this.id
        Write-Host 'name:           ' $this.name
        Write-Host '--------------------------------------'
    }
}

class TSalesPoints
{
    $Items

    TSalesPoints()
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
    [string]$regNumber      # Регистрационный номер ККТ
    [string]$name           # Имя ККТ
    [string]$serialNumber   # Заводской номер ККТ
    [string]$organizationId # id организации на которую зареган ККТ
    [string]$modelName      # модель ККТ
    #salesPointPeriods              
    [string]$salesPointPeriods_salesPointId # id Точки продаж на которой установленн ККТ
    #fnEntity                       
    [string]$fnEntity_serialNumber       # Серийный номер фискального накопителя(ФН) в ККТ  
    [string]$fnEntity_closeDate_date     # Дата окончания ФН     
    [string]$fnEntity_closeDate_daysLeft # Дней до окончания ФН 
    #
    [string]$salesPointName # Имя точки продаж на которой установленн ККТ
      

    [void]print_items()
    {
        Write-Host 'regNumber:                      ' $this.regNumber
        Write-Host 'name:                           ' $this.name
        Write-Host 'serialNumber:                   ' $this.serialNumber
        Write-Host 'organizationId:                 ' $this.organizationId
        Write-Host 'modelName:                      ' $this.modelName
        Write-Host 'salesPointPeriods_salesPointId: ' $this.salesPointPeriods_salesPointId
        Write-Host 'fnEntity_serialNumber:          ' $this.fnEntity_serialNumber
        Write-Host 'fnEntity_closeDate_date:        ' $this.fnEntity_closeDate_date
        Write-Host 'fnEntity_closeDate_daysLeft:    ' $this.fnEntity_closeDate_daysLeft
        Write-Host 'salesPointName:                 ' $this.salesPointName
        Write-Host '--------------------------------------'
    }
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

    $Kkts        # Список ККТ закреплённый за организацией
    $SalesPoints # Списка точек продаж организации

    TOrganization()
    {
        $this.Kkts        = New-Object TKkts
        $this.SalesPoints = New-Object TSalesPoints
    }
    
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
        Write-Host '-------------------------------------'
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

    $OrganizationsList  #
    $OrganizationsCount #

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

    [string]get_sales_point_name([string]$salesPointId)
    {
        [string]$result = ''

        #Write-Host $organizationId
        #Write-Host $salesPointId

        if ($this.OrganizationsCount -gt 0) 
        {
            foreach ($item in $this.OrganizationsList.items.SalesPoints.items) 
            {
                if ($item.id -eq $salesPointId)
                {
                    $result = $item.name
                    #Write-Host $item.name
                    #Write-Host '-----------'
                }
            }
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

    [void]get_sales_point()
    {
        # метод salespoints (недокументированный)
        # получение списка точек продаж организации

        if ($this.OrganizationsCount -gt 0)
        {
            foreach ($organization in $this.OrganizationsList.Items)
            {
                $URL     = $this.endPoint + 'v1/organizations/' + $organization.id + '/salesPoints'
                $Request = Invoke-WebRequest -Uri $URL -WebSession $this.WebSession -Method 'GET' 

                if ($Request.StatusCode -eq '200')
                {
                    $json = $Request.Content  | ConvertFrom-Json

                    foreach ($item in $json)
                    {
                        $salesPoint                = New-Object TSalesPoint
                        $salesPoint.organizationId = $item.organizationId
                        $salesPoint.id             = $item.id
                        $salesPoint.name           = $item.name

                        #$salesPoint.print_items()

                        $organization.SalesPoints.Items.Add($salesPoint)
                    }
                }
            }
        }
    }

    [void]get_cashboxes()
    {
        # метод cashboxes (https://kontur-ofd-api.readthedocs.io/ru/latest/http/cashboxes.html)
        # получаем списки ККТ по всем организациям доступным пользователю
        if ($this.OrganizationsCount -gt 0)
        {
            foreach ($organization in $this.OrganizationsList.Items) 
            {
                #Write-Host $organization.shortName

                $URL     = $this.endPoint + 'v1/organizations/' + $organization.id + '/cashboxes'
                $Request = Invoke-WebRequest -Uri $URL -WebSession $this.WebSession -Method 'GET'

                #Write-Host $Request.Content

                if ($Request.StatusCode -eq '200')
                {
                    $json = $Request.Content  | ConvertFrom-Json
                    foreach ($item in $json)
                    {
                         #Write-Host $item.regNumber
                         $Kkt = New-Object TKkt

                         $Kkt.regNumber      = $item.regNumber
                         $Kkt.name           = $item.name
                         $Kkt.serialNumber   = $item.serialNumber
                         $Kkt.organizationId = $item.organizationId
                         $Kkt.modelName      = $item.modelName
                         #salesPointPeriods
                            $Kkt.salesPointPeriods_salesPointId = $item.salesPointPeriods.salesPointId
                         #fnEntity
                            $Kkt.fnEntity_serialNumber       = $item.fnEntity.serialNumber
                            $Kkt.fnEntity_closeDate_date     = $item.fnEntity.closeDate.date
                            $Kkt.fnEntity_closeDate_daysLeft = $item.fnEntity.closeDate.daysLeft
                         $Kkt.salesPointName = $this.get_sales_point_name($Kkt.salesPointPeriods_salesPointId)

                         $Kkt.print_items()
                         
                         $organization.Kkts.Items.Add($Kkt)
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

$Kontur.get_sales_point()

$Kontur.get_cashboxes()
