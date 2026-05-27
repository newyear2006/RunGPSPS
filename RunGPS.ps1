# RunGPS-Powershellroutinen

Import-Module PowerHTML -ErrorAction Stop

function New-QueryString {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Query
    )

    ($Query.GetEnumerator() | ForEach-Object {
        '{0}={1}' -f `
            [uri]::EscapeDataString([string]$_.Key), `
            [uri]::EscapeDataString([string]$_.Value)
    }) -join '&'
}

function ConvertFrom-HtmlText {
    param([object]$Node)

    if ($null -eq $Node) {
        return ''
    }

    return ([System.Net.WebUtility]::HtmlDecode($Node.InnerText) -replace [char]160, ' ').Trim()
}

function ConvertTo-RunGpsDecimal {
    param([string]$Text)

    $clean = ([System.Net.WebUtility]::HtmlDecode($Text) -replace [char]160, ' ').Trim()
    $clean = $clean -replace '[^\d,.\-]', ''

    if ($clean -match ',' -and $clean -notmatch '\.') {
        $clean = $clean -replace ',', '.'
    }

    if ([string]::IsNullOrWhiteSpace($clean)) {
        return [decimal]0
    }

    return [decimal]::Parse(
        $clean,
        [System.Globalization.NumberStyles]::Number,
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

function ConvertTo-RunGpsInt {
    param([string]$Text)

    $clean = ([System.Net.WebUtility]::HtmlDecode($Text) -replace [char]160, ' ').Trim()
    $clean = $clean -replace '[^\d\-]', ''

    if ([string]::IsNullOrWhiteSpace($clean)) {
        return 0
    }

    return [int]$clean
}

function Get-RunGpsIdFromCell {
    param([object]$Cell)

    $link = $Cell.SelectSingleNode('.//a[@href]')
    if ($null -eq $link) {
        return $null
    }

    $href = $link.GetAttributeValue('href', '')

    if ($href -match '_(\d+)(?:\D*$|$)') {
        return [int]$matches[1]
    }

    if ($href -match '(?:routeID|trainingID)=(\d+)') {
        return [int]$matches[1]
    }

    return $null
}

function Select-HtmlNodes {
    param(
        [Parameter(Mandatory)]
        [object]$Node,

        [Parameter(Mandatory)]
        [string]$XPath
    )

    if ($null -eq $Node) {
        return @()
    }

    $nodes = $Node.SelectNodes($XPath)

    if ($null -eq $nodes) {
        return @()
    }

    return @($nodes | Where-Object { $null -ne $_ })
}

function ConvertFrom-RunGpsRoutesHtml {
    param([string]$Content)

    $doc = $Content | ConvertFrom-Html

    if ($null -eq $doc) {
        throw "HTML konnte nicht geparst werden. Content ist leer oder ungültig."
    }

    $rows = Select-HtmlNodes -Node $doc -XPath '//tr[td]'
	if ($rows.Count -eq 0) {
	    $debugFile = Join-Path $PWD 'debug-routes.html'
	    $Content | Set-Content -Path $debugFile -Encoding utf8
	
	    Write-Host "Keine <tr><td> Tabellenzeilen gefunden."
	    Write-Host "HTML-Laenge: $($Content.Length)"
	    Write-Host "Debug-Datei: $debugFile"
	
	    throw "Keine Tabellenzeilen in der Routen-HTML-Seite gefunden. Vermutlich ist die geladene Seite nicht die erwartete userRoutes.jsp-Tabelle."
	}

    foreach ($row in $rows) {
        $cells = Select-HtmlNodes -Node $row -XPath './td'

        if ($cells.Count -lt 9) {
            continue
        }

        $id = Get-RunGpsIdFromCell -Cell $cells[2]
        if ($null -eq $id) {
            continue
        }

        $values = @($cells | ForEach-Object { ConvertFrom-HtmlText $_ })

        [PSCustomObject]@{
            Datum    = Get-Date $values[0]
            Sportart = $values[1]
            ID       = $id
            Titel    = $values[2]
            Distanz  = ConvertTo-RunGpsDecimal $values[3]
            Läufe    = ConvertTo-RunGpsInt $values[4]
            Ort      = $values[5]
            Land     = $values[6]
            Aufstieg = ConvertTo-RunGpsInt $values[7]
            Abstieg  = ConvertTo-RunGpsInt $values[8]
        }
    }
}

function ConvertFrom-RunGpsTrainingsHtml {
    param([string]$Content)

    $doc = $Content | ConvertFrom-Html

    if ($null -eq $doc) {
        throw "HTML konnte nicht geparst werden. Content ist leer oder ungültig."
    }

    $rows = Select-HtmlNodes -Node $doc -XPath '//tr[td]'

    if ($rows.Count -eq 0) {
        $debugFile = Join-Path $PWD 'debug-trainings.html'
        $Content | Set-Content -Path $debugFile -Encoding utf8

        Write-Host "Keine <tr><td> Tabellenzeilen gefunden."
        Write-Host "HTML-Laenge: $($Content.Length)"
        Write-Host "Debug-Datei: $debugFile"

        throw "Keine Tabellenzeilen in der Trainings-HTML-Seite gefunden. Vermutlich ist die geladene Seite nicht die erwartete userTrainings.jsp-Tabelle."
    }

    foreach ($row in $rows) {
        $cells = Select-HtmlNodes -Node $row -XPath './td'

        if ($cells.Count -lt 15) {
            continue
        }

        $id = Get-RunGpsIdFromCell -Cell $cells[2]

        if ($null -eq $id) {
            continue
        }

        $values = @($cells | ForEach-Object { ConvertFrom-HtmlText $_ })

        [PSCustomObject]@{
            Datum            = Get-Date $values[0]
            Sportart         = $values[1]
            ID               = $id
            Titel            = $values[2]
            Distanz          = ConvertTo-RunGpsDecimal $values[3]
            Dauer            = [TimeSpan]$values[4]
            Kalorien         = ConvertTo-RunGpsInt $values[5]
            HerzfrequenzD    = ConvertTo-RunGpsDecimal $values[6]
            TrittfrequenzD   = ConvertTo-RunGpsDecimal $values[7]
            GeschwindigkeitD = ConvertTo-RunGpsDecimal $values[8]
            GeschwindigkeitDA = ConvertTo-RunGpsDecimal $values[9]
            HoeheMin         = ConvertTo-RunGpsInt $values[10]
            HoeheMax         = ConvertTo-RunGpsInt $values[11]
            Abstieg          = ConvertTo-RunGpsInt $values[12]
            Aufstieg         = ConvertTo-RunGpsInt $values[13]
            Gewicht          = ConvertTo-RunGpsInt $values[14]
            DistanzBereich   = Get-DistanzBereich -km (ConvertTo-RunGpsDecimal $values[3])
        }
    }
}

# https://d-fens.ch/2013/04/29/invoke-webrequest-does-not-save-all-cookies-in-sessionvariable/
function ConvertFrom-SetCookieHeader {
  [cmdletbinding()]
  Param(
    [Parameter(Mandatory=$true,Position=0)]
    [string] $SetCookieHeader
  )
  
  $res = @{}
  $res.SetCookieSplit = "([^=]+)=([^;]+);\ ";
  $SetCookie = $SetCookieHeader; #$r.Headers.'set-cookie';
  $SetCookie = $SetCookie.Replace('Secure,', 'Secure=1;').Replace('Secure;', 'Secure=1;');
  $SetCookie = $SetCookie.Replace('HTTPOnly,', 'HTTPOnly=1;').Replace('HTTPOnly;', 'HTTPOnly=1;');
  $MatchInfo = Select-String -InputObject $SetCookie $res.SetCookieSplit -AllMatches;
  $CookieHeader = "";
  foreach($Match in $MatchInfo.Matches) {
    if(!$Match.Success) { continue; }
      $CookieName = $Match.Groups.Value[1].Trim();
      $CookieValue = $Match.Groups.Value[2].Trim();
      #$CookiePath = $res.CookiePathDefault;
      $CookieDomain = $WebHost;
      switch($CookieName) {
        "version" { break; }
        "path" { break; }
        "domain" { break; }
        "expires" { break; }
        "HTTPOnly" { break; }
        "Secure" { break; }
        default {
        Write-Verbose ("'{0}' = '{1}'" -f $CookieName, $CookieValue);
        if(!$CookieHeader) {
          $CookieHeader = ("{0}={1}" -f $CookieName, $CookieValue);
        } else {
          $CookieHeader += ("; {0}={1}" -f $CookieName, $CookieValue);
        } # if
      }
    } # switch
  } # foreach Match
  return $CookieHeader;
} # ConvertFrom-SetCookieHeader

# parst einen übergebenen HTML-String und liest relevante Trainingsdaten aus und gibt diese als Objekt zurück
Function NewRoute {
    Param (
        $htmlRoute
    )
    $r=[PSCustomObject]@{Datum=(Get-Date $htmlRoute.children[0].innerhtml);
                      Sportart=($htmlRoute.children[1].innerText.Trim());
                      ID=[int]$htmlRoute.children[2].childNodes[0].pathname.substring($htmlRoute.children[2].childNodes[0].pathname.LastIndexOf("_")+1);
                      Titel=$htmlRoute.children[2].innerText;
                      Distanz=[decimal]$htmlRoute.children[3].innerText;
                      Laeufe=[int]$htmlRoute.children[4].innerText;
                      Ort=$htmlRoute.children[5].innerText;
                      Land=$htmlRoute.children[6].innerText;
                      Aufstieg=[int]$htmlRoute.children[7].innerText;
                      Abstieg=[int]$htmlRoute.children[8].innerText;
                      }
    $r
}

# $runGPS muss existieren!
Function SaveRoutesData {
    [Cmdletbinding()]
    Param(
        [String]$ID,
        [String]$Path
    )

    Write-Verbose "Saving $ID"
    SaveRouteGPSData -ID $ID -FileType GPX -Path $Path
    SaveRouteGPSData -ID $ID -FileType KML -Path $Path

}

# $runGPS muss existieren!
Function SaveRouteGPSData {
    [CmdletBinding()]
    Param(
        [String]$ID,
        [String]$FileType = "GPX",
        [String]$Path
    )
    
    $SaveFile = Join-Path -Path $Path -ChildPath "Route-$($ID).$FileType"

    Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/routePlanner/dlServices/$($FileType.ToLower()).jsp?routeID=$ID" -OutFile $SaveFile

}

# gibt zu einer KM-Zahl den Distanzbereich zurück
Function Get-DistanzBereich {
  [CmdletBinding()]
  Param ([decimal]$km)

  switch ($km) {
    {$km -ge 42.195} {'42195';break;}
    {$km -ge 21.098} {'21098';break;}
    {$km -ge 10.0} {'10000';break;}
    {$km -ge 5.0} {' 5000';break;}
    {$km -ge 1.0} {' 1000';break;}
    default {'    0'}
  }
}  

# parst einen übergebenen HTML-String und liest relevante Trainingsdaten aus und gibt diese als Objekt zurück
function NewTraining {
    Param (
        $htmlTraining
    )
    $t=[PSCustomObject]@{Datum=(Get-Date $htmlTraining.children[0].innerhtml);
                      Sportart=($htmlTraining.children[1].innerText.Trim());
                      ID=[int]$htmlTraining.children[2].childNodes[0].pathname.substring($htmlTraining.children[2].childNodes[0].pathname.LastIndexOf("_")+1);
                      Titel=$htmlTraining.children[2].innerText;
                      Distanz=$htmlTraining.children[3].innerText;
                      Dauer=[TimeSpan]$htmlTraining.children[4].innerText;
                      Kalorien=[int]$htmlTraining.children[5].innerText;
                      HerzfrequenzD=[decimal]$htmlTraining.children[6].innerText;
                      TrittfrequenzD=[decimal]$htmlTraining.children[7].innerText;
                      GeschwindigkeitD=[decimal]$htmlTraining.children[8].innerText;
                      GeschwindigkeitDA=[decimal]$htmlTraining.children[9].innerText;
                      HoeheMin=[int]$htmlTraining.children[10].innerText;
                      HoeheMax=[int]$htmlTraining.children[11].innerText;
                      Abstieg=[int]$htmlTraining.children[12].innerText;
                      Aufstieg=[int]$htmlTraining.children[13].innerText;
                      Gewicht=[int]$htmlTraining.children[14].innerText;
		      DistanzBereich=Get-DistanzBereich -km $htmlTraining.children[3].innerText;
                      }
    
    $t
}

# $runGPS muss existieren!
Function SaveTrainingsData {
    [Cmdletbinding()]
    Param(
        [String]$ID,
        [String]$Path,
	[switch]$Force
    )

    Write-Verbose "Saving $ID"
    SaveTrainingGPSData -ID $ID -FileType GPX -Path $Path -Force:$Force
    SaveTrainingGPSData -ID $ID -FileType KML -Path $Path -Force:$Force
    SaveTrainingGPSData -ID $ID -FileType TCX -Path $Path -Force:$Force

}

# $runGPS muss existieren!
Function SaveTrainingGPSData {
    [CmdletBinding()]
    Param(
        [String]$ID,
        [String]$FileType = "GPX",
        [String]$Path,
	[switch]$Force
    )

    $SaveFile = Join-Path -Path $Path -ChildPath "$($ID).$FileType"

    If (($Force) -or (-Not (Test-Path $SaveFile -Type Leaf))) {
      Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/services/training$($FileType).jsp?trainingID=$ID" -OutFile $SaveFile
    }
}

# ermittelt alle Trainings, Zeitraum und Sportart können optional angegeben werden
# $runGPS muss existieren!
# $user muss existieren!
Function Get-Trainings {
    [CmdletBinding()]
    Param(
        [DateTime]$FromDate=(Get-Date 1.1.2000),
        [DateTime]$ToDate=(Get-Date),
	[String]$Sport=""	# alle
    )

    $Trainings=@()
    $loop = $true
    while ($loop) {
      Write-Verbose "Lade $($FromDate.toString('d')) bis $($ToDate.toString('d'))"
	  $headers = @{
 		   Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
	  }
      $r=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=$($FromDate.toString('yyyy-MM-dd'))&endDate=$($ToDate.toString('yyyy-MM-dd'))&sport=$($Sport)&submitButton=Aktualisieren#" -Headers $headers
      If ($?) {
		$newTrainings = @(ConvertFrom-RunGpsTrainingsHtml -Content $r.Content)
		
		Write-Verbose "Konvertiere $($newTrainings.Count) Einträge"
		
		if ($newTrainings.Count -eq 0) {
		    $debugFile = Join-Path $PWD "debug-trainings-$($FromDate.ToString('yyyyMMdd'))-$($ToDate.ToString('yyyyMMdd')).html"
		    $r.Content | Set-Content -Path $debugFile -Encoding UTF8
		    throw "Keine Trainings gefunden. HTML wurde nach $debugFile geschrieben."
		}
		
		if ($newTrainings.Count -lt 1000) {
		    $loop = $false
		} else {
		    $FromDate = $newTrainings[0].Datum
		}
		
		$Trainings += $newTrainings
	  }
    }
  $Trainings | sort Datum
}

# ermittelt die im aktuellen Monat noch zu leistenden Werte, im Vergleich zum Vorjahresmonat
# $runGPS muss existieren!
Function Get-MonthToGo {
    [CmdletBinding()]
    Param(
	    [DateTime]$Datum=(Get-Date)
    )

    # Monatsanafang ermitteln
    $Anfang = Get-date -Day 1 -Month $Datum.Month -Year $Datum.Year -Hour 0 -Minute 0 -Second 0
    $Ende = $Anfang.AddMonths(1).AddSeconds(-1)
    
    Write-Verbose "Ermittle Daten im Zeitraum $Anfang bis $Ende"
    $t1=Get-Trainings -Sport Hiking -FromDate $Anfang -ToDate $Ende -Verbose
    $Anfang = $Anfang.AddYears(-1)
    $Ende = $Ende.AddYears(-1)
    $t2=Get-Trainings -Sport Hiking -FromDate $Anfang -ToDate $Ende -Verbose
    
    # Werte ausgeben, aktueller Monat
    $t1|select ID, datum, Distanz, Dauer, Distanzbereich|ft
    $t1|select ID, datum, Distanz, Dauer, Distanzbereich|measure -sum -Property Distanz
    
    # Vergleichsmonat aus Vorjahr:
    $t2|select ID, datum, Distanz, Dauer, Distanzbereich|ft
    
    Write-Host "TODO:"
    $t2|where datum -gt (Get-date).AddYears(-1) |select ID, datum, Distanz, Dauer, Distanzbereich|ft
    $t2|where datum -gt (Get-date).AddYears(-1) |select ID, datum, Distanz, Dauer, Distanzbereich|measure -sum -Property Distanz
}

function Get-HtmlAttribute {
    param(
        [Parameter(Mandatory)]
        [string]$Tag,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $pattern = '(?is)\b' + [regex]::Escape($Name) + '\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^>\s]+))'
    $match = [regex]::Match($Tag, $pattern)

    if (-not $match.Success) {
        return $null
    }

    foreach ($groupIndex in 1..3) {
        if ($match.Groups[$groupIndex].Success) {
            return [System.Net.WebUtility]::HtmlDecode($match.Groups[$groupIndex].Value)
        }
    }

    return $null
}

function Get-FirstHtmlForm {
    param(
        [Parameter(Mandatory)]
        [string]$Html
    )

    $match = [regex]::Match($Html, '(?is)<form\b(?<attrs>[^>]*)>(?<inner>.*?)</form>')

    if (-not $match.Success) {
        throw 'Kein HTML-Formular in login.jsp gefunden.'
    }

    [PSCustomObject]@{
        Attributes = $match.Groups['attrs'].Value
        InnerHtml  = $match.Groups['inner'].Value
    }
}

function Get-HtmlFormFields {
    param(
        [Parameter(Mandatory)]
        [string]$FormInnerHtml
    )

    $fields = [ordered]@{}

    $inputMatches = [regex]::Matches($FormInnerHtml, '(?is)<input\b[^>]*>')

    foreach ($inputMatch in $inputMatches) {
        $tag = $inputMatch.Value
        $name = Get-HtmlAttribute -Tag $tag -Name 'name'

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $value = Get-HtmlAttribute -Tag $tag -Name 'value'

        if ($null -eq $value) {
            $value = ''
        }

        $fields[$name] = $value
    }

    return $fields
}

function Get-FirstImageUri {
    param(
        [Parameter(Mandatory)]
        [string]$Html,

        [Parameter(Mandatory)]
        [uri]$BaseUri
    )

    $imgMatch = [regex]::Match($Html, '(?is)<img\b[^>]*>')

    if (-not $imgMatch.Success) {
        return $null
    }

    $tag = $imgMatch.Value

    foreach ($attrName in @('src', 'href')) {
        $raw = Get-HtmlAttribute -Tag $tag -Name $attrName

        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            return ([uri]::new($BaseUri, $raw)).AbsoluteUri
        }
    }

    return $null
}

function ConvertFrom-SecureStringToPlainText {
    param(
        [Parameter(Mandatory)]
        [securestring]$SecureString
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)

    try {
        [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Show-RunGpsCookies {
    param(
        [Parameter(Mandatory)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    foreach ($uriText in @(
        'https://www.rungps.net/',
        'https://www.gps-sport.net/',
        'http://www.gps-sport.net/'
    )) {
        $uri = [uri]$uriText
        $cookies = @($Session.Cookies.GetCookies($uri))

        Write-Host "Cookies for $uriText : $($cookies.Count)"

        foreach ($cookie in $cookies) {
            Write-Host ("  {0}; Domain={1}; Path={2}; Secure={3}" -f `
                $cookie.Name,
                $cookie.Domain,
                $cookie.Path,
                $cookie.Secure)
        }
    }
}

function Connect-RunGPS {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCredential]$Credential
    )

    $headers = @{
        'User-Agent'      = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36'
        'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8'
        'Accept-Language' = 'de-DE,de;q=0.9,en;q=0.8'
        'Cache-Control'   = 'no-cache'
    }

    $loginUri = [uri]'https://www.rungps.net/login.jsp'

    Write-Host "GET $loginUri"

    $loginPage = Invoke-WebRequest `
        -Uri $loginUri `
        -SessionVariable runGPS `
        -Headers $headers `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    $form = Get-FirstHtmlForm -Html $loginPage.Content
    $formAction = Get-HtmlAttribute -Tag $form.Attributes -Name 'action'

    if ([string]::IsNullOrWhiteSpace($formAction)) {
        $formAction = 'login.jsp'
    }

    $postUri = [uri]::new($loginUri, $formAction)

    $body = Get-HtmlFormFields -FormInnerHtml $form.InnerHtml
    $body['userName'] = $Credential.UserName
    $body['password1'] = ConvertFrom-SecureStringToPlainText -SecureString $Credential.Password

    Write-Host "POST $postUri"

    $loginPost = Invoke-WebRequest `
        -Uri $postUri `
        -Method Post `
        -Body $body `
        -ContentType 'application/x-www-form-urlencoded' `
        -WebSession $runGPS `
        -Headers $headers `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    Write-Host "Login response: HTTP $($loginPost.StatusCode), Content length: $($loginPost.Content.Length)"

    $firstImageUri = Get-FirstImageUri `
        -Html $loginPost.Content `
        -BaseUri $loginPost.BaseResponse.RequestMessage.RequestUri

    if ([string]::IsNullOrWhiteSpace($firstImageUri)) {
        throw 'Login erfolgreich? Es wurde kein Bridge-Image in der Login-Antwort gefunden.'
    }

    Write-Host "Bridge image request: $firstImageUri"

    $bridgeResponse = Invoke-WebRequest `
        -Uri $firstImageUri `
        -WebSession $runGPS `
        -Headers $headers `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    Write-Host "Bridge response: HTTP $($bridgeResponse.StatusCode), Content length: $($bridgeResponse.Content.Length)"

    $gpsSportCookieCount = @($runGPS.Cookies.GetCookies([uri]'https://www.gps-sport.net/')).Count

    if ($gpsSportCookieCount -eq 0) {
        Show-RunGpsCookies -Session $runGPS
        throw 'Nach dem Bridge-Image wurde kein Cookie für gps-sport.net erzeugt.'
    }

    Write-Host "Authentication OK. gps-sport.net cookies: $gpsSportCookieCount"

    return [Microsoft.PowerShell.Commands.WebRequestSession]$runGPS
}

Function Disconnect-RunGPS {
  [CmdletBinding()]
  Param(
    [Microsoft.PowerShell.Commands.WebRequestSession]$runGPS
  )

  $r=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/logout.jsp" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

}

# zum Prüfen, was auf der anderen Seite ankommt: https://pipedream.com/requestbin 
# früher: http://requestb.in/

####### Benutzer und Passwort müssen hinterlegt werden!!
$user='Benutzer'
$password='Passwort'
#######
# $cred=Get-Credential -Message 'Bitte Zugangsdaten für RunGPS-Anmeldung eingeben'
# zum Scripten im Workflow dies hinterlegen:
#  env:
#      RUNGPSUSER: ${{ secrets.RUNGPSUSER }}
#      RUNGPSPASSWORD: ${{ secrets.RUNGPSPASSWORD }}
# dann mittels $env:RUNGPSUSER und $env:RUNGPSPASSWORD ansprechen

$user=$env:RUNGPSUSER
"user-Länge: $($user.length)"
$password = ConvertTo-SecureString $env:RUNGPSPASSWORD -AsPlainText -Force
$cred = New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $user, $password
"und?"
$cred
$runGPS = Connect-RunGPS -Credential $cred

Write-Host "------------------- Testaufruf ------------------"
$headers = @{
    'User-Agent'      = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36'
    'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8'
    'Accept-Language' = 'de-DE,de;q=0.9,en;q=0.8'
}

$r2 = Invoke-WebRequest `
    -WebSession $runGPS `
    -Uri "https://www.gps-sport.net/userRoutes.jsp?userName=$([uri]::EscapeDataString($User))&startDate=2006-10-18&endDate=2023-02-06&sport=&searchTerm=&submitButton=Aktualisieren" `
    -Headers $headers `
    -MaximumRedirection 10 `
    -ErrorAction Stop

Write-Host "Routes response: HTTP $($r2.StatusCode), Content length: $($r2.Content.Length)"
Write-Host "------------------- nach Test------------------"



$headers = @{
    Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
}

# Datumsangaben im ISO-Format YYYY-MM-DDD, gibt maximal 1000 Einträge zurück, evtl. muss der Datumsbereich durch zwei
# oder mehrere Aufrufe gesplittet werden!
$routeQuery = [ordered]@{
    userName   = $User
    startDate  = '2006-10-18'
    endDate    = '2023-02-06'
    sport      = ''
    searchTerm = ''
}

$routeUri = 'https://www.gps-sport.net/userRoutes.jsp?' + (New-QueryString -Query $routeQuery)

$r2 = Invoke-WebRequest `
    -WebSession $runGPS `
    -Uri $routeUri `
    -Headers $headers `
    -MaximumRedirection 10 `
    -ErrorAction Stop

Write-Host "Routes response: HTTP $($r2.StatusCode), Content length: $($r2.Content.Length)"
Write-Host "Routes final URI: $($r2.BaseResponse.RequestMessage.RequestUri)"
Write-Host "Routes content preview:"
Write-Host ($r2.Content.Substring(0, [Math]::Min(1000, $r2.Content.Length)))

$r3=Invoke-WebRequest -WebSession $runGPS -Uri "https://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=2006-10-26&endDate=2017-12-31&sport=&submitButton=Aktualisieren#" -Headers $headers

$r4=Invoke-WebRequest -WebSession $runGPS -Uri "https://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=2018-01-01&endDate=2023-02-06&sport=&submitButton=Aktualisieren#" -Headers $headers

### ROUTEN
# Routes
$routes = @(ConvertFrom-RunGpsRoutesHtml -Content $r2.Content)
Write-Host "Routes found: $($routes.Count)"
if ($routes.Count -eq 0) {
    $debugFile = Join-Path $PWD 'debug-routes.html'
    $r2.Content | Set-Content -Path $debugFile -Encoding UTF8
    throw "Keine Routen gefunden. HTML wurde nach $debugFile geschrieben."
}

# erste Route zum Testen speichern
SaveRouteGPSData -ID $routes[0].ID -FileType GPX -Path C:\Temp\RunGPS

# alle ermittelten Routen speichern
$routes | % {SaveRoutesData -ID $_.ID -Path C:\Temp\RunGPS }

### TRAININGS
# Trainings
$trainings=Get-Trainings

# ein Training zum Testen als GPX-Datei speichern 
SaveTrainingGPSData -ID $trainings[0].ID -FileType GPX -Path C:\Temp\RunGPS

# alle Trainings von der Abfrage speichern
$trainings | % {SaveTrainingsData -ID $_.ID -Path C:\Temp\RunGPS }

# Auswertung alle Wanderungen nach Monaten kummuliert ausgeben
$trainings | where Sportart -eq 'Wandern'| select *, @{N='JahrMonat';E={$_.Datum.tostring("yyyy-MM")}}|group -Property JahrMonat|select count, @{N='JahrMonat';E={$_.Name}}, @{N='Km';E={$_.Group|% {$km=0} {$km+=$_.Distanz} {$km}}}

# Auswertung alle Wanderungen nach Jahren kumuliert ausgeben
$trainings | where Sportart -eq 'Wandern'| select *, @{N='Jahr';E={$_.Datum.Year}}|group -Property Jahr|select count, @{N='Jahr';E={$_.Name}}, @{N='Km';E={$_.Group|% {$km=0} {$km+=$_.Distanz} {$km}}}

# Auswertung nach Distanzbereichen
$trainings | where Sportart -eq 'Wandern'| group -Property Distanzbereich|select count, Name, @{N='Km';E={$_.Group|% {$km=0} {$km+=$_.Distanz} {$km}}}

# Ausgabe der gesamten Trainingszeit
New-Timespan -seconds ($trainings|select -ExpandProperty dauer| measure -Sum -Property totalseconds).sum

# Trainings speichern
$trainings|Export-Clixml -Path c:\temp\RunGPS\Trainings.xml

# Trainings wieder laden
$trainings=Import-Clixml -Path .\Trainings.xml
