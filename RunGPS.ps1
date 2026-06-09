# RunGPS-Powershellroutinen

$culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
[System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
[System.Globalization.CultureInfo]::DefaultThreadCurrentCulture = $culture
[System.Globalization.CultureInfo]::DefaultThreadCurrentUICulture = $culture

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
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$ID,

        [Parameter(Mandatory)]
        [String]$Path,

        [switch]$Force
    )

    SaveRouteGPSData -ID $ID -FileType GPX -Path $Path -Force:$Force
    # SaveRouteGPSData -ID $ID -FileType KML -Path $Path -Force:$Force
}

# $runGPS muss existieren!
Function SaveRouteGPSData {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$ID,

        [String]$FileType = "GPX",

        [Parameter(Mandatory)]
        [String]$Path,

        [switch]$Force
    )

    New-Item -ItemType Directory -Force -Path $Path | Out-Null

    $SaveFile = Join-Path -Path $Path -ChildPath "Route-$($ID).$FileType"

    if ((-not $Force) -and (Test-Path $SaveFile -PathType Leaf)) {
        Write-Host "Skip existing route file: $SaveFile"
        return
    }

    Write-Host "Download route $ID as $FileType -> $SaveFile"

    Invoke-WebRequest `
        -WebSession $runGPS `
        -Uri "https://www.gps-sport.net/routePlanner/dlServices/$($FileType.ToLower()).jsp?routeID=$ID" `
        -OutFile $SaveFile `
        -MaximumRedirection 10 `
        -ErrorAction Stop
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
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$ID,

        [Parameter(Mandatory)]
        [String]$Path,

        [switch]$Force
    )

    # SaveTrainingGPSData -ID $ID -FileType GPX -Path $Path -Force:$Force
    # SaveTrainingGPSData -ID $ID -FileType KML -Path $Path -Force:$Force
    SaveTrainingGPSData -ID $ID -FileType TCX -Path $Path -Force:$Force
}

# $runGPS muss existieren!
Function SaveTrainingGPSData {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [String]$ID,

        [String]$FileType = "GPX",

        [Parameter(Mandatory)]
        [String]$Path,

        [switch]$Force
    )

    New-Item -ItemType Directory -Force -Path $Path | Out-Null

    $SaveFile = Join-Path -Path $Path -ChildPath "$($ID).$FileType"

    if ((-not $Force) -and (Test-Path $SaveFile -PathType Leaf)) {
        Write-Host "Skip existing training file: $SaveFile"
        return
    }

    Write-Host "Download training $ID as $FileType -> $SaveFile"

    Invoke-WebRequest `
        -WebSession $runGPS `
        -Uri "https://www.gps-sport.net/services/training$($FileType).jsp?trainingID=$ID" `
        -OutFile $SaveFile `
        -MaximumRedirection 10 `
        -ErrorAction Stop
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
  $Trainings | Sort-Object -Property Datum
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

function Get-RunGpsBridgeImageUri {
    param(
        [Parameter(Mandatory)]
        [string]$Html,

        [Parameter(Mandatory)]
        [uri]$BaseUri
    )

    $imgMatches = [regex]::Matches($Html, '(?is)<img\b[^>]*>')

    foreach ($imgMatch in $imgMatches) {
        $tag = $imgMatch.Value

        foreach ($attrName in @('src', 'href')) {
            $raw = Get-HtmlAttribute -Tag $tag -Name $attrName

            if ([string]::IsNullOrWhiteSpace($raw)) {
                continue
            }

            $absolute = ([uri]::new($BaseUri, $raw)).AbsoluteUri

            if ($absolute -match '(?i)^https?://www\.gps-sport\.net/whitePixel\.jsp\?') {
                return $absolute
            }
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

function Convert-RunGpsTcxToGeoJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TcxPath,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [switch]$Force
    )

    if ((-not $Force) -and (Test-Path $OutFile -PathType Leaf)) {
        Write-Host "Skip existing GeoJSON: $OutFile"
        return
    }

    [xml]$xml = Get-Content -Path $TcxPath -Raw

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('tcx', 'http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2')

    $trackpoints = $xml.SelectNodes(
        '//tcx:Trackpoint[tcx:Position/tcx:LatitudeDegrees and tcx:Position/tcx:LongitudeDegrees]',
        $ns
    )

    if ($null -eq $trackpoints -or $trackpoints.Count -eq 0) {
        Write-Warning "Keine Koordinaten gefunden: $TcxPath"
        return
    }

    $coordinates = [System.Collections.Generic.List[object]]::new()

    foreach ($tp in $trackpoints) {
        $latText = $tp.Position.LatitudeDegrees
        $lonText = $tp.Position.LongitudeDegrees
        $altText = $tp.AltitudeMeters

        if ([string]::IsNullOrWhiteSpace($latText) -or [string]::IsNullOrWhiteSpace($lonText)) {
            continue
        }

        $lat = [double]::Parse($latText, [System.Globalization.CultureInfo]::InvariantCulture)
        $lon = [double]::Parse($lonText, [System.Globalization.CultureInfo]::InvariantCulture)

        if (-not [string]::IsNullOrWhiteSpace($altText)) {
            $alt = [double]::Parse($altText, [System.Globalization.CultureInfo]::InvariantCulture)
            $coordinates.Add(@($lon, $lat, $alt))
        }
        else {
            $coordinates.Add(@($lon, $lat))
        }
    }

    if ($coordinates.Count -lt 2) {
        Write-Warning "Zu wenige Koordinaten für LineString: $TcxPath"
        return
    }

    $id = [IO.Path]::GetFileNameWithoutExtension($TcxPath)

    $geoJson = [ordered]@{
        type     = 'FeatureCollection'
        features = @(
            [ordered]@{
                type       = 'Feature'
                properties = [ordered]@{
                    id   = $id
                    file = [IO.Path]::GetFileName($TcxPath)
                }
                geometry   = [ordered]@{
                    type        = 'LineString'
                    coordinates = $coordinates
                }
            }
        )
    }

    $outDir = Split-Path -Path $OutFile -Parent
    if (-not [string]::IsNullOrWhiteSpace($outDir)) {
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    }

    $geoJson |
        ConvertTo-Json -Depth 50 |
        Set-Content -Path $OutFile -Encoding utf8

    Write-Host "GeoJSON geschrieben: $OutFile; Punkte: $($coordinates.Count)"
}

function Convert-AllRunGpsTcxToGeoJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TcxDirectory,

        [Parameter(Mandatory)]
        [string]$GeoJsonDirectory,

        [switch]$Force
    )

    New-Item -ItemType Directory -Force -Path $GeoJsonDirectory | Out-Null

    Get-ChildItem -Path $TcxDirectory -Filter '*.tcx' -File |
        ForEach-Object {
            $outFile = Join-Path $GeoJsonDirectory ($_.BaseName + '.geojson')

            if ((-not $Force) -and (Test-Path $outFile -PathType Leaf)) {
                Write-Host "Skip existing GeoJSON: $outFile"
                return
            }

            Convert-RunGpsTcxToGeoJson -TcxPath $_.FullName -OutFile $outFile
        }
}

function Convert-RunGpsTcxToMapPoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TcxPath
    )

    [xml]$xml = Get-Content -Path $TcxPath -Raw

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('tcx', 'http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2')

    $trackpoint = $xml.SelectSingleNode('//tcx:Trackpoint[tcx:Position/tcx:LatitudeDegrees and tcx:Position/tcx:LongitudeDegrees]', $ns)

    if ($null -eq $trackpoint) {
        return $null
    }

    $lat = [double]::Parse(
        $trackpoint.Position.LatitudeDegrees,
        [System.Globalization.CultureInfo]::InvariantCulture
    )

    $lon = [double]::Parse(
        $trackpoint.Position.LongitudeDegrees,
        [System.Globalization.CultureInfo]::InvariantCulture
    )

    $id = [System.IO.Path]::GetFileNameWithoutExtension($TcxPath)

    [PSCustomObject]@{
        type       = 'Feature'
        geometry   = @{
            type        = 'Point'
            coordinates = @($lon, $lat)
        }
        properties = @{
            id   = $id
            file = (Split-Path $TcxPath -Leaf)
        }
    }
}

function New-RunGpsTrainingsGeoJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TrainingsTcxPath,

        [Parameter(Mandatory)]
        [string]$OutFile
    )

    $features = Get-ChildItem -Path $TrainingsTcxPath -Filter '*.TCX' -File |
        ForEach-Object {
            Convert-RunGpsTcxToMapPoint -TcxPath $_.FullName
        } |
        Where-Object { $null -ne $_ }

    $geoJson = [ordered]@{
        type     = 'FeatureCollection'
        features = @($features)
    }

	$outDir = Split-Path -Path $OutFile -Parent
	
	if (-not [string]::IsNullOrWhiteSpace($outDir)) {
	    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
	}

    $geoJson |
        ConvertTo-Json -Depth 20 |
        Set-Content -Path $OutFile -Encoding utf8

    Write-Host "GeoJSON geschrieben: $OutFile"
    Write-Host "Features: $(@($features).Count)"
}

Function Disconnect-RunGPS {
  [CmdletBinding()]
  Param(
    [Microsoft.PowerShell.Commands.WebRequestSession]$runGPS
  )

  $r=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/logout.jsp" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

}

function Get-Sha256Fingerprint {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text,

        [int]$Length = 16
    )

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hashBytes = [System.Security.Cryptography.SHA256]::HashData($bytes)
    $hash = ([BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()

    return $hash.Substring(0, [Math]::Min($Length, $hash.Length))
}

function Connect-RunGpsPwsh7 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserName,

        [Parameter(Mandatory)]
        [string]$Password
    )

    $ErrorActionPreference = 'Stop'

    $headers = @{
        'User-Agent'      = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36'
        'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8'
        'Accept-Language' = 'de-DE,de;q=0.9,en;q=0.8'
    }

    function Get-HtmlAttribute {
        param(
            [Parameter(Mandatory)][string]$Tag,
            [Parameter(Mandatory)][string]$Name
        )

        $pattern = '(?is)\b' + [regex]::Escape($Name) + '\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^>\s]+))'
        $match = [regex]::Match($Tag, $pattern)

        if (-not $match.Success) {
            return $null
        }

        foreach ($i in 1..3) {
            if ($match.Groups[$i].Success) {
                return [System.Net.WebUtility]::HtmlDecode($match.Groups[$i].Value)
            }
        }

        return $null
    }

    function Get-FirstForm {
        param([Parameter(Mandatory)][string]$Html)

        $match = [regex]::Match($Html, '(?is)<form\b(?<attrs>[^>]*)>(?<inner>.*?)</form>')

        if (-not $match.Success) {
            throw 'Login-Formular nicht gefunden.'
        }

        [PSCustomObject]@{
            Attributes = $match.Groups['attrs'].Value
            InnerHtml  = $match.Groups['inner'].Value
        }
    }

    function Get-FormFields {
        param([Parameter(Mandatory)][string]$FormInnerHtml)

        $fields = [ordered]@{}
        $inputs = [regex]::Matches($FormInnerHtml, '(?is)<input\b[^>]*>')

        foreach ($input in $inputs) {
            $tag = $input.Value
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

    function New-FormUrlEncodedBody {
        param([Parameter(Mandatory)][hashtable]$Fields)

        ($Fields.GetEnumerator() | ForEach-Object {
            '{0}={1}' -f `
                [uri]::EscapeDataString([string]$_.Key), `
                [uri]::EscapeDataString([string]$_.Value)
        }) -join '&'
    }

    function Get-WhitePixelUri {
        param(
            [Parameter(Mandatory)][string]$Html,
            [Parameter(Mandatory)][uri]$BaseUri
        )

        $images = [regex]::Matches($Html, '(?is)<img\b[^>]*>')

        foreach ($image in $images) {
            $tag = $image.Value

            foreach ($attr in @('src', 'href')) {
                $raw = Get-HtmlAttribute -Tag $tag -Name $attr

                if ([string]::IsNullOrWhiteSpace($raw)) {
                    continue
                }

                $absolute = ([uri]::new($BaseUri, $raw)).AbsoluteUri

                if ($absolute -match '(?i)^https?://www\.gps-sport\.net/whitePixel\.jsp\?') {
                    return $absolute
                }
            }
        }

        return $null
    }

    $loginUri = [uri]'https://www.rungps.net/login.jsp'

    Write-Host "GET $loginUri"

    $loginGet = Invoke-WebRequest `
        -Uri $loginUri `
        -SessionVariable runGpsSession `
        -Headers $headers `
        -MaximumRedirection 10

    $form = Get-FirstForm -Html $loginGet.Content

    $formAction = Get-HtmlAttribute -Tag $form.Attributes -Name 'action'

    if ([string]::IsNullOrWhiteSpace($formAction)) {
        $formAction = 'login.jsp'
    }

    $postUri = [uri]::new($loginUri, $formAction)

    $fields = Get-FormFields -FormInnerHtml $form.InnerHtml

    # Exakt der alte Kern: vorhandene Formularfelder behalten,
    # nur userName und password1 überschreiben.
    $fields['userName'] = $UserName
    $fields['password1'] = $Password

    $body = New-FormUrlEncodedBody -Fields $fields

    $postHeaders = $headers.Clone()
    $postHeaders['Referer'] = $loginUri.AbsoluteUri
    $postHeaders['Origin'] = 'https://www.rungps.net'

    Write-Host "POST $postUri"
    Write-Host "POST fields: $($fields.Keys -join ', ')"

    $loginPost = Invoke-WebRequest `
        -Uri $postUri `
        -Method Post `
        -Body $body `
        -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' `
        -WebSession $runGpsSession `
        -Headers $postHeaders `
        -MaximumRedirection 10

    Write-Host "Login POST: HTTP $($loginPost.StatusCode), Length=$($loginPost.Content.Length)"

    $whitePixelUri = Get-WhitePixelUri `
        -Html $loginPost.Content `
        -BaseUri $loginPost.BaseResponse.RequestMessage.RequestUri

    if ([string]::IsNullOrWhiteSpace($whitePixelUri)) {
        $debugFile = Join-Path (Get-Location) 'debug-login-post.html'
        $loginPost.Content | Set-Content -Path $debugFile -Encoding utf8

        throw "Login nicht erfolgreich oder unerwartete Antwort. Kein gps-sport.net/whitePixel.jsp gefunden. Debug: $debugFile"
    }

    Write-Host "Bridge image: $whitePixelUri"

    $bridge = Invoke-WebRequest `
        -Uri $whitePixelUri `
        -WebSession $runGpsSession `
        -Headers $headers `
        -MaximumRedirection 10

    Write-Host "Bridge: HTTP $($bridge.StatusCode), Length=$($bridge.Content.Length)"

    $gpsSportCookies = @($runGpsSession.Cookies.GetCookies([uri]'https://www.gps-sport.net/'))

    if ($gpsSportCookies.Count -eq 0) {
        throw 'Bridge wurde aufgerufen, aber es wurde kein Cookie für gps-sport.net erzeugt.'
    }

    Write-Host "OK: gps-sport.net cookies: $($gpsSportCookies.Count)"

    return [Microsoft.PowerShell.Commands.WebRequestSession]$runGpsSession
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

if ([string]::IsNullOrWhiteSpace($env:RUNGPSUSER)) {
    throw "Secret RUNGPSUSER fehlt oder ist leer."
}

if ([string]::IsNullOrWhiteSpace($env:RUNGPSPASSWORD)) {
    throw "Secret RUNGPSPASSWORD fehlt oder ist leer."
}

$user = $env:RUNGPSUSER
$passwordPlain = $env:RUNGPSPASSWORD

Write-Host "RUNGPSUSER length: $($user.Length)"
Write-Host "RUNGPSUSER sha256-16: $(Get-Sha256Fingerprint -Text $user)"

Write-Host "RUNGPSPASSWORD length: $($passwordPlain.Length)"
Write-Host "RUNGPSPASSWORD sha256-16: $(Get-Sha256Fingerprint -Text $passwordPlain)"

$securePassword = ConvertTo-SecureString $passwordPlain -AsPlainText -Force
$cred = [PSCredential]::new($user, $securePassword)

# $runGPS = Connect-RunGPS -Credential $cred
$runGPS = Connect-RunGpsPwsh7 -Username $user -Password $passwordPlain

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

$RunGPSDataRoot = Join-Path $PSScriptRoot 'RunGPSData'
$RunGPSRoutesPath = Join-Path $RunGPSDataRoot 'Routes'
$RunGPSTrainingsPath = Join-Path $RunGPSDataRoot 'Trainings'

New-Item -ItemType Directory -Force -Path $RunGPSRoutesPath | Out-Null
New-Item -ItemType Directory -Force -Path $RunGPSTrainingsPath | Out-Null

Write-Host "Routes path: $RunGPSRoutesPath"
Write-Host "Trainings path: $RunGPSTrainingsPath"


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

$routes = @($routes)
Write-Host "Routes found: $($routes.Count)"

foreach ($route in $routes) {
    SaveRoutesData -ID $route.ID -Path $RunGPSRoutesPath
}

$trainings = @(Get-Trainings)
Write-Host "Trainings found: $($trainings.Count)"

foreach ($training in $trainings) {
    SaveTrainingsData -ID $training.ID -Path $RunGPSTrainingsPath
}

$trainingsXml = Join-Path $RunGPSDataRoot 'Trainings.xml'
$trainings | Export-Clixml -Path $trainingsXml

$tcxPath = Join-Path $PSScriptRoot 'RunGPSData/Trainings'
$geoJsonPath = Join-Path $PSScriptRoot 'RunGPSData/TrngsGeoJSON'

Convert-AllRunGpsTcxToGeoJson `
    -TcxDirectory $tcxPath `
    -GeoJsonDirectory $geoJsonPath
	
New-RunGpsTrainingsGeoJson `
    -TrainingsTcxPath (Join-Path $PSScriptRoot 'RunGPSData/Trainings') `
    -OutFile (Join-Path $PSScriptRoot 'RunGPSData/Maps/trainings.geojson')
	
<# Rest ignorieren wenn obiges klappt

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

#>
