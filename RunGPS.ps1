# RunGPS-Powershellroutinen

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
                      ID=$htmlRoute.children[2].childNodes[0].pathname.substring($htmlRoute.children[2].childNodes[0].pathname.LastIndexOf("_")+1);
                      Titel=$htmlRoute.children[2].innerText;
                      Distanz=[decimal]$htmlRoute.children[3].innerText;
                      Läufe=[int]$htmlRoute.children[4].innerText;
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
    Param(
        [String]$ID,
        [String]$FileType = "GPX",
        [String]$Path
    )
    
    $SaveFile = Join-Path -Path $Path -ChildPath "Route-$($ID).$FileType"

    Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/routePlanner/dlServices/$($FileType.ToLower()).jsp?routeID=$ID" -OutFile $SaveFile

}

# parst einen übergebenen HTML-String und liest relevante Trainingsdaten aus und gibt diese als Objekt zurück
function NewTraining {
    Param (
        $htmlTraining
    )
    $t=[PSCustomObject]@{Datum=(Get-Date $htmlTraining.children[0].innerhtml);
                      Sportart=($htmlTraining.children[1].innerText.Trim());
                      ID=$htmlTraining.children[2].childNodes[0].pathname.substring($htmlTraining.children[2].childNodes[0].pathname.LastIndexOf("_")+1);
                      Titel=$htmlTraining.children[2].innerText;
                      Distanz=$htmlTraining.children[3].innerText;
                      Dauer=[TimeSpan]$htmlTraining.children[4].innerText;
                      Kalorien=[int]$htmlTraining.children[5].innerText;
                      HerzfrequenzD=[decimal]$htmlTraining.children[6].innerText;
                      TrittfrequenzD=[decimal]$htmlTraining.children[7].innerText;
                      GeschwindigkeitD=[decimal]$htmlTraining.children[8].innerText;
                      GeschwindigkeitDA=[decimal]$htmlTraining.children[9].innerText
                      HöheMin=[int]$htmlTraining.children[10].innerText
                      HöheMax=[int]$htmlTraining.children[11].innerText
                      Abstieg=[int]$htmlTraining.children[12].innerText
                      Aufstieg=[int]$htmlTraining.children[13].innerText
                      Gewicht=[int]$htmlTraining.children[14].innerText
                      }
    
    $t
}

# $runGPS muss existieren!
Function SaveTrainingsData {
    [Cmdletbinding()]
    Param(
        [String]$ID,
        [String]$Path
    )

    Write-Verbose "Saving $ID"
    SaveTrainingGPSData -ID $ID -FileType GPX -Path $Path
    SaveTrainingGPSData -ID $ID -FileType KML -Path $Path
    SaveTrainingGPSData -ID $ID -FileType TCX -Path $Path

}

# $runGPS muss existieren!
Function SaveTrainingGPSData {
    Param(
        [String]$ID,
        [String]$FileType = "GPX",
        [String]$Path
    )

    $SaveFile = Join-Path -Path $Path -ChildPath "$($ID).$FileType"

    Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/services/training$($FileType).jsp?trainingID=$ID" -OutFile $SaveFile

}

# zum Prüfen, was auf der anderen Seite ankommt: https://pipedream.com/requestbin 
# früher: http://requestb.in/

####### Benutzer und Passwort müssen hinterlegt werden!!
$user='Benutzer'
$password='Passwort'
#######
$uri = "https://www.rungps.net/login.jsp"

$l=Invoke-WebRequest -Uri $uri -SessionVariable runGPS
$l.Forms[0].Fields["userName"]=$user
$l.Forms[0].Fields["password1"]=$password
$l.Forms[0].Fields
$r=Invoke-WebRequest -Uri ('https://www.rungps.net/' + $l.Forms[0].Action) -Method Post -Body $l -WebSession $runGPS
#$runGPS.Cookies.GetCookies($uri)
# Dieses Image muss aufgerufen werden, damit die passenden Cookies für den Domainwechsel vorhanden sind!
# Der Domainwechsel findet zwischen rungps.net und gps-sport.net statt!
$wp=Invoke-WebRequest -WebSession $runGPS -Uri $r.ParsedHtml.images[0].href 

# Datumsangaben im ISO-Format YYYY-MM-DDD, gibt maximal 1000 Einträge zurück, evtl. muss der Datumsbereich durch zwei
# oder mehrere Aufrufe gesplittet werden!
$r2=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/userRoutes.jsp?userName=$($User)&startDate=2006-10-18&endDate=2023-02-06&sport=&searchTerm=&submitButton=Aktualisieren#" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

$r3=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=2006-10-26&endDate=2017-12-31&sport=&submitButton=Aktualisieren#" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

$r4=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=2018-01-01&endDate=2023-02-06&sport=&submitButton=Aktualisieren#" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"

### ROUTEN
# Routes
$htmlRoutes=$r2.ParsedHtml.body.childNodes[1].childNodes[0].childNodes[13].childNodes[1].childnodes
# Anzahl
$htmlRoutes.length
$routes=$htmlRoutes | % {$r=@()} {$r+=NewRoute $_} {$r}

# erste Route zum Testen speichern
SaveRouteGPSData -ID $routes[0].ID -FileType GPX -Path C:\Temp\RunGPS

# alle ermittelten Routen speichern
$routes | % {SaveRoutesData -ID $_.ID -Path C:\Temp\RunGPS }

### TRAININGS
# Trainings
$htmlTrainings=$r3.ParsedHtml.body.childNodes[2].childNodes[0].childNodes[1].childNodes
# Anzahl
$htmlTrainings.length
$trainings=$htmlTrainings | % {$t=@()} {$t+=NewTraining $_} {$t}

# ein Training zum Testen speichern
SaveTrainingGPSData -ID $trainings[0].ID -FileType GPX -Path C:\Temp\RunGPS

# alle Trainings von der ersten Abfrage speichern
$trainings | % {SaveTrainingsData -ID $_.ID -Path C:\Temp\RunGPS }

# Trainings
$htmlTrainings=$r4.ParsedHtml.body.childNodes[2].childNodes[0].childNodes[1].childNodes
# Anzahl
$htmlTrainings.length
$trainings=$htmlTrainings | % {$t=@()} {$t+=NewTraining $_} {$t}

# alle Trainings von der zweiten Abfrage speichern
$trainings | % {SaveTrainingsData -ID $_.ID -Path C:\Temp\RunGPS }

