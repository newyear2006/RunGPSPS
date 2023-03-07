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
                      ID=[int]$htmlRoute.children[2].childNodes[0].pathname.substring($htmlRoute.children[2].childNodes[0].pathname.LastIndexOf("_")+1);
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
    {$km -ge 5.0} {'5000';break;}
    {$km -ge 1.0} {'1000';break;}
    default {'0'}
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
                      HöheMin=[int]$htmlTraining.children[10].innerText;
                      HöheMax=[int]$htmlTraining.children[11].innerText;
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
      $r=Invoke-WebRequest -WebSession $runGPS -Uri "http://www.gps-sport.net/userTrainings.jsp?userName=$($user)&startDate=$($FromDate.toString('yyyy-MM-dd'))&endDate=$($ToDate.toString('yyyy-MM-dd'))&sport=$($Sport)&submitButton=Aktualisieren#" -ContentType "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
      If ($?) {
	  $htmlTrainings=$r.ParsedHtml.body.childNodes[2].childNodes[0].childNodes[1].childNodes
	Write-Verbose "Konvertiere $($htmlTrainings.Length) Einträge"
	  $newtrainings=$htmlTrainings | % {$t=@()} {$t+=NewTraining $_} {$t}
	If ($htmlTrainings.Length -lt 1000) {
          $loop = $false
        } else {
          $FromDate = $newTrainings[0].Datum
        }
        $Trainings += $newTrainings
      }
    }
  $Trainings | sort Datum
}

Function Connect-RunGPS {
  [CmdletBinding()]
  Param (
  	[PSCredential]$Credential
  )
  
  $uri = "https://www.rungps.net/login.jsp"

  # [Microsoft.PowerShell.Commands.WebRequestSession]$runGPS = $null
  $l=Invoke-WebRequest -Uri $uri -SessionVariable runGPS
  $l.Forms[0].Fields["userName"]=$Credential.Username
  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
  $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)  
  $l.Forms[0].Fields["password1"]=$UnsecurePassword
  # $l.Forms[0].Fields
  $r=Invoke-WebRequest -Uri ('https://www.rungps.net/' + $l.Forms[0].Action) -Method Post -Body $l -WebSession $runGPS
  #$runGPS.Cookies.GetCookies($uri)
  # Dieses Image muss aufgerufen werden, damit die passenden Cookies für den Domainwechsel vorhanden sind!
  # Der Domainwechsel findet zwischen rungps.net und gps-sport.net statt!
  $wp=Invoke-WebRequest -WebSession $runGPS -Uri $r.ParsedHtml.images[0].href 
  # wenn alles geklappt hat meldet $wp.Content: 0A 0A 0A 3C 21 2D 2D 4F 4B 2D 2D 3E 0A 0A 0A     ...<!--OK-->...
  If ($wp.Content -eq "`n`n`n<!--OK-->`n`n`n") {
    [Microsoft.PowerShell.Commands.WebRequestSession]$runGPS
  }
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
$cred=Get-Credential -Message 'Bitte Zugangsdaten für RunGPS-Anmeldung eingeben'
$runGPS = Connect-RunGPS -Credential $cred

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
$trainings=Get-Trainings

# ein Training zum Testen als GPX-Datei speichern 
SaveTrainingGPSData -ID $trainings[0].ID -FileType GPX -Path C:\Temp\RunGPS

# alle Trainings von der Abfrage speichern
$trainings | % {SaveTrainingsData -ID $_.ID -Path C:\Temp\RunGPS }

# Auswertung alle Wanderungen nach Monaten kummuliert ausgeben
$trainings | where Sportart -eq 'Wandern'| select *, @{N='JahrMonat';E={$_.Datum.tostring("yyyy-MM")}}|group -Property JahrMonat|select count, @{N='JahrMonat';E={$_.Name}}, @{N='Km';E={$_.Group|% {$km=0} {$km+=$_.Distanz} {$km}}}

# Auswertung alle Wanderungen nach Jahren kumuliert ausgeben
$trainings | where Sportart -eq 'Wandern'| select *, @{N='Jahr';E={$_.Datum.Year}}|group -Property Jahr|select count, @{N='Jahr';E={$_.Name}}, @{N='Km';E={$_.Group|% {$km=0} {$km+=$_.Distanz} {$km}}}

# Ausgabe der gesamten Trainingszeit
New-Timespan -seconds ($trainings|select -ExpandProperty dauer| measure -Sum -Property totalseconds).sum


# Trainings speichern
$trainings|Export-Clixml -Path c:\temp\RunGPS\Trainings.xml

# Trainings wieder laden
$trainings=Import-Clixml -Path .\Trainings.xml
