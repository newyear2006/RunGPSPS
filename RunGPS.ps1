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
