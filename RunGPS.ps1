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
