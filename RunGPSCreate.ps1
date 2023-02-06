# RunGPS.ps1 muss vorher geladen sein!

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
