# ConvertTo-Calibre.ps1

$calibreData = calibredb list --fields all --for-machine
$jsonData = $calibreData | convertfrom-json
$csv = $jsonData | Select-Object title,authors,pubdate,isbn,publisher | ConvertTo-CSV -NoTypeInformation
$csv | Out-File -Encoding utf8 -FilePath ~/Downloads/ToLibraryThing.csv
Invoke-Item ~/Downloads/ToLibraryThing.csv
