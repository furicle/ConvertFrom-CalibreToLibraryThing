# ConvertTo-Calibre.ps1

# set up a few variables for later
$dateStamp             = Get-Date -Format yyyy-MM-dd_HH-mm-ss
$DownloadsFolderPath   = "~/Downloads/"
$csvFileName           = "Import To LibraryThing $dateStamp.csv"
$LibraryThingImportURL = "https://www.librarything.com/import"
$calibreDBHelpURL      = "https://manual.calibre-ebook.com/generated/en/calibredb.html"

# check for calibredb
try {
    " Checking for Calibre..."
    calibredb.exe --version
}
catch {
    " Calibre command line database interface not found"
    " Typing calibredb.exe --version in a PowerShell window has to work for this script to function"
    " Please confirm Calibre installed and the calibredb command on your PATH"
    " See $calibreDBHelpURL for more info on the command"
    exit 1
}

# export library to json
$calibreData = calibredb.exe list --fields all --for-machine

# convert json to object
$jsonData = $calibreData | ConvertFrom-Json

# grab just four fields from object, converting the pubdate to a yyyy-MM-dd format
$trimmedData = $jsonData | Select-Object title, authors,@{n='pubdate';e={([datetime]$_.pubdate).ToString('yyyy-MM-dd')}}, isbn, publisher

# convert object to csv
$csvData = $trimmedData | ConvertTo-Csv -NoTypeInformation

# save csv in your downloads folder
$csvData | Out-File -Encoding utf8 -FilePath $DownloadsFolderPath/$csvFileName
$count = $csvData.Count
" "
" Saved CSV file -> $csvFileName "
" in folder      -> $DownloadsFolderPath "
" "
" It has $count books in it "
" "
" You may wish to review the whole file for sanity "
" before you upload to Library Thing "
" "

# open the csv with default editor for further review
" Opening csv file editor "
" REMEMBER ISBN numbers often start with a zero, so you "
" MUST import that column as text, not a number, "
" or your ISBN numbers will be mangled. "
Invoke-Item $DownloadsFolderPath/$csvFileName

# open the Library Thing Import Page
" Opening Library Thing Import Page "
" This may not work if you aren't already logged into Library Thing "
" You can just go $LibraryThingImportURL manually after logging in "
Start-Process $LibraryThingImportURL


# vim: set filetype=ps1 syntax=ps1 ts=4 sw=4 tw=0 :
