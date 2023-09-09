# ConvertTo-Calibre.ps1

# set up a few variables for later
$DownloadsFolderPath   = "~/Downloads/ToLibraryThing.csv"
$LibraryThingImportURL = "https://www.librarything.com/import" 
$calibreDBHelpURL      = "https://manual.calibre-ebook.com/generated/en/calibredb.html"

# check for calibredb
try {
	"Checking for Calibre..."
	calibredb.exe --version
}
catch {
	"Calibre command line database interface not found"
	"Typing calibredb.exe --version in a PowerShell window has to work for this script to function"
	"Please confirm Calibre installed and the calibredb command on your PATH" 
	"See $calibreDBHelpURL for more info on the command"
	exit 1
}

# export library to json
$calibreData = calibredb.exe list --fields all --for-machine

# convert json to object
$jsonData = $calibreData | ConvertFrom-JSON

# grab just four fields from object, converting the pubdate to a yyyy-MM-dd format
$trimmedData = $jsonData | Select-Object title,authors,pubdate,isbn,publisher 

# convert object to csv
$csvData = $trimmedData | ConvertTo-CSV -NoTypeInformation

# save csv in your downloads folder
$csvData | Out-File -Encoding utf8 -FilePath $DownloadsFolderPath
" Saved CSV file in your Downloads folder "
" Please review for sanity before you upload to Library Thing "

# open the csv with default editor for further review
" Opening csv file viewer "
Invoke-Item $DownloadsFolderPath 

# open the Library Thing Import Page
" Opening Library Thing Import Page "
" This may not work if you aren't already logged into Library Thing "
" You can just go $LibraryThingImportURL manually after logging in "
Invoke-Item $LibraryThingImportURL
