<#
.SYNOPSIS
	Help get book data from Calibre to Library Thing
.DESCRIPTION
	Exports JSON from the calibre database and converts it to csv files useful with Library Thing
.PARAMETER FileNamePrefix
	If specified, will use the as the provided prefix in the output csv file names
.PARAMETER FilePath
	If specified, will output csv files into that directory
.EXAMPLE
	PS> ./ConvertFrom-CalibreToLibraryThing.ps1
	Launches calibredb and exports all entries into two csv files, saved in your Downloads folder, then opens the Library Thing website.
.EXAMPLE
	PS> ./ConvertFrom-CalibreToLibraryThing.ps1 --SkipBrowser
	Launches calibredb and exports all entries into two csv files, saved in your Downloads folder, but does not open the Library Thing website.
.EXAMPLE
	PS> ./ConvertFrom-CalibreToLibraryThing.ps1 --SingleFile
	Launches calibredb and exports all entries into one csv file, saved in your Downloads folder
.EXAMPLE
	PS> ./ConvertFrom-CalibreToLibraryThing.ps1 --FilePath ./Out
	Launches calibredb and exports all entries into two csv files, saved in a folder called Out in the same directory as this conversion script
.EXAMPLE
	PS> ./ConvertFrom-CalibreToLibraryThing.ps1 --FileName CollectionA
	Launches calibredb and exports all entries into two csv files, named CollectionA-NoISBN-{currentdatetime}.csv and CollectionA-ISBNOnly-{currentdatetime}.csv
.LINK
	https://github.com/furicle/ConvertFrom-CalibreToLibraryThing
.NOTES
	Author: furicle | License: GPL-3
#>

# set up required variables
[CmdletBinding()]
param(
	[string]$FileNamePrefix  = "Import To LibraryThing",
	[string]$FilePath        = "~/Downloads/",
	[switch]$SkipBrowser
	)

# These variables aren't likely to be changed as an option
[string]$dateStamp             = Get-Date -Format yyyy-MM-dd_HH-mm-ss
[string]$csvFileName           = "$FileNamePrefix $dateStamp.csv"
[string]$LibraryThingImportURL = "https://www.librarything.com/import"
[string]$calibreDBHelpURL      = "https://manual.calibre-ebook.com/generated/en/calibredb.html"



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
$csvData | Out-File -Encoding utf8 -FilePath $OutPath/$csvFileName
$count = $csvData.Count

Write-Verbose " "
Write-Verbose " Saved CSV file -> $csvFileName "
Write-Verbose " in folder      -> $OutPath "
Write-Verbose " "
Write-Verbose " It has $count books in it "
Write-Verbose " "
Write-Verbose " You may wish to review the whole file for sanity "
Write-Verbose " before you upload to Library Thing "
Write-Verbose " "

# open the csv with default editor for further review
Write-Verbose " Opening csv file editor "
Write-Verbose " REMEMBER ISBN numbers often start with a zero, so you "
Write-Verbose " MUST import that column as text, not a number, "
Write-Verbose " or your ISBN numbers will be mangled. "

Invoke-Item $OutPath/$csvFileName

# open the Library Thing Import Page
Write-Verbose " Opening Library Thing Import Page "
Write-Verbose " This may not work if you aren't already logged into Library Thing "
Write-Verbose " You can just go $LibraryThingImportURL manually after logging in "

if ( ! $SkipBrowser ) {
	Start-Process $LibraryThingImportURL
} else {
	Write-Verbose " Not opening Library Thing Import Page "
	Write-Verbose " Visit $LibraryThingImportURL after logging in "
	Write-Verbose " to upload your results "
}

Write-Verbose ""
Write-Verbose " -------------"
Write-Verbose " - All done! - "
Write-Verbose " -------------"
Write-Verbose ""


# vim: set filetype=ps1 syntax=ps1 ts=4 sw=4 tw=0 :
