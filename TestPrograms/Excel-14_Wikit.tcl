# Test CawtExcel procedures to exchange data between Excel and Wikit tables.
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

set outPath [file join [pwd] "testOut"]

set outFileWiki1  [file join $outPath Excel-14_Wikit1.txt]
set outFileWiki2  [file join $outPath Excel-14_Wikit2.txt]
set testFileExcel [file join $outPath Excel-14_Wikit%s]
set outFileExcel1 [file join $outPath Excel-14_Wikit1%s]
set outFileExcel2 [file join $outPath Excel-14_Wikit2%s]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Open a new Excel instance, so we are able to get the extension string.
set appId1 [::Excel::OpenNew true]

set excelExt [::Excel::GetExtString $appId1]

# Delete Excel output file from previous test run.
set xlsTestFile [format $testFileExcel $excelExt]
file delete -force $xlsTestFile
set xlsOutFile1 [format $outFileExcel1 $excelExt]
file delete -force $xlsOutFile1
set xlsOutFile2 [format $outFileExcel2 $excelExt]
file delete -force $xlsOutFile2

# Create an Excel file with some test data.
set workbookId [::Excel::AddWorkbook $appId1]
set headerList { "Col-1" "Col-2" "Col-3" "Col-4" }
set dataList { 
    {"1" "2" "3" "None"}
    {"1.1" "1.2" "1.3" "Dot"}
    {"1,1" "1,2" "1,3" "Comma"}
    {"1|1" "1|2" "1|3" "Pipe"}
    {"1;1" "1;2" "1;3" "Semicolon"}
}

set worksheetId [::Excel::AddWorksheet $workbookId "WikitTest"]
::Excel::SetHeaderRow $worksheetId $headerList
::Excel::SetMatrixValues $worksheetId $dataList 2

puts "Copy worksheet to Wikit file $outFileWiki1"
::Excel::WorksheetToWikitFile $worksheetId $outFileWiki1 true

::Excel::SaveAs $workbookId $xlsTestFile
::Excel::Close $workbookId
::Excel::Quit $appId1

puts "Copy Wikit file $outFileWiki1 to Excel worksheet"
set appId2 [::Excel::OpenNew true]
set workbookId  [::Excel::AddWorkbook $appId2]
set worksheetId [::Excel::AddWorksheet $workbookId "WikitTable"]
::Excel::WikitFileToWorksheet $outFileWiki1 $worksheetId true
::Excel::SaveAs $workbookId $xlsOutFile1

puts "Copy Wikit file $outFileWiki1 to Excel file"
set appId3 [::Excel::WikitFileToExcelFile $outFileWiki1 $xlsOutFile2 true false]

puts "Copy Excel file $xlsOutFile1 to Wikit file"
::Excel::ExcelFileToWikitFile $xlsOutFile1 $outFileWiki2 "WikitTable" true true

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId2
    ::Excel::Quit $appId3
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
