# Test CawtExcel procedures to exchange data between Excel and the Windows clipboard.
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

set outPath [file join [pwd] "testOut"]

set outFile [file join $outPath Excel-15_Clipboard]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Open a new Excel instance, so we are able to get the extension string.
set appId [::Excel::OpenNew true]

set excelExt [::Excel::GetExtString $appId]

# Delete Excel output file from previous test run.
set xlsOutFile [format $outFile $excelExt]
file delete -force $xlsOutFile

# Create an Excel file with some test data.
set workbookId [::Excel::AddWorkbook $appId]
set headerList { "Col-1" "Col-2" "Col-3" "Col-4" }
set dataList { 
    {"1" "2" "3" "None"}
    {"1.1" "1.2" "1.3" "Dot"}
    {"1,1" "1,2" "1,3" "Comma"}
    {"1|1" "1|2" "1|3" "Pipe"}
    {"1;1" "1;2" "1;3" "Semicolon"}
}

set worksheetId1 [::Excel::AddWorksheet $workbookId "ClipboardSource"]
::Excel::SetHeaderRow $worksheetId1 $headerList
::Excel::SetMatrixValues $worksheetId1 $dataList 2

puts "Copy worksheet to clipboard"
::Excel::WorksheetToClipboard $worksheetId1 1 1  \
    [::Excel::GetNumUsedRows $worksheetId1] \
    [::Excel::GetNumUsedColumns $worksheetId1]

set worksheetId2 [::Excel::AddWorksheet $workbookId "ClipboardDest"]

puts "Copy clipboard to worksheet with offset"
::Excel::ClipboardToWorksheet $worksheetId2 3 2

::Excel::SaveAs $workbookId $xlsOutFile

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
