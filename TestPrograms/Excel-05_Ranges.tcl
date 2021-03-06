# Test CawtExcel procedures for retrieving the number of (used) rows and columns.
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-05_Ranges"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Test large row and column numbers.
set worksheetId [::Excel::AddWorksheet $workbookId "LargeColumn"]
set maxRows [::Excel::GetNumRows $worksheetId]
set maxCols [::Excel::GetNumColumns $worksheetId]
puts "Maximum number of rows   : $maxRows"
puts "Maximum number of columns: $maxCols"
# 256 and 65535 are the maximum number of columns and rows in Excel
# versions up to Excel 2003.
if { $maxCols > 256 } {
    set maxCols 500
}
if { $maxRows > 65536 } {
    set maxRows 70000
}
puts "Using as maximum row value   : $maxRows"
puts "Using as maximum column value: $maxCols ([::Excel::ColumnIntToChar $maxCols])"

::Excel::SetCellValue $worksheetId 1 1 "Cell-1-1"
::Excel::SetCellValue $worksheetId 1 $maxCols "Cell-1-$maxCols"
::Excel::SetCellValue $worksheetId $maxRows 1 "Cell-$maxRows-1"
::Excel::SetCellValue $worksheetId $maxRows $maxCols "Cell-$maxRows-$maxCols"

::Excel::ShowCellByIndex $worksheetId $maxRows $maxCols

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
