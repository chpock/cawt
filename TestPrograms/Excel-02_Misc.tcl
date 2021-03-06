# Test miscellaneous CawtExcel procedures like setting colors, fonts and column width,
# inserting formulas, hyperlinks and images, searching and page setup. 
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

# Number of test rows and columns being generated.
set numRows  10
set numCols   3

# Generate row list with test data
for { set i 1 } { $i <= $numCols } { incr i } {
    lappend rowList $i
}

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-02_Misc"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Select the first - already existing - worksheet, 
# set its name and fill it with data.
set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId 1]
::Excel::SetWorksheetName $worksheetId "ExcelMisc"

for { set row 1 } { $row <= $numRows } { incr row } {
    ::Excel::SetRowValues $worksheetId $row $rowList
}

# Use different range selection procedures and test various
# formatting and color procedures.
set rangeId [::Excel::SelectCellByIndex $worksheetId 2 1 true]
::Excel::SetRangeFillColor $rangeId 255 0 0
::Excel::SetRangeTextColor $rangeId 0 255 0
::Excel::SetRangeBorders $rangeId

set rangeId [::Excel::SelectCellByIndex $worksheetId 3 1 true]
::Excel::SetRangeFillColor $rangeId 0 255 0
::Excel::SetRangeTextColor $rangeId 0 0 255

set rangeId [::Excel::SelectCellByIndex $worksheetId 4 1 true]
::Excel::SetRangeFillColor $rangeId 0 0 255
::Excel::SetRangeTextColor $rangeId 255 0 0
::Excel::SetRangeBorders $rangeId $::Excel::xlThick

set rangeId [::Excel::SelectRangeByIndex $worksheetId 5 1 5 1 true]
::Excel::SetRangeFillColor $rangeId 255 0 0
::Excel::SetRangeTextColor $rangeId 0 255 0

set rangeId [::Excel::SelectRangeByIndex $worksheetId 6 1 7 2 true]
::Excel::SetRangeFillColor $rangeId 0 255 0
::Excel::SetRangeTextColor $rangeId 0 0 255
::Excel::SetRangeBorders $rangeId $::Excel::xlThin $::Excel::xlDash

set rangeId [::Excel::SelectRangeByString $worksheetId "A8:C10" true]
::Excel::SetRangeFillColor $rangeId 0 0 255
::Excel::SetRangeTextColor $rangeId 255 0 0

::Excel::SetRangeFormat $rangeId "real" [::Excel::GetLangNumberFormat "0" "000"]

# Test setting a formula.
set cell [::Excel::SelectCellByIndex $worksheetId 1 [expr $numCols + 2] true]
$cell Formula "=TODAY()"
puts "Formula:      [$cell Formula]"
puts "FormulaLocal: [$cell FormulaLocal]"

# Generate a text file for testing the hyperlink capabilities.
set fileName [file join [pwd] "testOut" "Excel-02_Misc.txt"]
set fp [open $fileName "w"]
puts $fp "This is the linked text file."
close $fp

::Excel::SetHyperlink $worksheetId 2 [expr $numCols + 2] \
                      [format "file://%s" $fileName] "Hyperlink"

# Test the search capabilities.
::Excel::SetCellValue $worksheetId 3 [expr $numCols + 2] "Hallo"
::Excel::SetCellValue $worksheetId 4 [expr $numCols + 2] "Holla"

set rangeId [::Excel::SelectCellByIndex $worksheetId 3 [expr $numCols + 2] true]
::Excel::SetRangeFontBold $rangeId true
::Excel::SetRangeBorders $rangeId $::Excel::xlThin $::Excel::xlContinuous 255 0 0
set rangeId [::Excel::SelectCellByIndex $worksheetId 4 [expr $numCols + 2] true]
::Excel::SetRangeFontItalic $rangeId true
::Excel::SetRangeBorders $rangeId $::Excel::xlThin $::Excel::xlContinuous 0 0 255

# Search only first 20 rows and columns for an existing string.
set str "Hallo"
set cell [::Excel::Search $worksheetId $str 1 1 20 20]
if { [llength $cell] == 2 } {
    set rowNum [lindex $cell 0]
    set colNum [lindex $cell 1]
    puts "Found string \"$str\" at cell [::Excel::ColumnIntToChar $colNum]$rowNum."
} else {
    puts "Error: Could not find string \"$str\"."
}

# Search only first 20 rows and columns for a non-existing string.
set str "HalliHallo"
set cell [::Excel::Search $worksheetId $str 1 1 20 20]
if { [llength $cell] == 0 } {
    puts "Did not find string \"$str\"."
} else {
    puts "Error: String \"$str\" should not be available in worksheet."
}

# Search whole worksheet for an existing string.
set str "Holla"
set cell [::Excel::Search $worksheetId $str]
if { [llength $cell] == 2 } {
    set rowNum [lindex $cell 0]
    set colNum [lindex $cell 1]
    puts "Found string \"$str\" at cell [::Excel::ColumnIntToChar $colNum]$rowNum."
} else {
    puts "Error: Could not find string \"$str\"."
}

# Test different ways of setting column width.
# Set all used colums to fit, except columns 1 and 2.
::Excel::SetColumnsWidth $worksheetId 1 [expr $numCols + 6] 0
::Excel::SetColumnWidth $worksheetId 1 20
::Excel::SetColumnWidth $worksheetId 2 10

# Test inserting and scaling an image into a worksheet.
set picId [::Excel::InsertImage $worksheetId [file join [pwd] "testIn/wish.gif"] 5 9]
::Excel::ScaleImage $picId 2 2.5

# Test copying a whole worksheet.
set copyWorksheetId [::Excel::AddWorksheet $workbookId "WorksheetCopy"]
::Excel::CopyWorksheet $worksheetId $copyWorksheetId

# Adjust the page setup of the worksheets.
::Excel::SetWorksheetOrientation $worksheetId $::Excel::xlLandscape
::Excel::SetWorksheetZoom $worksheetId 50

::Excel::SetWorksheetOrientation $copyWorksheetId $::Excel::xlPortrait
::Excel::SetWorksheetFitToPages $copyWorksheetId

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
