# Test CawtWord procedures for handling text.
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

# Open new Word instance and show the application window.
set appId [::Word::OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-03_Text"]
append wordFile [::Word::GetExtString $appId]
file delete -force $wordFile

set msg1 "This is a italic line of text in italic."
for { set i 0 } { $i < 20 } { incr i } {
    append msg2 "This is a large oops paragraph in bold. "
}

# Create a new document.
set docId [::Word::AddDocument $appId]

# Insert a short piece of text as one paragraph.
set range1 [::Word::AppendText $docId $msg1]
::Word::SetRangeFontItalic $range1 true
::Word::SetRangeHighlightColorByEnum $range1 $::Word::wdYellow
::Word::AppendParagraph $docId

# Insert a longer piece of text as one paragraph.
set range2 [::Word::AppendText $docId $msg2]
::Word::SetRangeFontBold $range2 true
::Word::AppendParagraph $docId

# Generate a text file for testing the hyperlink capabilities.
set fileName [file join [pwd] "testOut" "Word-03_Text.txt"]
set fp [open $fileName "w"]
puts $fp "This is the text file linked from Word."
close $fp

::Word::AppendParagraph $docId
set rangeLink [::Word::AppendText $docId "Dummy"]
::Word::SetHyperlink $docId $rangeLink [format "file://%s" $fileName] "File Link"
::Word::AppendParagraph $docId

# Insert lines of text. When we get to 7 inches from top of the
# document, insert a hard page break.
set pos [::Cawt::InchesToPoints 7]
while { true } {
    ::Word::AppendText $docId "More lines of text."
    ::Word::AppendParagraph $docId
    set endRange [::Word::GetEndRange $docId]
    if { $pos < [$endRange Information $::Word::wdVerticalPositionRelativeToPage] } {
        break
    }
}
$endRange Collapse $::Word::wdCollapseEnd
$endRange InsertBreak [expr int($::Word::wdPageBreak)]
$endRange Collapse $::Word::wdCollapseEnd
set rangeId [::Word::AppendText $docId "This is page 2."]
::Word::AddParagraph $rangeId "after"
set rangeId [::Word::AppendText $docId "There must be two paragraphs before this line."]
::Word::AddParagraph $rangeId "before"

::Word::SetRangeStartIndex $docId $rangeId "begin"
::Word::SetRangeEndIndex   $docId $rangeId 5
$rangeId Select
::Word::PrintRange $rangeId "SetRange and selection: "

# Save document as Word file.
puts "Saving as Word file: $wordFile"
::Word::SaveAs $docId $wordFile

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
