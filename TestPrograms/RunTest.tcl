# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { $argc < 1 } {
    puts ""
    puts "Usage: $argv0 Namespace \[CoverageOnly\]"
    puts ""
    puts "Perform the tests and check code coverage for specified namespace."
    puts "Namespaces available: Earth, Excel, Explorer, Ocr, Matlab, Ppt, Word"
    puts "If CoverageOnly is set to 1, only coverage checks are performed."
    exit 1
}

set nsName  [lindex $argv 0]
set nsLower [string tolower $nsName]

set coverageOnly 0
if { $argc > 1 } {
    set coverageOnly [lindex $argv 1]
}

if { $::tcl_platform(platform) eq "windows" } {
    set tclsh "tclsh.exe"
} else {
    set tclsh "tclsh"
}

proc runTest { testFile } {
    puts "Running test $testFile ..."
    exec $::tclsh $testFile auto
}

if { ! $coverageOnly } {
    catch { file mkdir testOut }
    foreach f [lsort [glob ${nsName}-*]] {
        runTest $f
    }
}

puts ""
puts "Checking $nsName test coverage ..."

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

set allProcList [lsort [info commands ${nsName}::*]]
foreach cmd $allProcList {
    if { ! [string match "*Obsolete:*" [info body $cmd]] } {
        lappend procList $cmd
    }
}

# We search the test scripts as well as the implementation files,
# as a procedure may be used by a higher-level procedure and thus 
# does not have to be tested separately.
set testFileList [lsort [glob "${nsName}-*.tcl" "../Cawt${nsName}/${nsLower}*.tcl"]]

foreach testFile $testFileList {
    puts "Scanning testfile $testFile"
    set fp [open $testFile "r"]
    while { [gets $fp line] >= 0 } {
        foreach cmd $procList {
            if { [string match "*${cmd}*" $line] } {
                #puts "Found proc $cmd in file $testFile"
                set found($cmd) 1
            }
        }
    }
    close $fp
}

set foundList [lsort [array names found *]]
foreach cmd $procList {
    if { [lsearch $foundList $cmd] < 0 } {
        puts "$cmd not yet tested"
    }
}

set numObsolete [expr [llength $allProcList] - [llength $procList]]
puts "[llength $procList] procedures checked ($numObsolete obsolete)"

exit 0
