# Test basic functionality of the CawtOcr package.
#
# Copyright: 2007-2013 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

set retVal [catch {package require cawt} pkgVersion]

set appId [::Ocr::Open]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

if { [lindex $argv 0] eq "auto" } {
    ::Ocr::Close $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
