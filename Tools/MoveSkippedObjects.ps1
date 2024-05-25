
using namespace System.Collections.Generic

$Directory = Split-Path -Parent(Get-Location) # Parent( (Get-Location == Tools) ) == HSort 


# To import module, Script has to be executed from .\HSort\Tools !
Import-Module -Name "$Directory\Modules\InitializeScript\InitializeScript.psm1"


<# 
    .SYNOPSIS
        Move skipped objects to a folder.
    .DESCRIPTION
        Move skipped objects to a folder named "Skipped [LibraryName]".
        The skipped objects are sorted into different sub-folders depending on
        the reason they were skipped.
    .INPUTS
        LibraryName
            The name of the library.
            To move the files that were skipped during the creation of this library.
        TargetDir
            The Directory to create "Skipped [LibraryName]" in. 
#>

$SkippedDir = "$HOME\AppData\Roaming\HSort\SkippedObjects"
$VisitedDir = "$HOME\AppData\Roaming\HSort\ApplicationData"

Show-Information -InformationArray ("INFORMATION",
        "==========================","This script is meant to be run after a Library was created from a Source-folder.",
"If objects from this Source-folder were skipped, this script will MOVE them to a folder of your choice and sort them.",
"This helps you to get an overview over Objects that you might want to add to the library manually,",
"(like Manga not following the E-Hentai naming-scheme), or sort otherwise."," ")

$UserInput = Get-UserInput -Dialog @{
    "Question" = "Would you like to continue?";
    "YesResponse" = " ";
    "NoResponse" = "Exiting"
}

if ($UserInput -eq "n"){
    exit
}


$LibraryName = Read-Host "Please enter the name of the library."
    
try {
    $SkippedXML = Import-Clixml -Path "$SkippedDir\Skipped $LibraryName.xml"
    
}
catch {
    Write-Information "$SkippedDir\Skipped $LibraryName.xml not found. Exiting"
    exit
}

try {
    $VisitedObjects = Import-Clixml -Path "$VisitedDir\VisitedObjects $LibraryName.xml"
}
catch {
    Write-Information "$VisitedDir\VisitedObjects $LibraryName.xml not found. Exiting"
    exit
}
        
$TargetDir = Read-Host "`nWhere would you like to save the folder containing the skipped objects?`n"

while ($True) {
    if (!(Test-Path $TargetDir)) {
        $TargetDir = Read-Host "This directory doesn't exist. Please enter a different path."
    }
    elseif (Test-Path "$TargetDir\Skipped $LibraryName") {
        $TargetDir = Read-Host "This directory already contains a folder named: Skipped $LibraryName. Please enter a different path."
    }
    else{
        break
    }
}


$PathsSkipped = [ordered]@{
    "Base"         = "$TargetDir\SkippedObjects from $LibraryName";
    # "NoMatch"      = "$TargetDir\SkippedObjects from $LibraryName\NoMatch";
    # "Duplicates"   = "$TargetDir\SkippedObjects from $LibraryName\Duplicates";
    # "HashMismatch" = "$TargetDir\SkippedObjects from $LibraryName\HashMismatch";
    # "Other"        = "$TargetDir\SkippedObjects from $LibraryName\Other"
}


foreach ($Path in $PathsSkipped.Keys) {
    if (!(Test-Path $PathsSkipped.$Path)) {
        $null = New-Item -ItemType "directory" -Path $PathsSkipped.$Path
    }
}

$ErrorCounter = 0
$TotalMoved = 0

$MovedObjects = [List[object]]::new()

Show-Information -InformationArray ("Please wait. Moving objects...`n")


foreach($Parent in $SkippedXML.Keys){

    $ParentName = Split-Path -Path $Parent -Leaf

    $null = New-Item -Name $ParentName -ItemType "directory" -Path $PathsSkipped.Base

    $ObjTargetDir = "$($PathsSkipped.Base)\$ParentName"

    foreach ($Object in $SkippedXML.$Parent.Keys) {
        
        $Source = $SkippedXML.$Parent.$Object.Path
        $Name = $SkippedXML.$Parent.$Object.ObjectName # The object name with extension
        $ParentDir = $SkippedXML.$Parent.$Object.ObjectParent
        
        Write-Information -MessageData "[Moving] $Name" -InformationAction Continue

        if ($SkippedXML.$Parent.$Object.Extension -eq "Folder") {

            try{ 
                $null = robocopy $Source "$ObjTargetDir\$Name" /E /DCOPY:DAT /move
                $TotalMoved += 1
            }
            catch {
                Write-Information "[Error moving] $Name" -InformationAction Continue
                $ErrorCounter += 1
                break
            }

            $VisitedObjects.Remove($Object)
            $MovedObjects.Add($Object)
            break 
        }
        else{
                
            try{
                $null = robocopy $ParentDir $ObjTargetDir $Name /move
                $TotalMoved += 1
            }
            catch {
                $ErrorCounter += 1
                Write-Information "[Error moving] $Name" -InformationAction Continue
                break
            }

            $VisitedObjects.Remove($Object)
            $MovedObjects.Add($Object)
            break 
        }

    }
}


# Update this instance of the Skipped hashtable |---
foreach ($Parent in $SkippedXML.Keys) {
    foreach ($Object in $MovedObjects) {
        $SkippedXML.$Parent.Remove($Object)
    }
}

# ---> Export and overwrite old Skipped hashtable
$SkippedXML | Export-Clixml -Path "$SkippedDir\Skipped $LibraryName.xml" -Force

$Summary = @"

Finished moving objects
=================================================

[Based on the last scan of: $LibraryName]

# Total objects moved: $TotalMoved
=================================================


# Errors: $ErrorCounter
=================================================

"@

$Summary | Out-Host
$Summary | Out-File -FilePath "$($PathsSkipped.Base)\MovedObjects Log.txt" -Encoding unicode -Force
