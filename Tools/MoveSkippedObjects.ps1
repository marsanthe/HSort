
using namespace System.Collections.Generic

$Directory = Split-Path -Parent(Get-Location) # Parent( (Get-Location == Tools) ) == HSort 


# To import module, Script has to be executed from .\HSort\Tools !
Import-Module -Name "$Directory\Modules\InitializeScript\InitializeScript.psm1"


<# 
    .SYNOPSIS
        Move skipped objects to a folder.
    .DESCRIPTION
        Move skipped objects to a folder named "Skipped [LibraryName]".
        The skipped objects are sorted into different sub-folders named after their parent folder.
    .INPUTS
        LibraryName
            The name of the library.
            To move the files that were skipped during the creation of this library.
        TargetDir
            The Directory to create "Skipped [LibraryName]" in. 
#>

$SkippedDir = "$HOME\AppData\Roaming\HSort\SkippedObjects"
$VisitedDir = "$HOME\AppData\Roaming\HSort\ApplicationData"

Show-Information -InformationArray (" ",
"INFORMATION",
"==========================",
"This script is meant to be run after a Library was created from a Source-folder.",
"If objects from this Source-folder were skipped, this script will MOVE them to a folder named:",
"SkippedObjects from [YourLibraryName]",
"Each object is in turn placed into a folder named by the object's parent-folder.",
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


$LibraryName = Read-Host "Please enter the name of the library in question"
    
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



$TargetDir = Read-Host "`nWhere would you like to create the SkippedObjects folder?`n"

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


$TotalMoved = 0
$TotalMoveError = 0

$MovedObjects = [List[object]]::new()
$MoveErrorObjects = [List[object]]::new()

Show-Information -InformationArray (" ","Please wait. Moving objects..."," ")


foreach($Parent in $SkippedXML.Keys){

    $ParentName = Split-Path -Path $Parent -Leaf

    $null = New-Item -Name $ParentName -ItemType "directory" -Path $PathsSkipped.Base

    $ObjTargetDir = "$($PathsSkipped.Base)\$ParentName"

    foreach ($Object in $SkippedXML.$Parent.Keys) {

        $ErrorFlag = 0
        
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
                $TotalMoveError += 1
                $ErrorFlag = 1
            }

            if($ErrorFlag -eq 0){
                $VisitedObjects.Remove($Object)
                $MovedObjects.Add($Object)
            }
            else{
                $MoveErrorObjects.Add($Object)
            }

        }
        else{

            try{
                $null = robocopy $ParentDir $ObjTargetDir $Name /move
                $TotalMoved += 1
            }
            catch {
                Write-Information "[Error moving] $Name" -InformationAction Continue
                $TotalMoveError += 1
                $ErrorFlag = 1
            }

            if ($ErrorFlag -eq 0) {
                $VisitedObjects.Remove($Object)
                $MovedObjects.Add($Object)
            }
            else {
                $MoveErrorObjects.Add($Object)
            }

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

Show-Information -InformationArray (" ", "Finished moving objects", " ")

$Head = @"

Summary of moved Objects
=================================================

[Based on the last scan of: $LibraryName]

# Total objects moved: $TotalMoved
=================================================

"@

$Head | Out-File -FilePath "$($PathsSkipped.Base)\MovedObjects Log.txt" -Encoding unicode -Append -Force

foreach ($Object in $MovedObjects) {
    $S = "$Object"
    $S | Out-File -FilePath "$($PathsSkipped.Base)\MovedObjects Log.txt" -Encoding unicode -Append -Force
}



$Subhead = @"

# Objects not moved: $TotalMoveError
=================================================

"@

$Subhead | Out-File -FilePath "$($PathsSkipped.Base)\MovedObjects Log.txt" -Encoding unicode -Append -Force

foreach ($Object in $MoveErrorObjects) {
    $S = "$Object"
    $S | Out-File -FilePath "$($PathsSkipped.Base)\MovedObjects Log.txt" -Encoding unicode -Append -Force
}
