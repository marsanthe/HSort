
$Directory = Split-Path -Parent(Get-Location) # Parent( (Get-Location == Tools) ) == HSort 


# To import module, Script has to be executed from .\HSort\Tools !
Import-Module -Name "$Directory\Modules\InitializeScript\InitializeScript.psm1"


<# 
    .SYNOPSIS
        Move skipped objects to a folder.
    .DESCRIPTION
        Move skipped objects to a folder named "Skipped [LibraryName]".
        The skipped objects are sorted into different folders depending on
        the reason they were skipped.
    .INPUTS
        LibraryName
            The name of the library.
            To move the files that were skipped during the creation of this library.
        TargetDir
            The Directory to create the folder containing the skipped files in. 
#>

$SourceDir = "$HOME\AppData\Roaming\HSort\SkippedObjects"

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
    $SkippedXML = Import-Clixml -Path "$SourceDir\Skipped $LibraryName.xml"
}
catch {
    Write-Information "$SourceDir\Skipped $LibraryName.xml not found. Exiting"
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
    "Base"         = "$TargetDir\Skipped $LibraryName";
    "NoMatch"      = "$TargetDir\Skipped $LibraryName\NoMatch";
    "Duplicates"   = "$TargetDir\Skipped $LibraryName\Duplicates";
    "HashMismatch" = "$TargetDir\Skipped $LibraryName\HashMismatch";
    "Other"      = "$TargetDir\Skipped $LibraryName\Other"
}


foreach ($Path in $PathsSkipped.Keys) {
    if (!(Test-Path $PathsSkipped.$Path)) {
        $null = New-Item -ItemType "directory" -Path $PathsSkipped.$Path
    }
}

$DuplicateCounter = 0
$MovedObjects = [List[object]]::new()

Show-Information -InformationArray ("Please wait. Moving objects...`n")

foreach($Reason in $SkippedXML.Keys){

    foreach($Object in $SkippedXML.$Reason.Keys){

        $Source = $SkippedXML.$Reason.$Object.Path
        $Name = $SkippedXML.$Reason.$Object.ObjectName
        $ParentDir = $SkippedXML.$Reason.$Object.ObjectParent

        Write-Information -MessageData "[Moving] $Name" -InformationAction Continue

        # Since an object can have more than one duplicate,
        # we have to make sure that every object moved to the duplicates
        # folder has a unique name
        if ($SkippedXML.$Reason.$Object.Reason -eq "Duplicate") {
            $DuplicateCounter += 1
        }

        if($SkippedXML.$Reason.$Object.Extension -eq "Folder") {

            switch(($SkippedXML.$Reason.$Object.Reason)){

                "NoMatch" {
                            try{$null = robocopy $Source "$($PathsSkipped.NoMatch)\$Name" /E /DCOPY:DAT /move ; $MovedObjects.Add()}
                            catch {Write-Information "Copy error" -InformationAction Continue};
                            break }
                

                "Duplicate" {
                            try { $null = robocopy $Source "$($PathsSkipped.Duplicates)\$Name" /E /DCOPY:DAT /Move;
                            Rename-Item -LiteralPath "$($PathsSkipped.Duplicates)\$Name" -NewName "ID $DuplicateCounter $Name" }
                            catch {Write-Information "Copy error" -InformationAction Continue};
                            break }

                "HashMismatch" {
                            try{$null = robocopy $Source "$($PathsSkipped.HashMismatch)\$Name" /E /DCOPY:DAT /move}
                            catch {Write-Information "Copy error" -InformationAction Continue};
                            break }
                
                # Only move Junk-Objects from this type of folder, but don't move folder itself.
                "UnsupportedOrMixedContent" {break}

                Default { 
                            try{$null = robocopy $Source "$($PathsSkipped.Other)\$Name" /E /DCOPY:DAT /move}
                            catch {Write-Information "Copy error" -InformationAction Continue};
                            break }
            }
        }
        else{

            switch(($SkippedXML.$Reason.$Object.Reason)){

                "NoMatch" {
                            try{ $null = robocopy $ParentDir $PathsSkipped.NoMatch $Name /move}
                            catch { Write-Information "Copy error" -InformationAction Continue};
                            break }

                # LiteralPath is neccessary
                # "$Name ID:$DuplicateCounter" doesn't work
                # $Name has to be at the end, since we need the extension from $Name
                "Duplicate" {
                            try{$null = robocopy $ParentDir $PathsSkipped.Duplicates $Name /move;
                            Rename-Item -LiteralPath "$($PathsSkipped.Duplicates)\$Name" -NewName "ID $DuplicateCounter $Name" }
                            catch {Write-Information "Copy error" -InformationAction Continue};
                            break }


                "HashMismatch" {
                            try{$null = robocopy $ParentDir $PathsSkipped.HashMismatch $Name /move}
                            catch {Write-Information "Copy error"} -InformationAction Continue;
                            break }

                Default {
                            try{$null = robocopy $ParentDir $PathsSkipped.Other $Name /move}
                            catch {Write-Information "Copy error"} -InformationAction Continue;
                            break }

            }
        }
    }
}
Write-Information -MessageData "`nFinished moving objects."

