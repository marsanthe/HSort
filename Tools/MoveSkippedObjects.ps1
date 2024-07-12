
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

$Directory = Split-Path -Parent(Get-Location) # Parent( (Get-Location == Tools) ) == HSort 

# To import module, Script has to be executed from .\HSort\Tools !
Import-Module -Name "$Directory\Modules\InitializeScript\InitializeScript.psm1"

#===========================================================================
# Functions
#===========================================================================

function Assert-ValidString{

    Param(
        [Parameter(Mandatory,Position = 0)]
        $string
    )

    [bool]$Value

    if ($string -match "^[\w\+\-]+$") {
        $Value = $true
    }
    else {
        $Value = $false
    }

    # Write-Host "Value: $Value"

    return $Value
}
function Read-SrcDirTree {

    Param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    
    $AllSrcDirTrees = @{}

    # Structure of the [SrcDirTree LibraryName.txt] file
    #
    # [SectionOne_Name]
    # key = value
    # .
    # .
    # .
    # key = value
    #
    # [SectionTwo_Name]
    # key = value
    # .
    # .
    # .
    # key = value
    #
    # ...
    
    # [SrcDirTree] must be defined here _initially_
    # to allow us to add the first SECTION to [AllSrcDirTrees]
    $SrcDirTree = [ordered]@{}
    $SectionCounter = 0

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        
        if ($line -ne "" -and -not $line.StartsWith("#")) {

            if ($line -match "^\[(?<Source>.+)\]") {

                $SectionCounter++

                if ($SrcDirTree.Count -gt 0) {
                    if(-not $AllSrcDirTrees.ContainsKey($Source)){
                        $null = $AllSrcDirTrees.Add($Source, $SrcDirTree)
                    }
                    # If a section for a source already exists,
                    # replace the old directory-tree with the new one.
                    #
                    # SrcDirTree.txt shall therefore be ordered ascending by date (Out-File ... -Append)
                    else{
                        $AllSrcDirTrees.$Source = $SrcDirTree
                    }
                    $SrcDirTree = [ordered]@{}
                }
                # Keep this order.
                # This must not be put into an else{} statement.
                $Source = $Matches.Source

            }
            # Match the parent node
            # Create new [CurrentPnode] key in [SrcDirTree]
            # that points to an empty list.
            elseif ($line -match "(?<Argument>^Pnode) = (?<PnodePath>.+)"){

                $CurrentPnode = $Matches.PnodePath
                $SrcDirTree.Add($CurrentPnode, [List[String]]::new())
            }
            # Match the child nodes
            # Add child nodes to the list of [CurrentPnode] 
            elseif ($line -match "(?<Argument>^Cnode) = (?<CnodePath>.+)"){
                $SrcDirTree.$CurrentPnode.Add($Matches.CnodePath)
            }
        }
    }

    # This adds the last Section to [$AllSrcDirTrees]
    if ($SrcDirTree.Count -gt 0) {
        if(-not $AllSrcDirTrees.ContainsKey($Source)){
            $null = $AllSrcDirTrees.Add($Source, $SrcDirTree)
        }
        else{ # Update $AllSrcDirTrees - SrcDirTree.txt shall be ordered ascending by date (Out-File ... -Append)
            $AllSrcDirTrees.$Source = $SrcDirTree
        }
        $SrcDirTree = [ordered]@{}
    }

    return $AllSrcDirTrees
}

function Select-FromHashtable{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $SelectionHashtable = [ordered]@{}

    Write-Information -MessageData " " -InformationAction Continue

    $i = 0
    foreach ($Key in $Hashtable.Keys) {

        $SelectionHashtable.Add("$i", $Key)

        Write-Information -MessageData "[$i] $Key" -InformationAction Continue
        $i += 1
    }

    Write-Information -MessageData " " -InformationAction Continue


    while ($true) {

        $UserInput = Read-Host "Enter a number to select a $Name. Enter [n] to exit."

        Write-Information -MessageData " " -InformationAction Continue

        if ($UserInput -match "^\d+$") {
            if ($SelectionHashtable.Contains($UserInput)) {
                $KeyName = $SelectionHashtable.$UserInput
                break
            }
            else {
                Write-Information -MessageData "Please enter a number corresponding to a $Name. Enter [n] to exit." -InformationAction Continue
            }
        }
        else {
            if ($UserInput -eq "n") {
                Write-Information -MessageData "Exiting..." -InformationAction Continue
                exit
            }
            else {
                Write-Information -MessageData "Please enter a number corresponding to a $Name. Enter [n] to exit." -InformationAction Continue
            }
        }
    }

    Write-Information -MessageData "You selected $KeyName" -InformationAction Continue

    return $KeyName
}

function Move-Folder {

    Param(

        [Parameter(Mandatory)]
        [string]$ObjSrcPath,

        [Parameter(Mandatory)]
        [string]$ExcludedFolderPath,

        [Parameter(Mandatory)]
        [string]$RelPath,

        [switch]$Pretend
    )

    $ObjTargetPath = "$ExcludedFolderPath\$RelPath"

    if(-not $Pretend){

        $null = robocopy $ObjSrcPath $ObjTargetPath  /move /njh /njs
        $RobocopyExitCode = $LASTEXITCODE
        if ($RobocopyExitCode -ne 1) {
            Write-Information -MessageData "Container doesn't exist." -InformationAction Continue
            $ErrorCounter += 1
        }
    }
    else{
        Write-Information -MessageData "`nMoving Folder..." -InformationAction Continue
        Write-Information -MessageData "[From] $ObjSrcPath" -InformationAction Continue
        Write-Information -MessageData "[To]   $ObjTargetPath" -InformationAction Continue
    }

}

function Move-File {

    Param(

        [Parameter(Mandatory)]
        [hashtable]$ExcludedObject,

        [Parameter(Mandatory)]
        [string]$ExcludedFolderPath,

        [switch]$Pretend
    )

    $ObjTargetPath = "$ExcludedFolderPath\$($ExcludedObject.RelativeParentPath)"
    
    $ObjTargetPath = $ObjTargetPath.trim("\")
    
    if(-not $Pretend){

        $null = robocopy $ExcludedObject.ParentPath $ObjTargetPath $ExcludedObject.ObjectName /mov /njh /njs
        $RobocopyExitCode = $LASTEXITCODE
        if ($RobocopyExitCode -ne 1) {
            Write-Information -MessageData "[Error moving container] $($ExcludedObject.ObjectName)." -InformationAction Continue
            $ErrorCounter += 1
        }
    }
    else{
        Write-Information -MessageData "`nMoving File..." -InformationAction Continue
        Write-Information -MessageData "[Name] $($ExcludedObject.ObjectName)" -InformationAction Continue
        Write-Information -MessageData "[From] $($ExcludedObject.ParentPath)" -InformationAction Continue
        Write-Information -MessageData "[To]   $ObjTargetPath" -InformationAction Continue
    }
}

#===========================================================================
# Begin Main
#===========================================================================

Write-Paragraph -InformationArray (" ",
"INFORMATION",
"==========================",
"This script is meant to be run after a Library was created from a source folder.",
"If objects from this source folder were excluded, this script will MOVE them to a folder named:",
"ExcludedObjects from [YourLibraryName]",

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


#===========================================================================
# User Input: Where to create [Excluded from Library] folder
#===========================================================================

while ($True) {
    $ExcludedFolderLocation = Read-Host "`nWhere would you like to create the [Excluded from $LibraryName] folder?`n"

    if (!(Test-Path $ExcludedFolderLocation)) {
        $ExcludedFolderLocation = Read-Host "This directory doesn't exist. Please enter a different path."
    }
    elseif (Test-Path "$ExcludedFolderLocation\ExcludedObjects from $LibraryName") {
        $ExcludedFolderLocation = Read-Host "This directory already contains a folder named: [Excluded $LibraryName].`nPlease enter a different path."
    }
    else {
        $ExcludedFolderName = "Excluded from $LibraryName"
        $ExcludedFolderPath = "$ExcludedFolderLocation\$ExcludedFolderName"
        $null = New-Item -Path $ExcludedFolderLocation -ItemType Directory -Name $ExcludedFolderName #-WhatIf
        break
    }
}


#===========================================================================
# User Input: Select target library
#===========================================================================

<# 
The excluded objects of which library do you want to move?
#>
if ((Test-Path "$HOME\AppData\Roaming\HSort\Settings\SettingsHistory.xml")){
    
    $SettingsHistory = Import-Clixml -LiteralPath "$HOME\AppData\Roaming\HSort\Settings\SettingsHistory.xml"
}
else{
    
    Write-Information -MessageData "SettingsHistory not found.`nExiting..."
    Exit
}

Write-Head("Select Target Library")
Write-Paragraph -InformationArray ("The excluded objects of which library do you want to move?,")
$LibraryName = Select-FromHashtable -Hashtable $SettingsHistory -Name "Library"





#===========================================================================
# Create Hashtable from [AllSrcDirTrees.txt]
#===========================================================================

#$SrcDirTree = Read-SrcDirTree -Path "$HOME\AppData\Roaming\HSort\LibraryFiles\$LibraryName\SrcDirTree $LibraryName.txt"
$AllSrcDirTrees = Read-SrcDirTree -Path "$HOME\AppData\Roaming\HSort\LibraryFiles\$LibraryName\SrcDirTree $LibraryName.txt"

#===========================================================================
# User Input: Select build folder of target library
# (In case the library was build from multiple folders.)
#===========================================================================

Write-Head("Select source folder")

Write-Paragraph -InformationArray ("If your library was build from multiple sources,",
"you can select the right source folder now.")

$SrcPath = Select-FromHashtable -Hashtable $AllSrcDirTrees -Name "Source folder"

$SrcDirTree = [ordered]@{}
$SrcDirTree = $AllSrcDirTrees.$SrcPath

Write-Host " "

foreach ($Pnode in $SrcDirTree.Keys){
    Write-Debug "Pnode: $Pnode"
    foreach($Cnode in $SrcDirTree.$Pnode){ # this is a list
        Write-Debug "Cnode: $Cnode"
    }
}



#===========================================================================
# User Input: Final confirmation before moving objects
#===========================================================================

while($true){
    $UserResponse = Read-Host "Press [y] to continue to move the Excluded files from $SrcPath.`nPress [n] to exit."
    if($UserResponse -eq "y"){
        break
    }
    elseif($UserResponse -eq "n"){
        Write-Information -MessageData "Moving Excluded objects cancelled.`nExiting..." -InformationAction Continue
        Exit
    }
    else{
        Write-Information -MessageData "Please enter [y] to continue or [n] to exit." -InformationAction Continue
    }
}

#===========================================================================
# Create pruned mirror of SrcDir 
#===========================================================================

# Pnode and Cnode are always relative paths !!!
foreach ($Pnode in $SrcDirTree.Keys) {

    # The full path with respect to ExcludedFolderPath
    $PnodeFullPath = "$ExcludedFolderPath\$Pnode"
    
    Write-Debug "Parent: $Pnode"

    if ( -not (Test-Path -LiteralPath "$ExcludedFolderPath\$Pnode")) {

        New-Item -Path $ExcludedFolderPath -ItemType Directory -Name $Pnode #-WhatIf
    }

    # List of direct child-nodes
    $Cnodes = $SrcDirTree.$Pnode

    # Check if root node is not a dead end
    if ($Cnodes[0] -ne $Pnode) {
        # ... create new folder for every Cnode in Pnode
        foreach ($Cnode in $Cnodes) {

            Write-Debug "Child: $Cnode"

            New-Item -Path $PnodeFullPath -ItemType Directory -Name $Cnode #-WhatIf

        }
    }
}

#===========================================================================
# Move Excluded objects
#===========================================================================


if ((Test-Path "$HOME\AppData\Roaming\HSort\LibraryFiles\$LibraryName\Excluded $LibraryName.xml")) {
    
    $ExcludedXML = Import-Clixml -LiteralPath "$HOME\AppData\Roaming\HSort\LibraryFiles\$LibraryName\Excluded $LibraryName.xml"

}
else {

    Write-Information -MessageData "Excluded xml not found.`nExiting..."
    Exit
}


Write-Head("Source folders")

$LibrarySourceName = Select-FromHashtable -Hashtable $ExcludedXML -Name "Source folder"

$Excluded = $ExcludedXML.$LibrarySourceName.Excluded

foreach($Object in $Excluded.Keys){
    if($Excluded.$Object.Extension -eq "Folder"){
        Move-Folder -ObjSrcPath $Excluded.$Object.Path -ExcludedFolderPath $ExcludedFolderPath -RelPath $Excluded.$Object.RelativePath #-Pretend
    }
    else{
        Move-File -ExcludedObject ($Excluded.$Object) -ExcludedFolderPath $ExcludedFolderPath #-Pretend
    }
}

#===========================================================================
# Restore Test Data - for Script Testing
#===========================================================================

while($true){
    $response = Read-Host "Restore test data?"
    if($response -eq "y"){
        $Name = (Get-Item -LiteralPath $SrcPath).BaseName
        Remove-Item -LiteralPath $SrcPath -Recurse -Force
        New-Item -Path "F:\TestsExt" -ItemType "Directory" -Name "$Name"
        $null = robocopy "F:\TestsExt\BackUp\$Name" "F:\TestsExt\$Name" /e /njh /njs
        break
    }
    elseif($Response -eq "n"){

        break
    }
    else{
        Write-Information -MessageData "Please enter [y] or [n]"
    }
}