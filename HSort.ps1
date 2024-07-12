
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

$global:ScriptVersion = "V0.1"
$global:CallDirectory = Get-Location # Accessed in ObjectData.psm1 for Tags.txt
$global:Timestamp = Get-Date -Format FileDateTime

# To import modules successfully, Script has to be executed from .\HSort !
Import-Module -Name "$CallDirectory\Modules\GetObjects\GetObjects.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateComicInfoXML\CreateComicInfoXML.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateLogFiles\CreateLogFiles.psm1"
Import-Module -Name "$CallDirectory\Modules\InitializeScript\InitializeScript.psm1"
Import-Module -Name "$CallDirectory\Modules\DiskSpace\DiskSpace.psm1"
Import-Module -Name "$CallDirectory\Modules\ObjectData\ObjectData.psm1" 

Import-Module -Name "$CallDirectory\Modules\ExcludedObjectsHandler\ExcludedObjectsHandler.psm1" 

Import-Module -Name "$CallDirectory\Modules\CheckEnvironment\CheckEnvironment.psm1" 

#Clear-Host


Class HsCounter {

    [int]$AllObjects

    [int]$ToSort

    [int]$Anthologies
    [int]$Artists
    [int]$Manga
    [int]$Conventions
    [int]$Doujinshi

    [int]$Duplicates
    [int]$NoMatch
    [int]$EFC # Excluded From Copying
    [int]$Variants

    [int]$Collections
    [int]$ToCopy
    [int]$CopyGood
    [int]$CopyErrors
    [int]$Excluded

    Counter() {

        $this.AllObjects = 0
        $this.ToSort = 0
        $this.Anthologies = 0
        $this.Artists = 0
        $this.Manga = 0
        $this.Conventions = 0
        $this.Doujinshi = 0
        $this.EFC = 0
        $this.Variants = 0
        $this.Duplicates = 0
        $this.NoMatch = 0
        $this.ToCopy = 0
        $this.CopyErrors = 0
        $this.Excluded = 0
        $this.CopyGood = 0
        $this.Collections = 0
    }

    [void]AddTitle([string]$PublishingType) {

        if ($PublishingType -eq "Anthology") {
            $this.Anthologies += 1
        }
        elseif ($PublishingType -eq "Manga") {
            $this.Manga += 1
        }
        elseif ($PublishingType -eq "Doujinshi") {
            $this.Doujinshi += 1
        }        
    }

    [void]AddSet([string]$PublishingType) {

        if ($PublishingType -eq "Anthology") {
            $this.Anthologies += 1
        }
        elseif ($PublishingType -eq "Manga") {
            $this.Manga += 1
            $this.Artists += 1
        }
        elseif ($PublishingType -eq "Doujinshi") {
            $this.Conventions += 1
            $this.Doujinshi += 1
        }   
    }

    [void]AddDuplicate(){
        $this.Duplicates += 1
        $this.EFC += 1
    }

    [void]AddNoMatch() {
        $this.EFC += 1
        $this.NoMatch += 1
    }

    [void]AddEFC(){
        $this.EFC += 1
    }

    [void]ComputeToCopy(){
        $this.ToCopy = ($this.AllObjects - $this.EFC)
    }

    [void]AddCopyGood(){
        $this.CopyGood += 1
    }

    [void]AddCopyError(){
        $this.CopyErrors += 1
    }

    [void]ComputeExcluded(){
        $this.Excluded = $this.EFC + $this.CopyErrors
    }

    [void]ComputeCollections(){
        # + 1 for anthologies
        $this.Collections = $this.Artists + $this.Conventions + 1
    }

    [void]AddVariant(){
        $this.Variants += 1
    }
}


#region


function Move-ScriptOutput {
    <#
    .SYNOPSIS
    For debugging only
    .DESCRIPTION
    Moves program-files from Roaming\HSort to Desktop\HSort-ProgramFiles and deletes HSort from Roaming.
    #>

    Param(
        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [string]$PSVersion, 

        [switch]$Delete
    )

    $CopyName = "$PSVersion $LibraryName $Timestamp"

    $TargetDir = "F:\ExportedProgramData"

    $null = New-Item -Path $TargetDir -Name $CopyName -ItemType "directory"
    $null = robocopy "$HOME\AppData\Roaming\HSort" "$TargetDir\$CopyName" /E /DCOPY:DAT

    if ($Delete) {
        Remove-Item -LiteralPath "$HOME\AppData\Roaming\HSort" -Recurse -Force
    }
}


function Add-Excluded {
    <#
    .DESCRIPTION
    If an object 

    - either doesn't match the E-Hentai naming convention in sorting
    - or failed to copy
     
    a hashtable of its properties (see below) is added to [ExcludedObjects].
    #>
    Param(

        [Parameter(Mandatory)]
        [hashtable]$ExcludedObjects,

        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Reason,

        [string]$Extension,

        [switch]$CopyError
    )

    if($CopyError){

        $ParentPath = $Object.ObjectParent
        
        $RelPath = Get-RelativePath -Path ($Object.ObjectSource) -StartDir $SrcDirName
        $RelParentPath = Get-RelativePath -Path $ParentPath -StartDir $SrcDirName

        $ExcludedObjectProperties = @{
            Path               = $Object.ObjectSource;
            RelativePath       = $RelPath;
            ParentPath         = $ParentPath;
            RelativeParentPath = $RelParentPath;
            ObjectName         = $Object.ObjectName;
            Reason             = $Reason;
            Extension          = $Object.Extension
        }

        $null = $ExcludedObjects.add("$($Object.ObjectSource)", $ExcludedObjectProperties)
    }
    else{

        $ParentPath = Split-Path -Parent $Object.FullName

        $RelPath = Get-RelativePath -Path ($Object.FullName) -StartDir $SrcDirName
        $RelParentPath = Get-RelativePath -Path $ParentPath -StartDir $SrcDirName
    
        $ExcludedObjectProperties = @{
            Path               = $Object.FullName;
            RelativePath       = $RelPath;
            ParentPath         = $ParentPath;
            RelativeParentPath = $RelParentPath;
            ObjectName         = $Object.Name;
            Reason             = $Reason;
            Extension          = $Extension
        }
    
        $null = $ExcludedObjects.add("$($Object.FullName)", $ExcludedObjectProperties)
    }

}


function Format-ObjectName {
    <#
    .SYNOPSIS
    Formats a string
    .DESCRIPTION
    Formats a string by removing empty parenthesis/brackets.

    The RegEx in SORTING result in false matches when
    a string contains empty parenthesis/brackets.
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$ObjectNameNEX
    )

    $Normalized = ((($ObjectNameNEX -replace '\( *\)', '') -replace '\{ *\}', '') -replace '\[ *\]', '')

    return $Normalized
}

function ConvertTo-SanitizedNameArray {
    <# 
    .SYNOPSIS
    Sanitizes the elements of a string-array.
    .DESCRIPTION
    Sanitizes all elements of a string array
        - Remove extraneous spaces
        - Replace brackets
        - Remove double extensions
        - Turn meta into a single string of the form "aaa,bbb,ccc,..."
    #>
    Param(
        [Parameter(Mandatory)]
        [array]$NameArray
    )

    for ($i = 1; $i -le ($NameArray.length - 1); $i += 1) {

        if ($NameArray[$i] -ne "") {

            # Remove extraneous spaces
            $Token = ($NameArray[$i] -replace ' {2,}', ' ')
            # Remove framing spaces
            $Token = $Token.trim()
            # Replace brackets
            $Token = (($Token -replace '\{|\[', '(') -replace '\}|\]', ')')

            # If token is Artist.
            if ($i -eq 2) {
                $Token =  ($Token -replace '\(|\)|,|\.', '-')
                $Token = $Token.ToUpper()
            }
            elseif($i -eq 3){

                # 17/05/2024 
                # A folder cannot end with a '.'
                # When you try to create a folder with [New-Item -Name "xxxx."]
                # a folder is created but '.' is silently removed.
                # This makes the folder inaccessible since TargetID and the
                # actual folder name no longer match.

                $Token =  $Token.Trim('.')
                
            }
            # If token is Meta.
            if ($i -eq 4) {


                # Remove false/double extension, example: "File.zip.cbz"
                # Keep in mind that NameArray was build from ObjectNameNEX.
                # So "File.zip.cbz" would arrive here as "File.zip" and ".zip" is removed.
                $Token = $Token.trim('.zip')
                $Token = $Token.trim('.rar')
                $Token = $Token.trim('.cbz')
                $Token = $Token.trim('.cbr')

                ## Turn meta into a single string of the form "aaa,bbb,ccc,..."
                
                # Initialize string to a known state
                $Token = (($Token -replace '\+', '') -replace ',', '')
                
                $Token = (($Token -replace '\{|\[|\(', '+') -replace '\}|\]|\)', '+')
                $Token = ($Token -replace '\+ \+', ',')
                $Token = $Token.trim('+')

            }

            $NameArray[$i] = $Token
        }
        else {
            $NameArray[$i] = ""
        }
    }
}

function Add-TitleToLibrary {
    <# 
    .SYNOPSIS
    Creates an entry in UserLibrary-hashtable.
    .DESCRIPTION
    Creates an entry for a matched object in UserLibrary if it's not a duplicate.
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserLibrary,

        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$HashedID,

        [switch]$Variant
    )

    $PublishingType = $NameArray[0]
    $Convention     = $NameArray[1]
    $Artist         = $NameArray[2]
    $Title          = $NameArray[3]
    
    if($NameArray[4]){
        $Meta = ($NameArray[4]).Split(",")
    }
    else{
        $Meta = @()
    }

    #===========================================================================
    # Title is NOT a variant
    #===========================================================================

    if(! $Variant){

        if ($PublishingType -eq "Manga") {
    
            $UserLibrary.$Artist[$Title] = @{
    
                "VariantList" = [List[string]]::new();
    
                $HashedID      = @{
                    ObjectSource    = $Object.FullName;
                    ObjectLocation  = "$($PathsLibrary.Artists)\$Artist";
                    FirstDiscovered = $Timestamp;
                    Meta            = $Meta
                }
    
            }
    
            # Add Base Variant
            $null = $UserLibrary.$Artist.$Title.VariantList.add($HashedID)
        }
    
        elseif ($PublishingType -eq "Doujinshi") {
    
            $UserLibrary.$Convention[$Title] = @{
    
                "VariantList" = [List[string]]::new();
    
                $HashedID      = @{
                    ObjectSource    = $Object.FullName;
                    ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention";
                    FirstDiscovered = $Timestamp;
                    Meta            = $Meta
                }
    
            }  
    
            $null = $UserLibrary.$Convention.$Title.VariantList.add($HashedID)
        }
    
        elseif ($PublishingType -eq "Anthology") {
    
            $UserLibrary.Anthology[$Title] = @{
    
                "VariantList" = [List[string]]::new();
    
                $HashedID      = @{
                    ObjectSource    = $Object.FullName;
                    ObjectLocation  = "$($PathsLibrary.Anthologies)\$Artist";
                    FirstDiscovered = $Timestamp;
                    Meta            = $Meta
                }
    
            }
    
            $null = $UserLibrary.Anthology.$Title.VariantList.add($HashedID)
        }
    }

    #===========================================================================
    # Title IS a variant
    #===========================================================================

    else{

        if ($PublishingType -eq "Manga") {

            $UserLibrary.$Artist.$Title[$HashedID] = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Artists)\$Artist";
                FirstDiscovered = $Timestamp;
                Meta            = $Meta
            }

            $null = $UserLibrary.$Artist.$Title.VariantList.add($HashedID)
        }
        elseif ($PublishingType -eq "Doujinshi") {

            $UserLibrary.$Convention.$Title[$HashedID] = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention";
                FirstDiscovered = $Timestamp; # change timestamp format here to yyyy-MM-dd
                Meta            = $Meta
            }
        
            $null = $UserLibrary.$Convention.$Title.VariantList.add($HashedID)
        }
        elseif ($PublishingType -eq "Anthology") {

            $UserLibrary.$Artist.$Title[$HashedID] = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$PathsLibrary.Anthologies\$Artist";
                FirstDiscovered = $Timestamp;
                Meta            = $Meta
            }

            $null = $UserLibrary.$Artist.$Title.VariantList.add($HashedID)
        }
    }
}

function Update-TitleAsVariant {
    <# 
    .SYNOPSIS
    Updates ObjectProperties, ObjectMeta and ObjectSelector hashtables.

    .DESCRIPTION
    If an object is identified as a variant, 
    the ObjectProperties, ObjectMeta and ObjectSelector hashtables
    have to be updated to replace Title with VariantTitle.

    The object title is modified to include the "Variant"-tag and a VariantNumber,
    indicating how many variants of a title exist.

    Modifying the title is neccessary to avoid copying two files of the same name
    to the same folder.
    #>
    Param(
        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [hashtable]$ObjectProperties,

        [Parameter(Mandatory)]
        [hashtable]$ObjectMeta,

        [Parameter(Mandatory)]
        [hashtable]$ObjectSelector
    )

    # Add variant-token to title
    $PublishingType = $NameArray[0]; $Convention = $NameArray[1]; $Artist = $NameArray[2]; $Title = $NameArray[3]

    switch ($PublishingType) {
        # -1 since the base title is added to VariantList as well
        "Manga" { $VariantNumber = ($UserLibrary.$Artist.$Title.VariantList.Count) - 1; Break }

        "Doujinshi" { $VariantNumber = ($UserLibrary.$Convention.$Title.VariantList.Count) - 1; Break }

        "Anthology" { $VariantNumber = ($UserLibrary.Anthology.$Title.VariantList.Count) - 1; Break }
        
    }

    $VariantTitle = "$Title Variant $VariantNumber"

    # Update Manga as Variant
    $ObjectProperties.TargetName = $VariantTitle
    $ObjectProperties.TargetID   = $VariantTitle + $ObjectProperties.NewExtension

    $ObjectMeta.Tags             = "$($ObjectMeta.Tags),Variant"

    $ObjectSelector.TargetName   = $VariantTitle

}

function Remove-FromLibrary{
    <# 
    .SYNOPSIS
    Removes an item from  UserLibrary.

    .DESCRIPTION
    Removes the entry of an object from UserLibrary if 
    the object couldn't be copied successfully.
    #>

    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserLibrary,

        [Parameter(Mandatory)]
        [hashtable]$ObjectSelector
    )
            
    $PublishingType = $ObjectSelector.PublishingType
    $Convention     = $ObjectSelector.Convention
    $Artist         = $ObjectSelector.Artist
    $Title          = $ObjectSelector.Title
    $HashedID       = $ObjectSelector.HashedID

    if($PublishingType -eq "Anthology"){

        $UserLibrary.$Artist.$Title.Remove($HashedID)
    }
    elseif($PublishingType -eq "Manga"){

        $UserLibrary.$Artist.$Title.Remove($HashedID)
    }
    elseif($PublishingType -eq "Doujinshi"){

        $UserLibrary.$Convention.$Title.Remove($HashedID)
    }
}

function Remove-ComicInfo{
    <#
    .SYNOPSIS
     Removes a ComicInfo-file from .\Library\ComicInfoFiles.

    .DESCRIPTION
    If an object fails to copy successfully its ComicInfo-file 
    is removed from .\Library\ComicInfoFiles
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$ObjectSelector
    )

    $TargetName = "$($ObjectSelector.TargetName)"

    Remove-Item -LiteralPath "$($PathsLibrary.ComicInfoFiles)\$TargetName" -Recurse -Force
}

function Export-AsCSV {
    <#
        .NOTES
        Expected formatting:
        Number,("Time,ObjectName")
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$TimeTable,

        [Parameter(Mandatory)]
        [string]$Type
    )

    foreach ($Number in $TimeTable.Keys) {
        "$Number,$($TimeTable.$Number)" | Out-File -FilePath "$($PathsLibrary.Logs)\TimeTable$Type $Timestamp.csv" -Encoding unicode -Append -Force
    }
}

function Get-Hash {
    <# 
    .SYNOPSIS
    Hashes a string.

    .DESCRIPTION
    Computes the SHA1-Hash of a string and returns the hash as a string.
    #>
    Param(
        [Parameter(Mandatory)]
        [String]$String
    )

    $Encoding = [system.Text.Encoding]::UTF8
    $SourceStringBytes = $Encoding.GetBytes($String) 

    # Creates a New SHA1-CryptoProvider instance
    $CryptoProvider = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider 

    # Compute hash
    $HashedSourceString = $CryptoProvider.ComputeHash($SourceStringBytes)
    $HashAsString = [System.Convert]::ToBase64String($HashedSourceString)

    return $HashAsString
}


#endregion



#===========================================================================#
#                                   MAIN                                    #
#===========================================================================#




Write-Paragraph -InformationArray (" ",
" _ _ _     _                      _          _____ _____         _   ",
    "| | | |___| |___ ___ _____ ___   | |_ ___   |  |  |   __|___ ___| |_ ",
    "| | | | -_| |  _| . |     | -_|  |  _| . |  |     |__   | . |  _|  _|",
    "|_____|___|_|___|___|_|_|_|___|  |_| |___|  |__|__|_____|___|_| |_|  `n",
    "=====================================================================",
    " ")


$RuntimeStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'
$SortingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'
$CopyingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'

$RuntimeStopwatch.Start()

$VersionMajor = $PSVersionTable.PSVersion.Major
$VersionMinor = $PSVersionTable.PSVersion.Minor
$PSVersion = "$VersionMajor.$VersionMinor"


#===========================================================================
# Check prerequisits
#===========================================================================


Write-Head("Checking Prerequisits")

$PSVersion_ExitCode = Assert-PSVersion -Major $VersionMajor -Minor $VersionMinor
$7zip = Confirm-SevenZip

if($PSVersion_ExitCode -ne 0){
    exit
}
else{
    if(! $7zip){
        exit
    }
}


#===========================================================================
# Define paths of HSort-Dir in Roaming 
#===========================================================================


# It's important that this is an ordered dictionary.
# Otherwise InitializeScript might try to create 
# files/folders in directories that don't exist yet.

$global:PathsProgram = [ordered]@{

    "Parent"      = "$HOME\AppData\Roaming";
    "Base"        = "$HOME\AppData\Roaming\HSort";

    "LibFiles"        = "$HOME\AppData\Roaming\HSort\LibraryFiles";

    "Settings"    = "$HOME\AppData\Roaming\HSort\Settings";
    "TxtSettings" = "$HOME\AppData\Roaming\HSort\Settings.txt";

    "Tmp"         = "$HOME\AppData\Roaming\HSort\Temp"
}


#===========================================================================
# Initialize script 
#===========================================================================


$Init_ExitCode = Initialize-Script


# Script continues IFF $Init_ExitCode -le 0.
if($Init_ExitCode -le 0){
    $SettingsHt = Import-Clixml -LiteralPath "$($PathsProgram.Settings)\ActiveSettings.xml"
    $DiskSpace_ExitCode = Assert-SufficientDiskSpace -SourceDir $SettingsHt.Source -LibraryParentFolder $SettingsHt.Target -Talkative

    if ($DiskSpace_ExitCode -eq 0) {
        $StartScript_ExitCode = Start-Script

        # Script aborted by user.
        if ($StartScript_ExitCode -eq 1) {
            Restore-Settings -Target $PathsProgram.Settings -Source $PathsProgram.Tmp
            exit
        }
    }
    else {
        exit
    }
}
else{
    exit
}


$global:PathsLibrary = [ordered]@{

    "Parent"         = "$($SettingsHt.Target)";
    "Base"           = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)";

    "Library"        = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)";

    "Source"         = $SettingsHt.Source;

    "Artists"        = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Manga by Artists";
    "Conventions"    = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Doujinshi by Conventions";
    "Anthologies"    = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Anthologies";

    "ComicInfoFiles" = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\ComicInfoFiles";

    "Logs"           = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\Logs"
}

$LibraryName = $SettingsHt.LibraryName
$LibSrc = $SettingsHt.Source;
$global:SrcDirName = (Get-Item $PathsLibrary.Source).Name


#===========================================================================
# Import UserLibrary
#===========================================================================


# === Create new library ===
# If (No Libraries exist) OR (no library of this name exists).
if (($Init_ExitCode -eq 0) -or ($Init_ExitCode -eq -1)) {

    Write-Debug "HSort: Creating new library folder"
    
    foreach ($Path in $PathsLibrary.Keys) {
        if ($Path -ne "Source" -and $Path -ne "Parent") {
            $null = New-Item -ItemType "directory" -Path $PathsLibrary.$Path
        }
    }

    $UserLibrary = @{} # See <Add-TitleToLibrary>

}
# === Update library ===
# If (SettingsHistory.xml contains Library(Name) AND Settings unchanged)
# OR
# If (SettingsHistory.xml contains Library(Name) AND ParentDirectory of Library is unchanged AND source is different)
elseif ( ($Init_ExitCode -eq -2) -or ($Init_ExitCode -eq -3) ) {

    try{ 
        $UserLibrary = Import-Clixml -Path "$($PathsProgram.LibFiles)\$LibraryName\UserLibrary $LibraryName.xml"
    }
    catch{
        Write-Information -MessageData "`nCouldn't import UserLibrary. Exiting...`n" -InformationAction Continue
        exit
    }
}


#===========================================================================
# Get objects to process
#===========================================================================


$DiscoveredObjects = @{}
$DiscoveredObjects = Get-Objects -SourceDir $PathsLibrary.Source -PathsProgram $PathsProgram -LibraryName $LibraryName

$SupportedObjects = [List[object]]::new()
$SupportedObjects = $DiscoveredObjects.SupportedObjects

if (! $DiscoveredObjects) {
    Write-Information -MessageData "Library already up-to-date. Exiting..." -InformationAction Continue
    exit
}

# [InitialDefinition] ExcludedObjects_Session
$ExcludedObjects_Session = @{} 
$ExcludedObjects_Session = $DiscoveredObjects.ExcludedObjects

# Folders that contain at least one file with an unsupported extension.
# 28/06/2024 Currently not in use
$UnsupportedFolderCounter = $DiscoveredObjects.UnsupportedFolderCounter

$UnsupportedFileCounter = $DiscoveredObjects.UnsupportedFileCounter

$HsCounter = [HsCounter]::New()
$HsCounter.ToSort = ($SupportedObjects.Count)
$HsCounter.EFC = ($UnsupportedFolderCounter + $UnsupportedFileCounter)
$HsCounter.AllObjects = ($SupportedObjects.Count + $UnsupportedFolderCounter + $UnsupportedFileCounter)


#===========================================================================
# Sort objects to process
#===========================================================================


$SortingProgress = 0
$SortingStopwatch.Start()

Write-Paragraph -InformationArray ("Sorting objects. Please wait."," ")

# SafeCopyFlag is accessed
# in [ObjectData.psm1] <Get-Properties> for Get-FileHash
if ($SettingsHt.SafeCopy -eq "False") {
    $Global:SafeCopyFlag = 1
}
else {
    $Global:SafeCopyFlag = 0
}

# Central hashtable.
# Used in COPYING to copy/sort found objects (Manga).
# Structure <HashedID> = <ObjectProperties>
$ToCopy = @{}

# Allows to access UserLibrary-Hashtable from to ToCopy, and therefore
# to remove an Object from UserLibrary in COPYING if it fails to copy.
$LibraryLookUp = @{}

# Iterate over all supported objects 
foreach ($Object in $SupportedObjects) {
    
    $NoMatchFlag = 0
    $IsDuplicate = $false

    $ObjectName = $Object.Name # < Not a path. Is [FileName.Extension] for Files or [FolderName] for Folders
    $ObjectPath = $Object.FullName
    $IsFile     = ($Object -is [system.io.fileinfo])

    if ($IsFile) {
        
        $Extension = [System.IO.Path]::GetExtension($ObjectPath)
        #$ObjectNameNEX = $ObjectName.Substring(0, $ObjectName.LastIndexOf('.')) # < ObjectName without extension
        $ObjectNameNEX = $ObjectName.trim($Extension)

        # CreationDate: The date the object was created on the user's system.
        $CreationDate = ([System.IO.File]::GetLastWriteTime($ObjectPath)).ToString("yyyy-MM-dd")
    }
    else {
        $ObjectNameNEX = $ObjectName
        $Extension     = "Folder"

        $FolderFirstElement = Get-ChildItem -LiteralPath $ObjectPath -Force -File | Select-Object -First 1
        $CreationDate = $FolderFirstElement.LastWriteTime.ToString("yyyy-MM-dd")
    }

    $NormalizedObjectName = Format-ObjectName -ObjectNameNEX $ObjectNameNEX

    if($NormalizedObjectName -match "\A\((?<Convention>[^\)]*)\)\s\[(?<Artist>[^\]]*)\](?<Title>[^\[{(]*)(?<Meta>.*)"){

        $NameArray = ("Doujinshi", $Matches.Convention, $Matches.Artist, $Matches.Title, $Matches.Meta)
    }
    elseif($NormalizedObjectName -match "\A\W(Anthology)\W(?<Title>[^\[{(]*)(?<Meta>.*)"){ # Anthologies must be matched before Manga
       
        $NameArray = ("Anthology", "", "Anthology", $Matches.Title, $Matches.Meta)  # For Anthologies [$Artist] is ["Anthology"] !
    }
    elseif($NormalizedObjectName -match "\A\[(?<Artist>[^\]]*)\](?<Title>[^\[{(]*)(?<Meta>.*)"){

        $NameArray = ("Manga", "", $Matches.Artist, $Matches.Title, $Matches.Meta)
    }
    else {
        $NoMatchFlag = 1
        $HsCounter.AddNoMatch()
        Add-Excluded -ExcludedObjects $ExcludedObjects_Session -Object $Object -Reason "NoMatch" -Extension $Extension
    }

    if($NoMatchFlag -eq 0){

        ConvertTo-SanitizedNameArray -NameArray $NameArray

        $PublishingType = $NameArray[0]; $Convention = $NameArray[1]; $Artist = $NameArray[2]; $Title = $NameArray[3]

        $Collection     = Select-Collection -NameArray $NameArray # is $Artist xor $Convention xor "Anthology"
        $Id             = New-Id -NameArray $NameArray
        $HashedID       = Get-Hash -String $Id
        $TargetName     = Get-TargetName -NameArray $NameArray

        [hashtable]$ObjectProperties = Get-Properties -Object $Object -TargetName $TargetName -NameArray $NameArray -Ext $Extension 
        [hashtable]$ObjectMeta       = Get-Meta -NameArray $NameArray -CreationDate $CreationDate
        [hashtable]$ObjectSelector   = New-Selector -NameArray $NameArray -HashedID $HashedID -TargetName $TargetName

        if ( ! $UserLibrary.ContainsKey($Collection) ) {

            $UserLibrary[$Collection] = @{}
            $null = New-Item -Path ($ObjectProperties.ObjectTarget) -ItemType "directory" # Create empty folder for $Artists/$Convention/AnthologY
            
            Add-TitleToLibrary -UserLibrary $UserLibrary -NameArray $NameArray -Object $Object -HashedID $HashedID
            $HsCounter.AddSet($PublishingType)
        }
        elseif ( ! $UserLibrary.$Collection.ContainsKey($Title) ) {

            Add-TitleToLibrary -UserLibrary $UserLibrary -NameArray $NameArray -Object $Object -HashedID $HashedID
            $HsCounter.AddTitle($PublishingType)
        }
        elseif ( $UserLibrary.$Collection.ContainsKey($Title) -and (! $UserLibrary.$Collection.$Title.VariantList.Contains($HashedID)) ) {
            
            Update-TitleAsVariant -NameArray $NameArray -ObjectProperties $ObjectProperties -ObjectMeta $ObjectMeta -ObjectSelector $ObjectSelector
            
            Add-TitleToLibrary -UserLibrary $UserLibrary -NameArray $NameArray -Object $Object -HashedID $HashedID -Variant
            $HsCounter.AddTitle($PublishingType)
            $HsCounter.AddVariant()
        }
        else {
            $IsDuplicate = $true
            $HsCounter.AddDuplicate()

            Add-Excluded -ExcludedObjects $ExcludedObjects_Session -Object $Object -Reason "Duplicate" -Extension $Extension
        }
        
        if($IsDuplicate -eq $false){
            $null = $ToCopy.add($HashedID, $ObjectProperties)
            $LibraryLookUp.Add($HashedID, $ObjectSelector)

            # Create ComicInfo-folder and add ComicInfo.xml 
            $null = New-Item -ItemType "directory" -Path "$($PathsLibrary.ComicInfoFiles)\$($ObjectProperties.TargetName)"
            New-ComicInfoFile -ObjectMeta $ObjectMeta -Path "$($PathsLibrary.ComicInfoFiles)\$($ObjectProperties.TargetName)"
        }
    }
    $SortingProgress += 1
    $SortingCompleted = ($SortingProgress / $HsCounter.ToSort) * 100
    Write-Progress -Id 0 -Activity "Sorting" -Status "$ObjectNameNEX" -PercentComplete $SortingCompleted
}

$TotalSortingTime = ($SortingStopwatch.Elapsed).toString()
$SortingStopwatch.Stop()

$ToCopy | Export-Clixml -LiteralPath "$HOME\Desktop\ToCopy.xml" -Force

Write-Output "Creating folder structure: Finished`n"
Write-Output "Sorting: Finished`n"
Start-Sleep -Seconds 1.0

$HsCounter.ComputeToCopy()


#===========================================================================
# Copy objects
#===========================================================================


Write-Paragraph -InformationArray ("Copying objects. This can take a while.", " ")


# For the progress bar.
# Ignores if the object is actually copied successfully or not.
$CopyProgress = 0

$CopyingStopwatch.Start()

foreach ($HashedID in $ToCopy.Keys) {

    $CopyErrorFlag = 0
    $ErrorType = "Error"
    $7zipExitCode = 0

    # Creating aliases
    $Parent      = $ToCopy.$HashedID.ObjectParent
    $TargetDir   = $ToCopy.$HashedID.ObjectTarget
    $SourceID    = $ToCopy.$HashedID.ObjectName # is Object.Name
    $TargetName  = $ToCopy.$HashedID.TargetName
    $TargetID    = $ToCopy.$HashedID.TargetID

    # Progress bar
    $CopyProgress += 1
    $CopyCompleted = ($CopyProgress / $HsCounter.ToCopy) * 100
    Write-Progress -Id 1 -Activity "Copying" -Status "$SourceID" -PercentComplete $CopyCompleted

    # Define ComicInfo.xml source
    $XML_Folder = "$($PathsLibrary.ComicInfoFiles)\$TargetName"
    $XML_FILE   = "$($PathsLibrary.ComicInfoFiles)\$TargetName\ComicInfo.xml"


    #============================ [Copy archives] ============================


    if ($ToCopy.$HashedID.Extension -ne "Folder") {

        $null = robocopy $Parent $TargetDir $SourceID /R:10 /njh /njs
        $RobocopyExitCode = $LASTEXITCODE

        if ($RobocopyExitCode -eq 1) {

            if($SafeCopyFlag -eq 0){
                $ToCopy.$HashedID.TargetHash =
                (Get-FileHash -LiteralPath "$TargetDir\$SourceID" -Algorithm MD5).hash
            }
            else{
                $ToCopy.$HashedID.TargetHash = 0
            }

            if ($ToCopy.$HashedID.SourceHash -eq $ToCopy.$HashedID.TargetHash) {
                
                try {
                    # Rename copied archive to $TargetID
                    Rename-Item -LiteralPath "$TargetDir\$SourceID" -NewName $TargetID
                }
                catch {
                    "TargetID: $TargetID`nTargetDir: $TargetDir`nSourceID: $SourceID`nException`n$_`n" | Out-File -LiteralPath "$($PathsProgram.Tmp)\ErrorDump.txt" -Encoding unicode -Append -Force
                    Add-Excluded -ExcludedObjects $ExcludedObjects_Session -Object $ToCopy.$HashedID -Reason "RenamingError" -CopyError
                    Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$HashedID
                    Remove-ComicInfo -ObjectSelector $LibraryLookUp.$HashedID

                    # Remove copied archive
                    Remove-Item -LiteralPath "$TargetDir\$SourceID"-Force

                    $HsCounter.AddCopyError()
                    Continue
                }
                
                #============================ [BEGIN: Insert ComicInfo.xml] ============================
                <#
                .NOTES
                Expectable errors:
                1.) 7Zip throws an error - probably because the archive is corrupt -
                    when trying to move ComicInfo.xml into the archive.
                2.) XML doesn't exist.
                
                PS7 doesn't catch SeZipErrors.
                PS5.1 doesn't save the 7z-ExitCode to the $7zipExitCode array.
                This means $7zipExitCode.length is _always_ 0, what requires the code below...
                #>

                if ($VersionMajor -eq 5){
                    try {
                        & $7zip a -bsp0 -bso0 "$TargetDir\$TargetID" $XML_FILE *>&1
                    }
                    catch {
                        $7zipExitCode = 999 # Faking 7zip-ExitCode for PS5.1
                    }
                }
                elseif ($VersionMajor -gt 5) {
                    $7zipExitCode = & $7zip a -bsp0 -bso0 "$TargetDir\$TargetID" $XML_FILE *>&1
                }

                if ($7zipExitCode -ne 0) {
                    $CopyErrorFlag = 1
                    $ErrorType = "XMLinsertionError"

                    Remove-Item -LiteralPath "$TargetDir\$TargetID" -Force # Remove copied archive
                }            
                #============================ [END: Insert ComicInfo.xml] ============================
            }
            else {
                $CopyErrorFlag = 1
                $ErrorType = "HashMismatch"
            }
        }
        else {
            $CopyErrorFlag = 1
            $ErrorType = "RobocopyError: $($RobocopyExitCode)"
        }
    }


    #============================ Copy folders ============================


    elseif ($ToCopy.$HashedID.Extension -eq "Folder") {
        
        try {
            # Create empty TargetName-Folder
            $null = New-Item -Path $TargetDir -Name $TargetName -ItemType "directory"
        }
        catch {
            Add-Excluded -ExcludedObjects $ExcludedObjects_Session -Object $ToCopy.$HashedID -Reason "ErrorCreatingFolder" -CopyError
            Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$HashedID
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$HashedID

            $HsCounter.AddCopyError()
            Continue
        }

        $null = robocopy ($ToCopy.$HashedID.ObjectSource) "$TargetDir\$TargetName"  /R:10 /njh /njs
        $RobocopyExitCode = $LASTEXITCODE

        if ($RobocopyExitCode -eq 1) {

            if($SafeCopyFlag -eq 0){
                # Get the folder size (instead of a hash) _before_ inserting the ComicInfo.XML file.
                # Update TargetHash of ObjectProperties (accessed with: $ToCopy.$HashedID.TargetHash)
                $ToCopy.$HashedID.TargetHash =
                    ((Get-ChildItem -LiteralPath "$TargetDir\$TargetName") | Measure-Object -Sum Length).sum
            }
            else{
                $ToCopy.$HashedID.TargetHash = 0
            }

            if ($ToCopy.$HashedID.SourceHash -eq $ToCopy.$HashedID.TargetHash) {

                #============================ [Insert ComicInfo.xml] ============================

                # Move XML into TargetName-folder
                $null = robocopy $XML_Folder "$TargetDir\$TargetName" "ComicInfo.xml" /R:10 /njh /njs

                #============================ [Create Archive] ============================

                # Try to create an Archive with the copied folder's content
                # and delete the (unzipped) folder if successful.
                
                $7zipTarget = "$TargetDir\$TargetID"
                $7zipSource = "$TargetDir\$TargetName\*"
                
                if ($VersionMajor -eq 5) {
                    try {
                        & $7zip a -mx3 -bsp0 -bso0 $7zipTarget $7zipSource *>&1
                    }
                    catch {
                        $7zipExitCode = 999 # Faking 7zip-ExitCode for PS5.1
                    }
                }
                elseif ($VersionMajor -gt 5) {
                    $7zipExitCode = & $7zip a -mx3 -bsp0 -bso0 $7zipTarget $7zipSource *>&1
                }

                if ($7zipExitCode -ne 0) {
                    $CopyErrorFlag = 1
                    $ErrorType = "ErrorCreatingArchive"
                }

            }
            else {
                $CopyErrorFlag = 1
                $ErrorType = "HashMismatch"
            }
            # Remove the unzipped folder - always.
            Remove-Item -LiteralPath "$TargetDir\$TargetName" -Recurse -Force
        }
        else {
            $CopyErrorFlag = 1
            $ErrorType = "RobocopyError: $($RobocopyExitCode)"
        }
    }

    if($CopyErrorFlag -eq 1){

        Add-Excluded -ExcludedObjects $ExcludedObjects_Session -Object $ToCopy.$HashedID -Reason $ErrorType -CopyError
        Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$HashedID
        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$HashedID
        $HsCounter.AddCopyError()        

    }
    elseif($CopyErrorFlag -eq 0){
        $HsCounter.AddCopyGood()
    }
}

$CopyingStopwatch.Stop()
$RuntimeStopwatch.Stop()
$TotalCopyingTime = ($CopyingStopwatch.Elapsed).toString()
$TotalRuntime = ($RuntimeStopwatch.Elapsed).toString()

Write-Information -MessageData "Copying: Finished`n" -InformationAction Continue

#===========================================================================
# Create logs and export program files
#===========================================================================


Write-Information -MessageData "Finalizing. Please wait.`n" -InformationAction Continue

# Write Log-Files
Update-CopiedLog -ObjectsToCopy $ToCopy -TargetDir $PathsLibrary.Logs
Update-ExcludedLog -ExcludedObjects $ExcludedObjects_Session -TargetDir $PathsLibrary.Logs

# Export/Update UserLibrary
$UserLibrary | Export-Clixml -Path "$($PathsProgram.LibFiles)\$LibraryName\UserLibrary $LibraryName.xml" -Force

Invoke-ExcludedObjectsHandler -ExcludedObjects_Session $ExcludedObjects_Session -LibraryFiles $PathsProgram.LibFiles -LibraryName $LibraryName -LibSrc $LibSrc -SrcDirTree $DiscoveredObjects.SrcDirTree


$HsCounter.ComputeExcluded()
$HsCounter.ComputeCollections()

Write-Head("Script finished")

$Summary = @"

SUMMARY [$Timestamp]
=================================================

[ ScriptVersion: $ScriptVersion ]
[ UsedPowershellVersion: $PSVersion ]
[ Settings.txt location: $($PathsProgram.TxtSettings) ]

[ Source Folder: $($PathsLibrary.Source) ]

> TotalRuntime: $TotalRuntime

> TotalSortingTime: $TotalSortingTime

> TotalCopyingTime: $TotalCopyingTime

# DiscoveredObjects: $($HsCounter.AllObjects)
=================================================

> Successfully Copied Objects: $($HsCounter.CopyGood)

> CopyCompleteness (should be 0): $($HsCounter.ToCopy - $CopyProgress )

> Total number of Collections: $($HsCounter.Collections)

> Found $($HsCounter.Manga) Manga from $($HsCounter.Artists) Artists

> Found $($HsCounter.Doujinshi) Doujinshi from $($HsCounter.Conventions) Conventions

> Found $($HsCounter.Anthologies) Anthologies

> Found $($HsCounter.Variants) Variants


# Excluded Objects: $($HsCounter.Excluded)
=================================================

> Duplicates: $($HsCounter.Duplicates)

> UnmatchedObjects: $($HsCounter.NoMatch)

> Copy Errors: $($HsCounter.CopyErrors)

> Single files with
  wrong extension (.mov,.mp4,.pdf,...): $UnsupportedFileCounter

> Leaf-Folders that contain at least
  one file that is NOT a (.jpg,.jpeg,.png,.txt): $UnsupportedFolderCounter

=================================================
=================================================

"@

$Summary | Out-Host
$Summary | Out-File -FilePath "$($PathsLibrary.Logs)\ProgramLog.txt" -Encoding unicode -Append -Force

# For debugging only
# Move-ScriptOutput -LibraryName $SettingsHt.LibraryName -PSVersion $PSVersion -Delete

$SortingProgress = 0
$CopyProgress = 0

