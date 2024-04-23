using namespace System.Collections.Generic

<# 
.NOTES
    Accessed in ObjectData.psm1 for Tags.txt
#>
$global:CallDirectory = Get-Location

Import-Module -Name "$CallDirectory\Modules\GetObjects\GetObjects.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateComicInfoXML\CreateComicInfoXML.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateLogFiles\CreateLogFiles.psm1"
Import-Module -Name "$CallDirectory\Modules\InitializeScript\InitializeScript.psm1"
Import-Module -Name "$CallDirectory\Modules\DiskSpace\DiskSpace.psm1"
Import-Module -Name "$CallDirectory\Modules\ObjectData\ObjectData.psm1" 

$ErrorActionPreference = "Stop"

Clear-Host

$global:Timestamp = Get-Date -Format "hh_mm_ss_dd_MM_yyyy"

### Begin: Classes ###

<# 
Class ObjectCounter {

    [int]$Manga
    [int]$Doujinshi
    [int]$Convention
    [int]$Anthology

    [int]$Title

    EventCounter(){

        $this.Manga = 0
        $this.Doujinshi = 0
        $this.Convention = 0
        $this.Anthology = 0
    }

}
#>

Class EventCounter {

    <# 
        .SYNOPSIS
        Counting events.
    #>

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

    [int]$ToCopy
    [int]$CopyGood
    [int]$CopyErrors
    [int]$Skipped

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
        $this.Skipped = 0
        $this.CopyGood = 0
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

    [void]ComputeSkipped(){
        $this.Skipped = $this.EFC + $this.CopyErrors
    }

    [void]AddVariant(){
        $this.Variants += 1
    }
}

### End: Classes ###


#region

### Begin: Functions ###


function Copy-ScriptOutput {
    <#
    .DESCRIPTION
        For debugging only.
        Copy HSort folder to HSort-ProgramFiles and delete HSort from Roaming.
    #>

    Param(
        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [string]$PSVersion, 

        [switch]$Delete
    )

    $CopyName = "$PSVersion $LibraryName $Timestamp"
    $TargetParent = "$HOME\Desktop"
    $Target = "$HOME\Desktop\HSortProgramFiles"

    if (!(Test-Path $Target)) {
        $null = New-Item -Path $TargetParent -Name "HSortProgamFiles" -ItemType "directory"
    }

    $null = New-Item -Path $Target -Name $CopyName -ItemType "directory"
    $null = robocopy "$HOME\AppData\Roaming\HSort" "$Target\$CopyName" /E /DCOPY:DAT

    if ($Delete) {
        Remove-Item -LiteralPath "$HOME\AppData\Roaming\HSort" -Recurse -Force
    }
}

function Show-String {

    param(
        [Parameter(Mandatory)]
        # String array
        [string[]]$StringArray
    )

    for ($i = 0; $i -le ($StringArray.length - 1); $i++) {
        Write-Information -MessageData $StringArray[$i] -InformationAction Continue
    }
}

function Add-Skipped {
    <#
    .DESCRIPTION
        Objects that were skipped during SORTING.

        The content of SkippedObjects is formatted and
        written to SkippedObjects.txt in UserLibrary\Logs.
    #>

    Param(

        [Parameter(Mandatory)]
        [hashtable]$SkippedObjects,

        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Reason,

        [Parameter(Mandatory)]
        [string]$Extension
    )

    $SkippedObjectProperties = @{
        ObjectParent = (Split-Path -Parent $Object.FullName);
        Path         = $Object.FullName;
        ObjectName   = $Object.Name;
        Reason       = $Reason;
        Extension    = $Extension
    }

    if ($SkippedObjects.ContainsKey($Reason)) {
        $null = $SkippedObjects.$Reason.add("$($Object.FullName)", $SkippedObjectProperties)
    }
    else {
        $SkippedObjects[$Reason] = @{}
        $null = $SkippedObjects.$Reason.add("$($Object.FullName)", $SkippedObjectProperties)
    }
    

}

function Add-NoCopy {
    <#
    .DESCRIPTION
        Objects that caused errors in COPYING are stored in SkippedObjects.
        The content of SkippedObjects is formatted and
        written to  UserLibrary\Logs\SkippedObjects.txt.
    #>

    Param(
        [Parameter(Mandatory)]
        [hashtable]$SkippedObjects,
    
        [Parameter(Mandatory)]
        [hashtable]$Object,

        [Parameter(Mandatory)]
        [string]$Reason
    )  

    $SkippedObjectProperties = @{
        ObjectParent = $Object.ObjectParent;
        Path         = $Object.ObjectSource;
        ObjectName   = $Object.ObjectName;
        Reason       = $Reason;
        Extension    = $Object.Extension
    }

    if($SkippedObjects.ContainsKey($Reason)){
        $null = $SkippedObjects.$Reason.Add("$($Object.ObjectSource)", $SkippedObjectProperties)
    }
    else{
        $SkippedObjects[$Reason] = @{}
        $null = $SkippedObjects.$Reason.Add("$($Object.ObjectSource)", $SkippedObjectProperties)
    }

}

function Format-ObjectName {
    <#
    .DESCRIPTION
        Remove empty parenthesis/brackets.
        The RegEx patterns used in SORTING result in false matches when
        encountering empty parenthesis/brackets.
    #>

    Param(
        [Parameter(Mandatory)]
        [string]$ObjectNameNEX
    )

    $Normalized = ((($ObjectNameNEX -replace '\( *\)', '') -replace '\{ *\}', '') -replace '\[ *\]', '')

    return $Normalized
}

function Edit-ObjectNameArray {

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray
    )

    for ($i = 1; $i -le ($ObjectNameArray.length - 1); $i += 1) {

        if ($ObjectNameArray[$i] -ne "") {

            # Remove extraneous spaces
            $Token = ($ObjectNameArray[$i] -replace ' {2,}', ' ')
            # Remove framing spaces
            $Token = $Token.trim()
            # Replace brackets
            $Token = (($Token -replace '\{|\[', '(') -replace '\}|\]', ')')

            # If token is Artist.
            if ($i -eq 2) {
                # $Token = Edit-SpecialChars -TokenString $Token
                $Token =  ($Token -replace '\(|\)|,|\.', '-')
                $Token = $Token.ToUpper()
            }

            # If token is Meta.
            if ($i -eq 4) {

                # Remove double extensions,
                # example: "File.zip.cbz"
                $Token = $Token.trim('.zip')
                $Token = $Token.trim('.rar')
                $Token = $Token.trim('.cbz')
                $Token = $Token.trim('.cbr')

                $Token = (($Token -replace '\+', '') -replace ',', '')
                $Token = (($Token -replace '\{|\[|\(', '+') -replace '\}|\]|\)', '+')
                $Token = ($Token -replace '\+ \+', ',')
                $Token = $Token.trim('+')

            }

            $ObjectNameArray[$i] = $Token
        }
        else {
            $ObjectNameArray[$i] = ""
        }
    }
}

function Add-TitleToLibrary {
    <# 
    .NOTES
        The title is only a PART of the object name.
        Rename NewName to something like: ReFoName (re-formatted name)
    #>
    
    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserLibrary,

        [Parameter(Mandatory)]
        [array]$ObjectNameArray,

        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$NewName
    )

    $PublishingType = $ObjectNameArray[0]
    $Convention = $ObjectNameArray[1]
    $Artist = $ObjectNameArray[2]
    $Title = $ObjectNameArray[3]                        

    
    if ($PublishingType -eq "Manga") {

        $UserLibrary.$Artist[$Title] = @{

            "VariantList" = [List[string]]::new();

            $NewName      = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Artists)\$Artist";
                FirstDiscovered = $Timestamp
            }

        }

        # Add Base Variant
        $null = $UserLibrary.$Artist.$Title.VariantList.add($NewName)
    }

    elseif ($PublishingType -eq "Doujinshi") {

        $UserLibrary.$Convention[$Title] = @{

            "VariantList" = [List[string]]::new();

            $NewName      = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention";
                FirstDiscovered = $Timestamp;
            }

        }  

        $null = $UserLibrary.$Convention.$Title.VariantList.add($NewName)
    }

    elseif ($PublishingType -eq "Anthology") {

        $UserLibrary.Anthology[$Title] = @{

            "VariantList" = [List[string]]::new();

            $NewName      = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Anthologies)\$Artist";
                FirstDiscovered = $Timestamp;
            }

        }

        $null = $UserLibrary.Anthology.$Title.VariantList.add($NewName)
    }
}

function New-VariantNameArray {

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [array]$VariantNameArray
    )

    # Read from $NameArray 
    $PublishingType = $NameArray[0]; $Convention = $NameArray[1]; $Artist = $NameArray[2]; $Title = $NameArray[3]

    switch ($PublishingType) {

        "Manga" { $VariantNumber = ($UserLibrary.$Artist.$Title.VariantList.Count) - 1; Break }

        "Doujinshi" { $VariantNumber = ($UserLibrary.$Convention.$Title.VariantList.Count) - 1; Break }

        "Anthology" { $VariantNumber = ($UserLibrary.Anthology.$Title.VariantList.Count) - 1; Break }
        
    }

    $VariantTitle = "$Title Variant $VariantNumber"

    # Copy Values from NameArray to VariantNameArray, but replace title with variant title
    for ($i = 0; $i -le ($NameArray.length - 1); $i ++) {

        # If Title, replace with $VariantTitle
        if ($i -eq 3) {
            $VariantNameArray[$i] = $VariantTitle
        }
        # Else copy from $NameArray
        else {
            $VariantNameArray[$i] = $NameArray[$i] 
        }
    }

}

function Add-VariantToLibrary {

    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserLibrary,

        [Parameter(Mandatory)]
        [array]$ObjectNameArray,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$NewExtension,

        [Parameter(Mandatory)]
        [string]$NewName,

        [Parameter(Mandatory)]
        [string]$VariantName
    )

    $PublishingType = $ObjectNameArray[0]
    $Convention = $ObjectNameArray[1]
    $Artist = $ObjectNameArray[2]
    $Title = $ObjectNameArray[3]

    if ($PublishingType -eq "Manga") {

        $UserLibrary.$Artist.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$($PathsLibrary.Artists)\$Artist\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp
        }

        # Add $NewObjectName - not $VariantObjectName
        $null = $UserLibrary.$Artist.$Title.VariantList.add($NewName)
    }
    elseif ($PublishingType -eq "Doujinshi") {

        $UserLibrary.$Convention.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp # change timestamp format here to yyyy-MM-dd
        }
        
        $null = $UserLibrary.$Convention.$Title.VariantList.add($NewName)
    }
    elseif ($PublishingType -eq "Anthology") {

        $UserLibrary.$Artist.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$PathsLibrary.Anthologies\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp
        }

        $null = $UserLibrary.$Artist.$Title.VariantList.add($NewName)
    }
}

function Add-Duplicate{
    <#
    .DESCRIPTION
        Add duplicates to Duplicates-Hashtable.

        The key is the normalized and sanitized Title.
        The value is a list of paths.
        A title can have more than one duplicate.
    #>

    Param(
        [Parameter(Mandatory)]
        [hashtable]$Duplicates,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$ObjectPath
    )

    if (!($Duplicates.ContainsKey($Title))) {
        $Duplicates[$Title] = [List[string]]::new()
        $null = $Duplicates.$Title.Add($ObjectPath)
    }
    else {
        $null = $Duplicates.$Title.Add($ObjectPath)
    }
}

function Remove-FromLibrary{

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
    $VariantName    = $ObjectSelector.VariantObjectName
    
    $NewName = "$($ObjectSelector.NewName)"

    if($PublishingType -eq "Anthology"){

        if($VariantName -ne ""){
            $UserLibrary.$Artist.$Title.remove($VariantName)
        }
        else{
            $UserLibrary.$Artist.$Title.remove($NewName)
        }
    }
    elseif($PublishingType -eq "Manga"){

        if($VariantName -ne ""){
            $UserLibrary.$Artist.$Title.remove($VariantName)
        }
        else{
            $UserLibrary.$Artist.$Title.remove($NewName)
        }
    }
    elseif($PublishingType -eq "Doujinshi"){
        
        if($VariantName -ne ""){
            $UserLibrary.$Convention.$Title.remove($VariantName)
        }
        else{
            $UserLibrary.$Convention.$Title.remove($NewName)
        }
    }
}

function Remove-ComicInfo{
    <# 
    .DESCRIPTION
        Removes a ComicInfo-file from .\Library\ComicInfoFiles
        if the object fails to  copy.
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$ObjectSelector
    )

    $VariantName = $ObjectSelector.VariantObjectName
    $NewName = "$($ObjectSelector.NewName)"

    if($VariantName -ne ""){
        Remove-Item -LiteralPath "$($PathsLibrary.ComicInfoFiles)\$VariantName" -Recurse -Force
    }
    else{
        Remove-Item -LiteralPath "$($PathsLibrary.ComicInfoFiles)\$NewName" -Recurse -Force
    }
}

function Export-AsCSV {

    <# 
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
#endregion


### Begin: Main ###

Show-String -StringArray (" _ _ _     _                      _          _____ _____         _   ",
    "| | | |___| |___ ___ _____ ___   | |_ ___   |  |  |   __|___ ___| |_ ",
    "| | | | -_| |  _| . |     | -_|  |  _| . |  |     |__   | . |  _|  _|",
    "|_____|___|_|___|___|_|_|_|___|  |_| |___|  |__|__|_____|___|_| |_|  `n",
    "=====================================================================",
    " ",
" ")


### Begin: Setup ###

$ScriptVersion = "V0.1"

$RuntimeStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'
$SortingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'
$CopyingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'

$RuntimeStopwatch.Start()

### Begin: Assert-PSVersion >= 5.1 ###
$VersionMajor = $PSVersionTable.PSVersion.Major
$VersionMinor = $PSVersionTable.PSVersion.Minor
$PSVersion = "$VersionMajor.$VersionMinor"

if ($VersionMajor -lt 5) {
    Write-Output "This script requires at least Powershell Version 5.1`nExiting Script."
    exit
}
elseif (($VersionMajor -eq 5) -and ($VersionMinor -lt 1)) {
    Write-Output "This script requires at least Powershell Version 5.1`nExiting Script."
    exit
}
### End: Assert-PSVersion >= 5.1 ###


### Begin: Assert-7ZipInstalled ###
$64BitPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
$32BitPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

$App64Found = $False
$App32Found = $False

if([System.Environment]::Is64BitOperatingSystem){
    $64BitApps = Get-ItemProperty "HKLM:\$64BitPath"
    foreach($App in $64BitApps){
        if($App.Publisher -eq "Igor Pavlov"){
            $SevenZip64 = $App
            $App64Found = $True
        }
    }
    if($App64Found -eq $False){
        $32BitApps = Get-ItemProperty "HKLM:\$32BitPath"
        foreach($App in $32BitApps){
            if($App.Publisher -eq "Igor Pavlov"){
                $SevenZip32 = $App
                $App32Found = $True
            }
        }      
    }
}
else{
    $32BitApps = Get-ItemProperty "HKLM:\$32BitPath"
    foreach($App in $32BitApps){
        if($App.Publisher -eq "Igor Pavlov"){
            $SevenZip32 = $App
            $App32Found = $True
        }
    }        
}

if($App64Found -eq $true){
    $SevenZipPath = $SevenZip64.InstallLocation
}
elseif($App32Found -eq $true){
    $SevenZipPath = $SevenZip32.InstallLocation
}
else{
    Write-Information -MessageData "7-Zip not found. Exiting" -InformationAction Continue
    exit
}
### End: Assert-7ZipInstalled ###


$7zip = "$SevenZipPath\7z.exe"
### End: Setup ###



### Begin: ProgramData-Paths ###

# It's important that this is an ordered dictionary.
# Otherwise InitializeScript might try to create files/folders in folders that don't exist yet.

$PathsProgram = [ordered]@{

    "Parent"      = "$HOME\AppData\Roaming";
    "Base"        = "$HOME\AppData\Roaming\HSort";

    "AppData"     = "$HOME\AppData\Roaming\HSort\ApplicationData";

    "Libs"        = "$HOME\AppData\Roaming\HSort\LibraryFiles";

    "Settings"    = "$HOME\AppData\Roaming\HSort\Settings";
    "TxtSettings" = "$HOME\AppData\Roaming\HSort\Settings.txt";

    "Copied"      = "$HOME\AppData\Roaming\HSort\CopiedObjects";
    "Skipped"     = "$HOME\AppData\Roaming\HSort\SkippedObjects";

    "Tmp"         = "$HOME\AppData\Roaming\HSort\Temp"
}

### End: ProgramData-Paths ###


### Begin:  ProgramState ###

$InitializeScript_ExitCode = Initialize-Script -ScriptVersion $ScriptVersion -PathsProgram $PathsProgram

# Script continues IFF $InitializeScript_ExitCode -le 0.
if ($InitializeScript_ExitCode -le 0) {

    $SettingsHt = Import-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml"
    $ConfirmSettings_ExitCode = Confirm-Settings -CurrentSettings $SettingsHt

    if ($ConfirmSettings_ExitCode -eq 0) {
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
    else {
        exit
    }
}
else {
    exit
}

### End:  ProgramState ###



### Begin: Paths-LibraryFolder ###

$global:PathsLibrary = [ordered]@{

    "Parent"         = "$($SettingsHt.Target)";
    "Base"           = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)";

    "Library"        = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)";

    "Source"         = $SettingsHt.Source;

    "Artists"        = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Artists";
    "Conventions"    = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Conventions";
    "Anthologies"    = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)\Anthologies";

    "ComicInfoFiles" = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\ComicInfoFiles";

    "Logs"           = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\Logs"
}

### End: Paths-LibraryFolder ###



### Begin: OnInitialization ###

# No Libraries exist OR no library of this name exists.
if (($InitializeScript_ExitCode -eq 0) -or ($InitializeScript_ExitCode -eq -1)) {

    Write-Information -MessageData "Creating Library structure..." -InformationAction Continue

    foreach ($Path in $PathsLibrary.Keys) {
        if ($Path -ne "Source" -and $Path -ne "Parent") {
            $null = New-Item -ItemType "directory" -Path $PathsLibrary.$Path
        }
    }

    # UserLibraryXML to serialize
    $UserLibrary = @{}

    # Hashtable containing all visited objects (by object path).
    # Checked against ToProcessList.
    $VisitedObjects = @{}

    $Duplicates = @{}
}

# If Library(Name) is in SettingsArchive and ParentDirectory of Library is unchanged.
elseif($InitializeScript_ExitCode -eq -2) {

    $UserLibrary = Import-Clixml -Path "$($PathsProgram.Libs)\UserLibrary $($SettingsHt.LibraryName).xml"

    try {
        $VisitedObjects = Import-Clixml -Path "$($PathsProgram.AppData)\VisitedObjects $($SettingsHt.LibraryName).xml"
    }
    catch {
        $VisitedObjects = @{}
    }


    try {
        $Duplicates = Import-Clixml -Path "$($PathsProgram.AppData)\Duplicates $($SettingsHt.LibraryName).xml"
    }
    catch {
        $Duplicates = @{}
    }
}

<#
.NOTES
    Accessed in ObjectData.psm1
    for Get-FileHash
.DESCRIPTION
    SafeCopyFlag = 0
    Calculate and compare Source and Target hash of Files
    or Source and Target size for Folders

    SafeCopyFlag = 1
    Don't calculate file-hash/folder-size, instead set them to 0
    so that a comparison always evaluates to true.
.SOURCE
    GetObjects.psm1
#>

if($SettingsHt.SafeCopy -eq "False"){
    $Global:SafeCopyFlag = 1
}
else{
    $Global:SafeCopyFlag = 0
}

### End:  OnInitialization ###



### Begin: GetObjects ###

New-Graph -SourceDir $PathsLibrary.Source

$FoundObjectsHt = @{}
$FoundObjectsHt = Get-Objects -SourceDir $PathsLibrary.Source

$ToProcessLst = [List[object]]::new()
$ToProcessLst = $FoundObjectsHt.ToProcess

$SkippedObjects = @{} #< Why Global? 
$SkippedObjects = $FoundObjectsHt.SkippedObjects

# Folders that contain at least one file 
# with an unsupported extension.
$BadFolderCounter = $FoundObjectsHt.BadFolderCounter

$WrongExtensionCounter = $FoundObjectsHt.WrongExtensionCounter

$EventCounter = [EventCounter]::New()
$EventCounter.ToSort = ($ToProcessLst.Count)
$EventCounter.EFC = ($BadFolderCounter + $WrongExtensionCounter)
$EventCounter.AllObjects = ($ToProcessLst.Count + $BadFolderCounter + $WrongExtensionCounter)

### End: GetObjects ###



### Begin: Sorting ###

Show-String -StringArray (" ",
    "Sorting objects. Please wait.",
    " ")

$SortingStopwatch.Start()

<# 
    .DESCRIPTION
    The hashtable that is used in COPYING to copy
    objects to their proper destination.

    Structure
    <$Object.Name> = <$ObjectProperties>
#>
$ToCopy = @{}

# Allows to access UserLibrary-Hashtable from to ToCopy, and therefore
# to remove an Object from UserLibrary in COPYING if it fails to copy.
$LibraryLookUp = @{}

# For the progress bar. Incremented, when object is in ToProcessLst
# but Object is not in VisitedObjects _yet_ .
$SortingProgress = 0

# SortedObjectsCounter
# Becomes the Value of each Key added to VisitedObjects, but has no real function.
$SortedObjectsCounter = 0

$TagsHt = Import-Tags #< Get custom tags, if Tags.txt exists in HSort

foreach ($Object in $ToProcessLst) {
    
    $NoMatchFlag = 0
    $IsDuplicate = $false
    $IsVariant = $false

    $ObjectName = $Object.Name
    $ObjectPath = $Object.FullName
    $IsFile     = ($Object -is [system.io.fileinfo])

    Write-Information -MessageData "[Sorting]$ObjectName" -InformationAction Continue

    if ($IsFile) {

        $ObjectNameNEX = $ObjectName.Substring(0, $ObjectName.LastIndexOf('.'))
        $Extension = [System.IO.Path]::GetExtension($Object.FullName)

        # Date: The date the object was created on your system.
        $CreationDate = ([System.IO.File]::GetLastWriteTime($Object.FullName)).ToString("yyyy-MM-dd")
        
    }
    else {

        $ObjectNameNEX = $ObjectName
        $Extension     = "Folder"

        $FolderFirstElement = Get-ChildItem -LiteralPath $Object.FullName -Force -File | Select-Object -First 1
        $CreationDate = $FolderFirstElement.LastWriteTime.ToString("yyyy-MM-dd")

    }

    if ( (!($VisitedObjects.ContainsKey($ObjectPath))) ) {

        $SortedObjectsCounter += 1

        $null = $VisitedObjects.Add("$ObjectPath", $SortedObjectsCounter)

        $NormalizedObjectName = Format-ObjectName -ObjectNameNEX $ObjectNameNEX

        # Doujinshji
        if($NormalizedObjectName -match "\A\((?<Convention>[^\)]*)\)\s\[(?<Artist>[^\]]*)\](?<Title>[^\[{(]*)(?<Meta>.*)"){
            $ObjectNameArray = ("Doujinshi", $Matches.Convention, $Matches.Artist, $Matches.Title, $Matches.Meta)
        }
        # Anthologies
        elseif($NormalizedObjectName -match "\A\W(Anthology)\W(?<Title>[^\[{(]*)(?<Meta>.*)"){

            # Every Anthology has $Artist defined as "Anthology" !
            $ObjectNameArray = ("Anthology", "", "Anthology", $Matches.Title, $Matches.Meta)
        }
        # Manga
        elseif($NormalizedObjectName -match "\A\[(?<Artist>[^\]]*)\](?<Title>[^\[{(]*)(?<Meta>.*)"){
            $ObjectNameArray = ("Manga", "", $Matches.Artist, $Matches.Title, $Matches.Meta)
        }
        # NoMatch
        else {
            $NoMatchFlag = 1
            $EventCounter.AddNoMatch()
            Add-Skipped -SkippedObjects $SkippedObjects -Object $Object -Reason "NoMatch" -Extension $Extension
        }

        if($NoMatchFlag -eq 0){

            Edit-ObjectNameArray -ObjectNameArray $ObjectNameArray

            $PublishingType = $ObjectNameArray[0]; $Convention = $ObjectNameArray[1]; $Artist = $ObjectNameArray[2]; $Title = $ObjectNameArray[3]; $MetaToken = $ObjectNameArray[4]

            $Collection = Select-Collection -ObjectNameArray $ObjectNameArray

            # Import-Tags moved up. Import tags only once!
            
            $TagsCSV = Select-Tags -MetaTokenString $MetaToken -Title $Title -TagsHt $TagsHt

            $NewObjectName = New-ObjectName -ObjectNameArray $ObjectNameArray
    
            [hashtable]$ObjectMeta = Write-Meta -ObjectNameArray $ObjectNameArray -TagsCSV $TagsCSV -CreationDate $CreationDate
    
            [hashtable]$ObjectProperties = Write-Properties -Object $Object -NewName $NewObjectName -NameArray $ObjectNameArray -NameNEX $ObjectNameNEX -Ext $Extension
    
            [hashtable]$ObjectSelector = New-Selector -NameArray $ObjectNameArray -NewName $NewObjectName
    
            if (!$UserLibrary.ContainsKey($Collection)) {
    
                $UserLibrary[$Collection] = @{}
                
                Add-TitleToLibrary -UserLibrary $UserLibrary -ObjectNameArray $ObjectNameArray -Object $Object -NewName $NewObjectName
                
                $null = New-Item -Path ($ObjectProperties.ObjectTarget) -ItemType "directory"
                
                $EventCounter.AddSet($PublishingType)
    
            }
            elseif ( ! $UserLibrary.$Collection.ContainsKey($Title) ) {
    
                Add-TitleToLibrary -UserLibrary $UserLibrary -ObjectNameArray $ObjectNameArray -Object $Object -NewName $NewObjectName
                
                $EventCounter.AddTitle($PublishingType)
            
            }
            elseif ($UserLibrary.$Collection.ContainsKey($Title) -and (! $UserLibrary.$Collection.$Title.VariantList.Contains($NewObjectName))) {

                $IsVariant = $true
                
                $VariantNameArray = ("PublishingType","Convention","Artist","Title","Meta")
                New-VariantNameArray -NameArray $ObjectNameArray -VariantNameArray $VariantNameArray
                $VariantObjectName = New-ObjectName -ObjectNameArray $VariantNameArray
                
                # Update ObjectNewName in ObjectProperties. Needed in Copying.
                $ObjectProperties.ObjectNewName = $VariantObjectName
                # Special ObjectSelector for variants.
                $ObjectSelector.VariantObjectName = "$VariantObjectName"
                
                $ObjectMeta.Tags = "$($ObjectMeta.Tags),Variant"
                
                Add-VariantToLibrary -UserLibrary $UserLibrary -ObjectNameArray $ObjectNameArray -NewExtension $NewExtension -NewName $NewObjectName -VariantName $VariantObjectName
                $EventCounter.AddTitle($PublishingType)
                $EventCounter.AddVariant()
    
            }
            else {
                $IsDuplicate = $true
                $EventCounter.AddDuplicate()
    
                Add-Duplicate -Duplicates $Duplicates -Title $Title -ObjectPath $ObjectPath
                Add-Skipped -SkippedObjects $SkippedObjects -Object $Object -Reason "Duplicate" -Extension $Extension
            }
            
            if($IsDuplicate -eq $false){

                $null = $ToCopy.add($ObjectName, $ObjectProperties)
                $LibraryLookUp.Add($ObjectName, $ObjectSelector)

                if($IsVariant -eq $true){
                    $null = New-Item -ItemType "directory" -Path "$($PathsLibrary.ComicInfoFiles)\$VariantObjectName"
                    New-ComicInfoFile -ObjectMeta $ObjectMeta -Path "$($PathsLibrary.ComicInfoFiles)\$VariantObjectName"
                }
                else{
                    $null = New-Item -ItemType "directory" -Path "$($PathsLibrary.ComicInfoFiles)\$NewObjectName"
                    New-ComicInfoFile -ObjectMeta $ObjectMeta -Path "$($PathsLibrary.ComicInfoFiles)\$NewObjectName"
                }
            }
        }
    }
    else { #< Object already visited.
        $EventCounter.AddEFC()
    }
    $SortingProgress += 1
    $SortingCompleted = ($SortingProgress / $EventCounter.ToSort) * 100
    Write-Progress -Id 0 -Activity "Sorting" -Status "$ObjectNameNEX" -PercentComplete $SortingCompleted
}

$VisitedObjects |  Export-Clixml -Path "$($PathsProgram.AppData)\VisitedObjects $($SettingsHt.LibraryName).xml" -Force
$Duplicates |  Export-Clixml -Path "$($PathsProgram.AppData)\Duplicates $($SettingsHt.LibraryName).xml" -Force

$TotalSortingTime = ($SortingStopwatch.Elapsed).toString()
$SortingStopwatch.Stop()

Write-Output "Creating folder structure: Done`n"
Write-Output "Analyzing objects: Done`n"
Start-Sleep -Seconds 1.0

$EventCounter.ComputeToCopy()

### End: Sorting ###



### Begin: Copying ###

Show-String -StringArray (" ","Copying objects. This can take a while."," ")

# Hashtable of successfully copied objects.
# This allows the user to delete the successfully
# sorted and copied Objects if they wish to do so
# by running DeleteCopiedObjects.ps1 from Tools.
$CopiedObjects = @{}

# For progress bar.
# Ignores if the object actually gets copied successfully or not.
$CopyProgress = 0

$CopyingStopwatch.Start()

foreach ($ObjectName in $ToCopy.Keys) {

    $CopyProgress += 1
    $CopyCompleted = ($CopyProgress / $EventCounter.ToCopy) * 100
    Write-Progress -Id 1 -Activity "Copying" -Status "$ObjectName" -PercentComplete $CopyCompleted

    $Parent  = $ToCopy.$ObjectName.ObjectParent
    $Target  = $ToCopy.$ObjectName.ObjectTarget
    $Name    = $ToCopy.$ObjectName.ObjectName
    $NewName = $ToCopy.$ObjectName.ObjectNewName

    # Define XML source
    $XML = "$($PathsLibrary.ComicInfoFiles)\$NewName\ComicInfo.xml"
    $XMLParent = "$($PathsLibrary.ComicInfoFiles)\$NewName"

    ### Begin: CopyArchive ###

    if ($ToCopy.$ObjectName.Extension -ne "Folder") {

        <#
        .DESCRIPTION
            Possible errors:
            Probable case: A file of this NAME already exists.
            NAME is here identical to the name of the source file.
        #>

        $null = robocopy $Parent $Target $Name  /njh /njs
        $RobocopyExitCode = $LASTEXITCODE

        if ($RobocopyExitCode -eq 1) {

            if($SafeCopyFlag -eq 0){
                $ToCopy.$ObjectName.TargetHash =
                (Get-FileHash -LiteralPath "$Target\$Name" -Algorithm MD5).hash
            }
            else{
                $ToCopy.$ObjectName.TargetHash = 0
            }

            if ($ToCopy.$ObjectName.SourceHash -eq $ToCopy.$ObjectName.TargetHash) {

                $ArchiveName = "$NewName$($ToCopy.$ObjectName.NewExtension)"

                <#
                .DESCRIPTION
                    Possible error:
                    A file of this NAME already exists.
                    NAME is the newly sanitized name from SORTING with an added extension.
                #>

                try {
                    Rename-Item -LiteralPath "$Target\$Name" -NewName $ArchiveName
                }
                catch {

                    Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "RenamingError"
                    Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                    Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                    Remove-Item -LiteralPath $XMLpath -Force

                    $EventCounter.AddCopyError()
                    Continue
                }

                ### Begin: InsertXML ###

                <#
                .DESCRIPTION
                    Possible errors:
                        1.) 7Zip throws an error - probably because the archive is corrupt,
                            when trying to move ComicInfo.xml into the archive.
                        2.) XML doesn't exist.

                .NOTES
                    PS7 doesn't catch SeZipErrors.
                    PS5.1 doesn't save the 7z-ExitCode to the $SevenZipError array.
                    This means $SevenZipError.length is _always_ 0, what requires the code below...
                #>

                if ($VersionMajor -eq 5){

                    try {

                        & $7zip a -bsp0 -bso0 "$Target\$ArchiveName" $XML *>&1
                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)
                        $EventCounter.AddCopyGood()

                    }
                    catch {

                        Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "XMLinsertionError"
                        Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-Item -LiteralPath "$Target\$ArchiveName" -Force

                        $EventCounter.AddCopyError()
                    }

                }
                elseif ($VersionMajor -gt 5) {

                    $SevenZipError = & $7zip a -bsp0 -bso0 "$Target\$ArchiveName" $XML *>&1

                    if ($SevenZipError.length -eq 0) {

                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)
                        $EventCounter.AddCopyGood()

                    }
                    else {

                        Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "XMLinsertionError"
                        Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-Item -LiteralPath "$Target\$ArchiveName" -Force

                        $EventCounter.AddCopyError()
                    }
                }
                ### End: InsertXML ###
            }
            else {

                Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "HashMismatch"
                Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                $EventCounter.AddCopyError()
            }
        }
        else {

            Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "RobocopyError: $($RobocopyExitCode)"
            Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

            $EventCounter.AddCopyError()
        }
    }
    ### End: CopyArchive ###

    ### Begin: CopyFolder ###
    elseif ($ToCopy.$ObjectName.Extension -eq "Folder") {

        <#
        .DESCRIPTION
            Create NewName folder.
        .NOTES
            Possible errors:
            A folder of this NAME alredy exists.
            NAME is the newly sanitized name from SORTING.
        #>

        try {
            $null = New-Item -Path $Target -Name $NewName -ItemType "directory"
        }
        catch {

            Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "ErrorCreatingFolder"
            Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

            $EventCounter.AddCopyError()
            Continue
        }

        $null = robocopy ($ToCopy.$ObjectName.ObjectSource) "$Target\$NewName"  /njh /njs
        $RobocopyExitCode = $LASTEXITCODE

        if ($RobocopyExitCode -eq 1) {

            if($SafeCopyFlag -eq 0){
                # Get the folder size (instead of a hash) _before_ inserting the ComicInfo.XML file.
                $ToCopy.$ObjectName.TargetHash =
                    ((Get-ChildItem -LiteralPath "$Target\$NewName") | Measure-Object -Sum Length).sum
            }
            else{
                $ToCopy.$ObjectName.TargetHash = 0
            }

            if ($ToCopy.$ObjectName.SourceHash -eq $ToCopy.$ObjectName.TargetHash) {
                <#
                .DESCRIPTION
                    Try to create an Archive with the copied folder's content
                    and delete the (unzipped) folder if successful.
                .NOTES
                    Possible errors:
                    Generic 7Zip error.
                #>

                $null = robocopy $XMLParent "$Target\$NewName" "ComicInfo.xml"  /njh /njs

                $Target7zip = "$Target\$NewName.cbz"; $Source7zip = "$Target\$NewName\*"
                
                if ($VersionMajor -eq 5) {

                    try {
                        & $7zip a -mx3 -bsp0 -bso0 $Target7zip $Source7zip *>&1
                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)

                        $EventCounter.AddCopyGood()
                    }
                    catch {

                        Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "ErrorCreatingArchive"
                        Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                        $EventCounter.AddCopyError()
                    }

                }
                elseif ($VersionMajor -gt 5) {

                    $SevenZipError = & $7zip a -mx3 -bsp0 -bso0 $Target7zip $Source7zip *>&1

                    if ($SevenZipError -eq 0) {

                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)

                        $EventCounter.AddCopyGood()
                    }
                    else {

                        Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "ErrorCreatingArchive"
                        Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                        $EventCounter.AddCopyError()
                    }
                }
            }
            else { #< Hash mismatch

                Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "HashMismatch"
                Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
                Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                $EventCounter.AddCopyError()
            }
            # Whatever happens above, remove the unzipped folder.
            Remove-Item -LiteralPath "$Target\$NewName" -Recurse -Force
        }
        else {

            Add-NoCopy -SkippedObjects $SkippedObjects -Object $ToCopy.$ObjectName -Reason "RobocopyError: $($RobocopyExitCode)"
            Remove-FromLibrary -UserLibrary $UserLibrary -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

            $EventCounter.AddCopyError()
        }
    }
    ### End: CopyFolder ###

    Add-LogEntry -ObjectProperties $ToCopy.$ObjectName -Path $PathsLibrary.Logs
}

$CopyingStopwatch.Stop()
$RuntimeStopwatch.Stop()
$TotalCopyingTime = ($CopyingStopwatch.Elapsed).toString()
$TotalRuntime = ($RuntimeStopwatch.Elapsed).toString()

### End: Copying ###



### Begin: Finalizing ###

Write-Information -MessageData "Finalizing. Please wait." -InformationAction Continue

# Write Log-Files
foreach ($Reason in $SkippedObjects.Keys) {
    foreach($Object in $SkippedObjects.$Reason.Keys){

        Add-SkippedLogEntry -SkippedObjectProperties $SkippedObjects.$Reason.$Object -Path $PathsLibrary.Logs
    }
}


$UserLibrary | Export-Clixml -Path "$($PathsProgram.Libs)\UserLibrary $($SettingsHt.LibraryName).xml" -Force

# Serialize SkippedObjects - necessary for CopySkippedObjects.ps1 in Tools.
$SkippedObjects | Export-Clixml -Path "$($PathsProgram.Skipped)\Skipped $($SettingsHt.LibraryName).xml" -Force

# Serialize CopiedObjectsHt - necessary DeleteOriginals.ps1 in Tools.
$CopiedObjects | Export-Clixml -Path "$($PathsProgram.Copied)\CopiedObjects $($SettingsHt.LibraryName).xml" -Force

$EventCounter.ComputeSkipped()


Show-String -StringArray (" ","Script finished"," ")

$Summary = @"

SUMMARY [$Timestamp]
=================================================

[ ScriptVersion: $ScriptVersion ]
[ UsedPowershellVersion: $($VersionMajor).$($VersionMinor) ]
[ Settings.txt location: $($PathsProgram.TxtSettings) ]

> TotalRuntime: $TotalRuntime

> TotalSortingTime: $TotalSortingTime

> TotalCopyingTime: $TotalCopyingTime

# FoundObjects: $($EventCounter.AllObjects)
=================================================

> Successfully Copied Objects: $($EventCounter.CopyGood)

> CopyCompleteness (should be 0): $($EventCounter.ToCopy - $CopyProgress )

> Found $($EventCounter.Manga) Manga from $($EventCounter.Artists) Artists

> Found $($EventCounter.Doujinshi) Doujinshi from $($EventCounter.Conventions) Conventions

> Found $($EventCounter.Anthologies) Anthologies

> Found $($EventCounter.Variants) Variants


# Skipped Objects: $($EventCounter.Skipped)
=================================================

> Duplicates: $($EventCounter.Duplicates)

> UnmatchedObjects: $($EventCounter.NoMatch)

> Copy Errors: $($EventCounter.CopyErrors)

> Single files with
  wrong extension (.mov,.mp4,.pdf,...): $WrongExtensionCounter

> Leaf-Folders that contain at least
  one file that is NOT a (.jpg,.jpeg,.png,.txt): $BadFolderCounter

=================================================
=================================================

"@

$Summary | Out-Host
$Summary | Out-File -FilePath "$($PathsLibrary.Logs)\ProgramLog $Timestamp.txt" -Encoding unicode -Append -Force

# For debugging only
Copy-ScriptOutput -LibraryName $SettingsHt.LibraryName -PSVersion $PSVersion -Delete

$SortingProgress = 0
$CopyProgress = 0

### End: Finalizing ###
