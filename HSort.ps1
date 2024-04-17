using namespace System.Collections.Generic

$CallDirectory = Get-Location

Import-Module -Name "$CallDirectory\Modules\GetObjects\GetObjects.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateComicInfoXML\CreateComicInfoXML.psm1"
Import-Module -Name "$CallDirectory\Modules\CreateLogFiles\CreateLogFiles.psm1"
Import-Module -Name "$CallDirectory\Modules\InitializeScript\InitializeScript.psm1"
Import-Module -Name "$CallDirectory\Modules\DiskSpace\DiskSpace.psm1"

$ErrorActionPreference = "Stop"

Clear-Host

$global:Timestamp = Get-Date -Format "dd_MM_yyyy"

### Begin: Classes ###

Class Counter {

    <# 
        .SYNOPSIS
        Counting instances.
    #>

    [int]$AllObjects
    [int]$SortedObjects
    [int]$Anthologies
    [int]$Artists
    [int]$Manga
    [int]$Conventions
    [int]$Doujinshi
    [int]$Skipped
    [int]$Duplicates
    [int]$NoMatch
    [int]$Matches
    [int]$CopyErrors
    [int]$NotProcessed

    Counter() {

        $this.AllObjects = 0
        $this.SortedObjects = 0
        $this.Anthologies = 0
        $this.Artists = 0
        $this.Manga = 0
        $this.Conventions = 0
        $this.Doujinshi = 0
        $this.Skipped = 0
        $this.Duplicates = 0
        $this.NoMatch = 0
        $this.Matches = 0
        $this.CopyErrors = 0
        $this.NotProcessed = 0
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
        $this.Skipped += 1
    }

    [void]AddNoMatch() {
        $this.Skipped += 1
        $this.NoMatch += 1
    }

    [void]AddSkipped(){
        $this.Skipped += 1
    }

    [void]ComputeMatches(){
        $this.Matches = ($this.AllObjects - $this.Skipped)
    }

    [void]AddCopyError(){
        $this.CopyErrors += 1
    }

    [void]ComputeNotProcessed(){
        $this.NotProcessed = $this.Skipped + $this.CopyErrors
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

    $Timestamp = Get-Date -Format "dd_MM_yyyy HH_mm_ss"
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

function Read-Tokens {

    Param(
        [Parameter(Mandatory)]
        [string]$Path 
    )

    $TokenTable = @{}
    $TokenList = [List[string]]::new()

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        
        if ($line -ne "" -and -not $line.StartsWith("#")) {

            if ($line -match "^\[(?<Category>.+)\]") {

                if ($TokenList.Count -gt 0) {
                    $null = $TokenTable.Add($Category, $TokenList)
                    $TokenList = [List[string]]::new()
                }

                $Category = $Matches.Category

            }
            elseif ($Line.StartsWith("%")) {

                # The file has to end with any number of "%%%%"
                $TokenTable.Add($Category, $TokenList)

            }
            else {
                $null = $TokenList.Add($line)
            }

        }
    
    }
    return $TokenTable
}


function Import-TokenSet {

    Param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $TokenTable = Read-Tokens -Path $Path

    $TokenSet = [HashSet[string]]::new()
    
    foreach ($Category in $TokenTable.Keys) {
        foreach ($Token in $TokenTable.$Category) {
            $null = $TokenSet.Add($Token)
        }
    }

    return $TokenSet
}

function Get-TokenSet{

    <# 
        .NOTES
        $TokenSet is referenced by Select-MetaTags
    #>

    $TokenSetImportFlag = 0

    if(Test-Path "$($PathsProgram.AppData)\Tokens.txt"){

        $global:TokenSet = Import-TokenSet -Path "$($PathsProgram.AppData)\Tokens.txt"

        # In case $TokenSet evals to $null
        if(-not $TokenSet){
            $TokenSetImportFlag = 1
        }
    }
    else{
        $TokenSetImportFlag = 2
    }

    # If Tokens.txt doesn't exist or TokenSet evals to $null
    if ($TokenSetImportFlag -gt 0) {

        $TokenArray = ("English", "Eng", "Japanese", "Digital", "Censored", "Uncensored", "Decensored", "Full color")

        $global:TokenSet = [HashSet[string]]::new()

        # Since PS 5.1 has no range()...
        for ($i = 0; $i -le ($TokenArray.Length - 1); $i += 1 ) {
            $null = $TokenSet.Add($TokenArray[$i])
        }
    }
}

function Add-NoCopy {
    <#
    .DESCRIPTION
        Objects that caused errors in COPYING are stored in SkippedObjects.
        The content of SkippedObjects is formatted and
        written to SkippedObjects.txt in UserLibrary\Logs.
    #>

    Param(
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

    $null = $SkippedObjects.add("$($Object.ObjectSource)", $SkippedObjectProperties)

}

function Add-Skipped{
    <#
    .DESCRIPTION
        Objects that were skipped during SORTING.

        The content of SkippedObjects is formatted and
        written to SkippedObjects.txt in UserLibrary\Logs.
    #>

    Param(
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

    $null = $SkippedObjects.add("$($Object.FullName)", $SkippedObjectProperties)

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

function Convert-ListToString {
    Param(
        [Parameter(Mandatory)]
        [List[String]]$List
    )

    $String = ""                                
 
    for ($i = 0; $i -le ($List.Count - 1); $i++) {
        if ($i -le ($List.Count - 2)) {
            $String += "$($List[$i]),"
        }
        else {
            $String += "$($List[$i])"
        }
    }

    return $String
} 


function Select-MetaTags {
    <#
    .DESCRIPTION
        Select valid Tokens from Meta-Section of the object name.
        Return string of comma seperated tags
        to be stored in ObjectMeta.
    #>

    Param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$MetaString
    )

    if ($MetaString -ne "") {
        
        # Array
        MetaTokenArray = $MetaString.Split(",")

        $MetaTokenList = [List[string]]::new()

        for($i = 0; $i -le (MetaTokenArray.length -1); $i++){

            # Remove all non-word characters.
            # 02/04/2024 Use string-invariants
            $Token = MetaTokenArray[$i] -replace '[^a-zA-Z]', ''

            if ($TokenSet.Contains($Token)) {

                $MetaTokenList.Add($Token)

            }
        }
        
        $MetaTags = Convert-ListToString -List $MetaTokenList
    }
    else{
        $MetaTags = ""
    }


    return $MetaTags

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

function New-ObjectName {
    <#
    .INPUTS
        Array of sanitized and normalized tokens.
    .OUTPUTS
        NewObjectName as String
        The file name of any object stored in Library.
    .DESCRIPTION
        Put everything except the title in parenthesis.
    .PARAMETER
    #>

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray
    )

    $Name = ""

    for ($i = 1; $i -le ($ObjectNameArray.length - 1); $i += 1) {

        $Token = $ObjectNameArray[$i]

        if ($Token -ne "") {

            if ($i -eq 1) {
                $Token = "($Token)"
            }
            # If token is Artist
            elseif ($i -eq 2) {
                $Token = "($Token)"
            }
            elseif($i -eq 4){
                $Token = "($Token)"
            }

        }

        if ($Name -eq "") {
            $Name += $Token
        }

        # Add spce between Tokens
        else {
            $Name += " $Token"
        }
    }

    $Name = $Name.trim()

    return $Name
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

function Format-ObjectName{
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

function Write-Meta{
    <# 
    .NOTES
        16/04/2024
        Kavita 0.8 changed how collections work.

    #>    

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$TagsCSV,

        [Parameter(Mandatory)]
        [string]$CreationDate
    )

    $PublishingType = $ObjectNameArray[0]

    # Empty string if Object is not Doujinshi.
    $Convention     = $ObjectNameArray[1]

    # Artist is "Anthology" for Anthologies.
    $Artist         = $ObjectNameArray[2]
    $Title          = $ObjectNameArray[3]

    $DateArray = $CreationDate.split("-")
    $yyyy      = $DateArray[0]
    $MM        = $DateArray[1]

    switch($PublishingType)
    {

        "Manga" { $ObjectTarget = "$($PathsLibrary.Artists)\$Artist"; $SeriesGroup = $Artist; Break}

        "Doujinshi" { $ObjectTarget = "$($PathsLibrary.Conventions)\$Convention"; $SeriesGroup = $Convention; Break }

        "Anthology" { $ObjectTarget = $PathsLibrary.Anthologies; $SeriesGroup = $Artist; Break }
        
    }

    if(! $TagsCSV -eq ""){

        $Tags = "$PublishingType,$yyyy,$MM,$TagsCSV"
    }
    else{
        $Tags = "$PublishingType,$yyyy,$MM"
    }

    return @{
        PublishingType = $PublishingType;
        Format         = "Special";
        Convention     = $Convention;
        Artist         = $Artist;
        SeriesGroup    = $SeriesGroup;
        Series         = $Title;
        Title          = $Title;
        Tags           = $Tags;
        ObjectTarget   = $ObjectTarget
    }

}

function Add-TitleToLibrary{
    <# 
        .NOTES
        The title is only a PART of the object name.
        Rename NewName to something like: ReFoName (re-formatted name)
    #>
    
    Param(
        [Parameter(Mandatory)]
        [hashtable]$LibraryContent,

        [Parameter(Mandatory)]
        [array]$ObjectNameArray,

        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$NewName
    )

    $PublishingType = $ObjectNameArray[0]
    $Convention     = $ObjectNameArray[1]
    $Artist         = $ObjectNameArray[2]
    $Title          = $ObjectNameArray[3]                        

    
    if($PublishingType -eq "Manga"){

        $LibraryContent.$Artist[$Title] = @{
            "VariantList" = [List[string]]::new();
            $NewName = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Artists)\$Artist";
                FirstDiscovered = $Timestamp 
            }
        }

        # Add Base Variant
        $null = $LibraryContent.$Artist.$Title.VariantList.add($NewName)
    }

    elseif($PublishingType -eq "Doujinshi"){

        $LibraryContent.$Convention[$Title] = @{
            "VariantList" = [List[string]]::new();
            $NewName = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention";
                FirstDiscovered = $Timestamp;
            }
        }  

        $null = $LibraryContent.$Convention.$Title.VariantList.add($NewName)
    }

    elseif($PublishingType -eq "Anthology"){

        $LibraryContent.Anthology[$Title] = @{
            "VariantList" = [List[string]]::new();
           $NewName = @{
                ObjectSource    = $Object.FullName;
                ObjectLocation  = "$($PathsLibrary.Anthologies)\$Artist";
                FirstDiscovered = $Timestamp;
            }
        }

        $null = $LibraryContent.Anthology.$Title.VariantList.add($NewName)
    }
}

function Add-VariantToLibrary{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$LibraryContent,

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
    $Convention     = $ObjectNameArray[1]
    $Artist         = $ObjectNameArray[2]
    $Title          = $ObjectNameArray[3]

    if($PublishingType -eq "Manga"){

        $LibraryContent.$Artist.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$($PathsLibrary.Artists)\$Artist\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp
        }

        # Add $NewObjectName - not $VariantObjectName
        $null = $LibraryContent.$Artist.$Title.VariantList.add($NewName)
    }
    elseif($PublishingType -eq "Doujinshi"){

        $LibraryContent.$Convention.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$($PathsLibrary.Conventions)\$Convention\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp # change timestamp format here to yyyy-MM-dd
        }
        
        $null = $LibraryContent.$Convention.$Title.VariantList.add($NewName)
    }
    elseif($PublishingType -eq "Anthology"){

        $LibraryContent.$Artist.$Title[$VariantName] = @{
            ObjectSource    = $Object.FullName;
            ObjectLocation  = "$PathsLibrary.Anthologies\$VariantName$NewExtension";
            FirstDiscovered = $Timestamp
        }

        $null = $LibraryContent.$Artist.$Title.VariantList.add($NewName)
    }
}


function Write-Properties{

    Param(
        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$NewName,

        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [string]$NameNEX,

        [Parameter(Mandatory)]
        [string]$Ext
    )

    if (!($Ext -eq "Folder")) {

        # Kavita doesn't read ComicInfo.xml files from 7z archives.
        if ($Extension -eq '.zip') {
            $NewExtension = '.cbz'
        }
        elseif ($Extension -eq '.rar') {
            $NewExtension = '.cbr'
        }
        else {
            $NewExtension = $Extension
        }

        $SourceHash = (Get-FileHash -LiteralPath $Object.FullName -Algorithm MD5).hash

    }
    else {
        $NewExtension  = ""
        $SourceHash = (($Object | Get-ChildItem) | Measure-Object -Sum Length).sum
    }

    $PublishingType = $NameArray[0]
    $Convention     = $NameArray[1]
    $Artist         = $NameArray[2]

    if($PublishingType -eq "Anthology"){

        $ObjectTarget = "$($PathsLibrary.Anthologies)\$Artist"

    }
    elseif($PublishingType -eq "Manga"){

        $ObjectTarget = "$($PathsLibrary.Artists)\$Artist"

    }
    elseif($PublishingType -eq "Doujinshi"){

        $ObjectTarget = "$($PathsLibrary.Conventions)\$Convention"

    }

    return @{
        ObjectSource  = ($Object.FullName);
        ObjectNewName = $NewName;
        ObjectName    = $Object.Name;
        ObjectNameNEX = $NameNEX;
        Extension     = $Ext;
        NewExtension  = $NewExtension;
        ObjectTarget  = $ObjectTarget;
        ObjectParent  = (Split-Path -Parent $Object.FullName);
        SourceHash    = $SourceHash;
        TargetHash    = $Null
    }

}

function New-Selector{
    <# 
    .DESCRIPTION
    Allows to remove a Title from the Library
    in COPY, if copying fails

    .OUTPUTS
    Hashtable that allows to access a title in Library
    #>

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [string]$NewName,

        [string]$VariantName
    )

    begin{

        if($PSBoundParameters.ContainsKey('VariantName')){
            $VariantObjectName = $VariantName
        }
        else{
            $VariantObjectName = ""
        }
    }

    process{

        $PublishingType = $NameArray[0]
        $Convention     = $NameArray[1]
        $Artist         = $NameArray[2]
        $Title          = $NameArray[3]
        
        return @{
            "PublishingType" = $PublishingType
            "Artist" = $Artist
            "Convention" = $Convention
            "Title" = $Title
            "VariantObjectName" = $VariantObjectName
            "NewName" = $NewName
        }

    }
}

function Remove-FromLibrary{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$LibraryContent,

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
            $LibraryContent.$Artist.$Title.remove($VariantName)
        }
        else{
            $LibraryContent.$Artist.$Title.remove($NewName)
        }
    }
    elseif($PublishingType -eq "Manga"){

        if($VariantName -ne ""){
            $LibraryContent.$Artist.$Title.remove($VariantName)
        }
        else{
            $LibraryContent.$Artist.$Title.remove($NewName)
        }
    }
    elseif($PublishingType -eq "Doujinshi"){
        
        if($VariantName -ne ""){
            $LibraryContent.$Convention.$Title.remove($VariantName)
        }
        else{
            $LibraryContent.$Convention.$Title.remove($NewName)
        }
    }
}

function New-VariantNameArray{

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [array]$VariantNameArray
    )

    $PublishingType = $ObjectNameArray[0]
    $Convention = $ObjectNameArray[1]
    $Artist = $ObjectNameArray[2]
    $Title = $ObjectNameArray[3]
    $Meta = $ObjectNameArray[4]

    switch ($PublishingType) {

        "Manga" { $VariantNumber = ($LibraryContent.$Artist.$Title.VariantList.Count) - 1; Break }

        "Doujinshi" { $VariantNumber = ($LibraryContent.$Convention.$Title.VariantList.Count) - 1; Break }

        "Anthology" { $VariantNumber = ($LibraryContent.Anthology.$Title.VariantList.Count) - 1; Break }
        
    }

    $Title += " Variant $VariantNumber"

    $VariantNameArray[0] = $PublishingType
    $VariantNameArray[1] = $Convention
    $VariantNameArray[2] = $Artist
    $VariantNameArray[3] = $Title
    $VariantNameArray[4] = $Meta

}


function Read-Creator{

    <# 
        Until I can come up with a better name...

        Creator := $Convention for Doujinshi

        Creator := $Artist for Manga

        Creator := "Anthology" for Anthologies
        (And Artist:= "Anthology" as well...)

    #>    

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray
    )

    if($ObjectNameArray[0] -eq "Anthology"){
        $Creator = "Anthology"
    }
    
    elseif($ObjectNameArray[0] -eq "Manga"){
        $Creator = "$($ObjectNameArray[2])"
    }
    
    elseif($ObjectNameArray[0] -eq "Doujinshi"){

        $Creator = "$($ObjectNameArray[1])"

    }

    return $Creator
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

<# 
    function Select-TitleTags{
        Param(
            [Parameter(Mandatory)]
            [string]$Title
        )
        $Inseki = [System.Collections.Generic.HashSet[String]] @("Sister", "Onee", "Onee-chan", "Imouto","Brother")
        $TitleTokens = [System.Collections.Generic.HashSet[String]] @(($Title.Split(" ")))

        if($TitleTokens.Overlaps($Inseki)){

        }
    }
#>

### End: Functions ###

#endregion




Show-String -StringArray (" _ _ _     _                      _          _____ _____         _   ",
    "| | | |___| |___ ___ _____ ___   | |_ ___   |  |  |   __|___ ___| |_ ",
    "| | | | -_| |  _| . |     | -_|  |  _| . |  |     |__   | . |  _|  _|",
    "|_____|___|_|___|___|_|_|_|___|  |_| |___|  |__|__|_____|___|_| |_|  `n",
    "=====================================================================",
    " ",
    " ")


# trap {
#     Copy-ScriptOutput -LibraryName $SettingsHt.LibraryName -PSVersion $PSVersion -Delete
# }




### Begin: Setup ###

$ScriptVersion = "V004"

$7zip = "$env:ProgramFiles\7-Zip\7z.exe"

$RuntimeStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'

$SortingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'

$CopyingStopwatch = New-Object -TypeName 'System.Diagnostics.Stopwatch'

$RuntimeStopwatch.Start()

# Assert that PSVersion >= 5.1
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

$script:PathsLibrary = [ordered]@{

    "Parent"         = "$($SettingsHt.Target)";
    "Base"           = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)";

    "Library"        = "$($SettingsHt.Target)\$($SettingsHt.LibraryName)\$($SettingsHt.LibraryName)";

    "Source"      = $SettingsHt.Source;

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

    # LibraryContentXML to serialize
    $LibraryContent = @{}

    # Hashtable containing all visited objects (by object path).
    # Checked against ToProcessList.
    $VisitedObjects = @{}

    $Duplicates = @{}
}

# If Library(Name) is in SettingsArchive and ParentDirectory of Library is unchanged.
elseif($InitializeScript_ExitCode -eq -2) {

    $LibraryContent = Import-Clixml -Path "$($PathsProgram.Libs)\LibraryContent $($SettingsHt.LibraryName).xml"

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

### End:  OnInitialization ###



### Begin: GetObjects ###

New-Graph -SourceDir $PathsLibrary.Source

# Hashtable containing all skipped objects.
$FoundObjectsHt = @{}
$FoundObjectsHt = Get-Objects -SourceDir $PathsLibrary.Source

$ToProcessLst = [List[object]]::new()
$ToProcessLst = $FoundObjectsHt.ToProcess

$global:SkippedObjects = @{}
$SkippedObjects = $FoundObjectsHt.SkippedObjects

# Folders that contain at least one file with an unsupported extension.
$BadFolderCounter = $FoundObjectsHt.BadFolderCounter

$WrongExtensionCounter = $FoundObjectsHt.WrongExtensionCounter

$Counter = [Counter]::New()

$Counter.Skipped = ($BadFolderCounter + $WrongExtensionCounter)

$Counter.AllObjects = ($ToProcessLst.Count + $BadFolderCounter + $WrongExtensionCounter)

### End: GetObjects ###



### Begin: Sorting ###

Show-String -StringArray (" ",
    "Sorting objects. Please wait.",
    " ")

$SortingStopwatch.Start()

# Create hashtable with <$Object.Name> = <$ObjectProperties>"
# $ObjectProperties are defined in SORTING.
# $ObjectProperties allows us to copy and therefore sort the objects to the correct folders.
$ToCopy = @{}

# Allows to remove an Object from LibraryContent in COPYING
# in case the Object fails to copy.
# Allows to access LibraryContent-Hashtable from to ToCopy.
$LibraryLookUp = @{}

# For the progress bar. Incremented, when object is in ToProcessLst
# but Object is not in VisitedObjects _yet_ .
$SortingProgress = 0

# SortedObjectsCounter
# Becomes the Value of each Key added to VisitedObjects, but has no real function.
$SortedObjectsCounter = 0

# Get custom tokens, if Tokens.txt exists in AppData
Get-TokenSet

foreach ($Object in $ToProcessLst) {
    
    $NoMatchFlag = 0
    $IsDuplicate = $false
    $IsVariant = $false

    $ObjectName = $Object.Name
    $ObjectPath = $Object.FullName
    $IsFile     = ($Object -is [system.io.fileinfo])

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
            $Counter.AddNoMatch()
            Add-Skipped -Object $Object -Reason "NoMatch" -Extension $Extension
        }

        if($NoMatchFlag -eq 0){

            Edit-ObjectNameArray -ObjectNameArray $ObjectNameArray
            
            $PublishingType = $ObjectNameArray[0]; $Convention = $ObjectNameArray[1]; $Artist = $ObjectNameArray[2]; $Title = $ObjectNameArray[3]; $Meta = $ObjectNameArray[4]

            $Creator = Read-Creator -ObjectNameArray $ObjectNameArray
    
            $MetaTags = Select-MetaTags -MetaString $Meta
            
            $NewObjectName = New-ObjectName -ObjectNameArray $ObjectNameArray
    
            [hashtable]$ObjectMeta = Write-Meta -ObjectNameArray $ObjectNameArray -TagsCSV $MetaTags -CreationDate $CreationDate
    
            [hashtable]$ObjectProperties = Write-Properties -Object $Object -NewName $NewObjectName -NameArray $ObjectNameArray -NameNEX $ObjectNameNEX -Ext $Extension
    
            [hashtable]$ObjectSelector = New-Selector -NameArray $ObjectNameArray -NewName $NewObjectName
    
            if (!$LibraryContent.ContainsKey($Creator)) {
    
                $Counter.AddSet($PublishingType)
    
                $LibraryContent[$Creator] = @{}
                
                Add-TitleToLibrary -LibraryContent $LibraryContent -ObjectNameArray $ObjectNameArray -Object $Object -NewName $NewObjectName
                
                $null = New-Item -Path ($ObjectProperties.ObjectTarget) -ItemType "directory"
    
            }
            elseif ( ! $LibraryContent.$Creator.ContainsKey($Title) ) {
    
                $Counter.AddTitle($PublishingType)
    
                Add-TitleToLibrary -LibraryContent $LibraryContent -ObjectNameArray $ObjectNameArray -Object $Object -NewName $NewObjectName
            
            }
            elseif ($LibraryContent.$Creator.ContainsKey($Title) -and (! $LibraryContent.$Creator.$Title.VariantList.Contains($NewObjectName))) {

                $IsVariant = $true
    
                $Counter.AddTitle($PublishingType)
    
                $VariantNameArray = ("PublishingType","Convention","Artist","Title","Meta")
                New-VariantNameArray -NameArray $ObjectNameArray -VariantNameArray $VariantNameArray
                
                # Important: Overwrite $NewObjectname !
                $VariantObjectName = New-ObjectName -ObjectNameArray $VariantNameArray
    
                # Update ObjectNewName in ObjectProperties. Needed in Copying.
                $ObjectProperties.ObjectNewName = $VariantObjectName
                # Specific ObjectSelector for variants.
                $ObjectSelector.VariantObjectName = "$VariantObjectName"

                $ObjectMeta.Tags = "$($ObjectMeta.Tags),Variant"
    
                Add-VariantToLibrary -LibraryContent $LibraryContent -ObjectNameArray $ObjectNameArray -NewExtension $NewExtension -NewName $NewObjectName -VariantName $VariantObjectName
    
            }
            else {
                $IsDuplicate = $true
                $Counter.AddDuplicate()
    
                Add-Duplicate -Duplicates $Duplicates -Title $Title -ObjectPath $ObjectPath
                Add-Skipped -Object $Object -Reason "Duplicate" -Extension $Extension
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
        $SortingProgress += 1
        $SortingCompleted = ($SortingProgress / $Counter.AllObjects) * 100
        Write-Progress -Id 0 -Activity "Sorting" -Status "$ObjectNameNEX" -PercentComplete $SortingCompleted
    }
    # Object already visited.
    else {
        $Counter.AddSkipped()
    }
}

$VisitedObjects |  Export-Clixml -Path "$($PathsProgram.AppData)\VisitedObjects $($SettingsHt.LibraryName).xml" -Force
$Duplicates |  Export-Clixml -Path "$($PathsProgram.AppData)\Duplicates $($SettingsHt.LibraryName).xml" -Force

$TotalSortingTime = ($SortingStopwatch.Elapsed).toString()
$SortingStopwatch.Stop()

Write-Output "Creating folder structure: Done`n"
Write-Output "Analyzing objects: Done`n"
Start-Sleep -Seconds 1.0

$Counter.ComputeMatches()

### End: Sorting ###



### Begin: Copying ###

Show-String -StringArray (" ",
"Copying objects. This can take a while.",
" ")


# Hashtable of successfully copied objects.
# This allows the user to delete the successfully
# sorted and copied Objects if they wish to do so
# by running DeleteCopiedObjects.ps1 from Tools.
$CopiedObjects = @{}

# For progress bar.
# Ignores if the object actually gets copied successfully or not.
$CopyProgress = 0

# Set to 1 if SourceHash doesn't match TargetHash at least one time.
$IntegrityFlag = 0

$CopyingStopwatch.Start()

foreach ($ObjectName in $ToCopy.Keys) {

    $CopyProgress += 1
    $CopyCompleted = ($CopyProgress / $Counter.Matches) * 100
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

            $ToCopy.$ObjectName.TargetHash =
            (Get-FileHash -LiteralPath "$Target\$Name" -Algorithm MD5).hash

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
                    Add-NoCopy -Object $ToCopy.$ObjectName -Reason "RenamingError"

                    Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                    Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                    Remove-Item -LiteralPath $XMLpath -Force
                    $Counter.AddCopyError()
                    Continue
                }

                <#
                .DESCRIPTION
                    Possible errors:
                        1.) 7Zip throws an error - probably because the archive is corrupt,
                        when trying to move ComicInfo.xml into the archive.

                        2.) XML doesn't exist.

                .NOTES
                    PS7 doesn't catch SeZipErrors, and
                    PS5.1 doesn't save the ExitCode into the $SevenZipError array,
                    so that $SevenZipError.length is always 0,
                    what leads to the abomination below...
                #>

                ### Begin: InsertXML ###

                if ($VersionMajor -eq 5){

                    try {
                        & $7zip a -bsp0 -bso0 "$Target\$ArchiveName" $XML *>&1
                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)
                    }
                    catch {
                        Add-NoCopy -Object $ToCopy.$ObjectName -Reason "XMLinsertionError"

                        Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                        Remove-Item -LiteralPath "$Target\$ArchiveName" -Force
                        $Counter.AddCopyError()
                    }

                }
                elseif ($VersionMajor -gt 5) {

                    $SevenZipError = & $7zip a -bsp0 -bso0 "$Target\$ArchiveName" $XML *>&1

                    if ($SevenZipError.length -eq 0) {
                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)
                    }
                    else {
                        Add-NoCopy -Object $ToCopy.$ObjectName -Reason "XMLinsertionError"

                        Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName

                        Remove-Item -LiteralPath "$Target\$ArchiveName" -Force
                        $Counter.AddCopyError()
                    }
                }

                ### End: InsertXML ###
            }
            else {
                $IntegrityFlag += 1

                Add-NoCopy -Object $ToCopy.$ObjectName -Reason "HashMismatch"
                Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                $Counter.AddCopyError()
            }
        }
        else {
            Add-NoCopy -Object $ToCopy.$ObjectName -Reason "RobocopyError: $($RobocopyExitCode)"
            Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
            $Counter.AddCopyError()
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
            Add-NoCopy -Object $ToCopy.$ObjectName -Reason "ErrorCreatingFolder"
            Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
            $Counter.AddCopyError()
            Continue
        }

        $null = robocopy ($ToCopy.$ObjectName.ObjectSource) "$Target\$NewName"  /njh /njs
        $RobocopyExitCode = $LASTEXITCODE

        if ($RobocopyExitCode -eq 1) {

            # Get the folder size (instead of a hash) _before_ inserting the ComicInfo.XML file.
            $ToCopy.$ObjectName.TargetHash =
                ((Get-ChildItem -LiteralPath "$Target\$NewName") | Measure-Object -Sum Length).sum

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
                    }
                    catch {
                        Add-NoCopy -Object $ToCopy.$ObjectName -Reason "ErrorCreatingArchive"
                        Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                        $Counter.AddCopyError()
                    }

                }
                elseif ($VersionMajor -gt 5) {

                    $SevenZipError = & $7zip a -mx3 -bsp0 -bso0 $Target7zip $Source7zip *>&1

                    if ($SevenZipError -eq 0) {
                        $null = $CopiedObjects.Add($ObjectName, $ToCopy.$ObjectName)
                    }
                    else {
                        Add-NoCopy -Object $ToCopy.$ObjectName -Reason "ErrorCreatingArchive"
                        Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                        Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                        $Counter.AddCopyError()
                    }
                }
            }

            # Hash mismatch
            else {
                $IntegrityFlag += 1

                Add-NoCopy -Object $ToCopy.$ObjectName -Reason "HashMismatch"
                Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
                Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
                $Counter.AddCopyError()
            }
            # Whatever happens above, remove the unzipped folder.
            Remove-Item -LiteralPath "$Target\$NewName" -Recurse -Force
        }
        else {
            Add-NoCopy -Object $ToCopy.$ObjectName -Reason "RobocopyError: $($RobocopyExitCode)"
            Remove-FromLibrary -LibraryContent $LibraryContent -ObjectSelector $LibraryLookUp.$ObjectName
            Remove-ComicInfo -ObjectSelector $LibraryLookUp.$ObjectName
            $Counter.AddCopyError()
        }
    }
    ### End: CopyFolder ###

    Add-LogEntry -ObjectProperties $ToCopy.$ObjectName -Path $PathsLibrary.Logs
}

$SortingProgress = 0
$CopyProgress = 0

$TotalCopyingTime = ($CopyingStopwatch.Elapsed).toString()
$CopyingStopwatch.Stop()

$TotalRuntime = ($RuntimeStopwatch.Elapsed).toString()
$RuntimeStopwatch.Stop()

### End: Copying ###



### Begin: Finalizing ###

Write-Information -MessageData "Finalizing. Please wait." -InformationAction Continue

# Write Log-Files
foreach ($SkippedObjectProperties in $SkippedObjects.Keys) {
    Add-SkippedLogEntry -SkippedObjectProperties $SkippedObjects.$SkippedObjectProperties -Path $PathsLibrary.Logs
}

$LibraryContent | Export-Clixml -Path "$($PathsProgram.Libs)\LibraryContent $($SettingsHt.LibraryName).xml" -Force

# Serialize SkippedObjects - necessary for the CopySkippedObjects script.
$SkippedObjects | Export-Clixml -Path "$($PathsProgram.Skipped)\Skipped $($SettingsHt.LibraryName).xml" -Force

# Serialize CopiedObjectsHt - necessary for the DeleteOriginals script.
$CopiedObjects | Export-Clixml -Path "$($PathsProgram.Copied)\CopiedObjects $($SettingsHt.LibraryName).xml" -Force

$Counter.ComputeNotProcessed()



Show-String -StringArray (" ",
"Script finished", 
" ")


$Summary = @"

SUMMARY [$Timestamp]
=================================================

[ ScriptVersion: $ScriptVersion ]
[ UsedPowershellVersion: $($VersionMajor).$($VersionMinor) ]
[ Settings.txt location: $($PathsProgram.TxtSettings) ]

> TotalRuntime: $TotalRuntime

> TotalSortingTime: $TotalSortingTime

> TotalCopyingTime: $TotalCopyingTime

# FoundObjects: $($Counter.AllObjects)
=================================================

> Successfully Copied Objects: $($CopiedObjects.Count)

> Found $($Counter.Manga) Manga from $($Counter.Artists) Artists

> Found $($Counter.Doujinshi) Doujinshi from $($Counter.Conventions) Conventions

> Found $($Counter.Anthologies) Anthologies


# Objects not copied: $($Counter.NotProcessed)
=================================================

> Duplicates: $($Counter.Duplicates)

> UnmatchedObjects: $($Counter.NoMatch)

> Copy Errors: $($Counter.CopyErrors)

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

### End: Finalizing ###
