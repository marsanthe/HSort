
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

    $NewObName = ""

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
            elseif ($i -eq 4) {
                $Token = "($Token)"
            }

        }

        if ($NewObName -eq "") {
            $NewObName += $Token
        }

        # Add space between Tokens
        else {
            $NewObName += " $Token"
        }
    }

    $NewObName = $NewObName.trim()

    return $NewObName
}
function Select-Collection {
    <# 
    .OUTPUTS
    String
    #>

    <# 
        Until I can come up with a better name...

        Collection := $Convention for Doujinshi

        Collection := $Artist for Manga

        Collection := "Anthology" for Anthologies
        (And Artist:= "Anthology" as well...)

    #>    

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray
    )

    if ($ObjectNameArray[0] -eq "Anthology") {
        $Collection = "Anthology"
    }
    
    elseif ($ObjectNameArray[0] -eq "Manga") {
        $Collection = "$($ObjectNameArray[2])"
    }
    
    elseif ($ObjectNameArray[0] -eq "Doujinshi") {

        $Collection = "$($ObjectNameArray[1])"

    }

    return $Collection
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

function Read-TagsFromFile {

    Param(
        [Parameter(Mandatory)]
        [string]$Path 
    )

    $TagsHt = @{}
    $TagSet = [HashSet[string]]::new()

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        
        if ($line -ne "" -and -not $line.StartsWith("#")) {

            if ($line -match "^\[(?<Category>.+)\]") {

                if ($TagSet.Count -gt 0) {
                    $null = $TagsHt.Add($Category, $TagSet)
                    $TagSet = [List[string]]::new()
                }

                $Category = $Matches.Category

            }
            elseif ($Line.StartsWith("%")) {

                # The file has to end with any number of "%%%%"
                $TagsHt.Add($Category, $TagSet)

            }
            else {
                $null = $TagSet.Add(($line.Trim()))
            }
        }
    }

    #$TagsHt | Export-Clixml -Path "$($PathsLibrary.Logs)\TagsHt.xml" -Force

    return $TagsHt
}


function Import-Tags {
    <# 
    .NOTES
        $TagsHt is referenced by Select-Tags
    #>

    $ImportFlag = 0

    if (Test-Path "$CallDirectory\Tags.txt") {
        
        $TagsHt = Read-TagsFromFile -Path "$CallDirectory\Tags.txt"
        
        # In case $TagsHt evals to $null
        if (-not $TagsHt) {
            $ImportFlag = 1
        }
    }
    else {
        $ImportFlag = 2
    }

    # If Tags.txt doesn't exist or TagsHt evals to $null (Read error)
    if ($ImportFlag -gt 0) {

        $TagSetMeta = [System.Collections.Generic.HashSet[String]] @("english", "eng", "japanese", "digital", "censored", "uncensored", "decensored", "full color")

        $TagsHt = @{
            "Meta" = $TagSetMeta
        }
    }

    return $TagsHt
}
function Select-Tags {
    <#
    .DESCRIPTION
        Select valid Tokens from object name.
    #>

    Param(

        [array]$ObjectNameArray,
        
        [hashtable]$TagsHt

        #[AllowEmptyString()]
        #[string]$MetaTokenString


        #[string]$Title,

    )

    $TagList = [List[string]]::new()

    ### Select Tags from Title ###

    $Title = $ObjectNameArray[3]
    
    $Title = $Title.ToLower()
    $TitleTokenSet = [System.Collections.Generic.HashSet[String]] @(($Title.Split(" ")))

    foreach ($Category in $TagsHt.Keys) {
        if($TagsHt.$Category.Count -gt 0){
            if($Category -ne "Meta"){
                if ($TitleTokenSet.Overlaps($TagsHt.$Category)) {
                    $TagList.Add("$Category")
                }
            }
        }
    }

    ### Select Tags from Meta ###

    $MetaTokenString = $ObjectNameArray[4]

    if ($MetaTokenString -ne "") {
        
        # reduce variance
        $MetaTokenString = $MetaTokenString.ToLower()
        # Array
        $MetaTokenArray = $MetaTokenString.Split(",")

        for ($i = 0; $i -le ($MetaTokenArray.length - 1); $i++) {

            $Token = $MetaTokenArray[$i]

            if ($TagsHt.Meta.Contains($Token)) {

                $TagList.Add($Token)

            }
        }
 
    }

    if ($TagList.Count -gt 0) {
        $TagsCSV = Convert-ListToString -List $TagList
    }
    else{
        $TagsCSV = ""
    }

    return $TagsCSV

}

function Write-Meta {
    <# 
    .NOTES
        16/04/2024
        Kavita 0.8 changed how collections work.

    #>    

    Param(
        [Parameter(Mandatory)]
        [array]$ObjectNameArray,

        [Parameter(Mandatory)]
        [string]$CreationDate,

        [Parameter(Mandatory)]
        [hashtable]$TagsHT
    )

    $PublishingType = $ObjectNameArray[0]
    $Convention = $ObjectNameArray[1] # Empty string if Object is not Doujinshi.
    $Artist = $ObjectNameArray[2] # Artist is "Anthology" for Anthologies.
    $Title = $ObjectNameArray[3]

    $DateArray = $CreationDate.split("-")
    $yyyy = $DateArray[0]; $MM = $DateArray[1]

    $TagsCSV = Select-Tags -ObjectNameArray $ObjectNameArray -TagsHt $TagsHt

    switch ($PublishingType) {

        "Manga" { $ObjectTarget = "$($PathsLibrary.Artists)\$Artist"; $SeriesGroup = $Artist; Break }

        "Doujinshi" { $ObjectTarget = "$($PathsLibrary.Conventions)\$Convention"; $SeriesGroup = $Convention; Break }

        "Anthology" { $ObjectTarget = $PathsLibrary.Anthologies; $SeriesGroup = $Artist; Break }
        
    }

    if (! $TagsCSV -eq "") {

        $Tags = "$PublishingType,$yyyy,$MM,$TagsCSV"
    }
    else {
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

function Write-Properties {

    Param(
        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$NewName,

        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [string]$Ext
    )

    if (!($Ext -eq "Folder")) {

        # Kavita doesn't read ComicInfo.xml files from 7z archives.
        if ($Ext -eq '.zip') {
            $NewExtension = '.cbz'
        }
        elseif ($Ext -eq '.rar') {
            $NewExtension = '.cbr'
        }
        else {
            $NewExtension = $Ext
        }

        if($SafeCopyFlag -eq 0){

            $SourceHash = (Get-FileHash -LiteralPath $Object.FullName -Algorithm MD5).hash
        }
        else{
            $SourceHash = 0
        }

    }
    else {

        $NewExtension = ".cbz"

        if($SafeCopyFlag -eq 0){
            $SourceHash = (($Object | Get-ChildItem) | Measure-Object -Sum Length).sum
        }
        else{
            $SourceHash = 0
        }
    }

    $PublishingType = $NameArray[0]
    $Convention = $NameArray[1]
    $Artist = $NameArray[2]

    if ($PublishingType -eq "Anthology") {

        $ObjectTarget = "$($PathsLibrary.Anthologies)\$Artist"

    }
    elseif ($PublishingType -eq "Manga") {

        $ObjectTarget = "$($PathsLibrary.Artists)\$Artist"

    }
    elseif ($PublishingType -eq "Doujinshi") {

        $ObjectTarget = "$($PathsLibrary.Conventions)\$Convention"

    }

    return @{
        ObjectSource  = ($Object.FullName);
        ObjectNewName = $NewName;
        ObjectName    = $Object.Name;
        Extension     = $Ext; # Required to check if Object is file or folder in COPY
        NewExtension  = $NewExtension;
        ObjectTarget  = $ObjectTarget;
        ObjectParent  = (Split-Path -Parent $Object.FullName);
        SourceHash    = $SourceHash;
        TargetHash    = $Null
    }

}

function New-Selector {
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

    begin {

        if ($PSBoundParameters.ContainsKey('VariantName')) {
            $VariantObjectName = $VariantName
        }
        else {
            $VariantObjectName = ""
        }
    }

    process {

        $PublishingType = $NameArray[0]
        $Convention = $NameArray[1]
        $Artist = $NameArray[2]
        $Title = $NameArray[3]
        
        return @{
            "PublishingType"    = $PublishingType
            "Artist"            = $Artist
            "Convention"        = $Convention
            "Title"             = $Title
            "VariantObjectName" = $VariantObjectName
            "NewName"           = $NewName
        }

    }
}

Export-ModuleMember -Function New-ObjectName,Select-Collection,Convert-ListToString,Import-Tags,Select-Tags,Write-Meta,Write-Properties,New-Selector -Variable Name,Collection,String,TagsCSV,TagsHt