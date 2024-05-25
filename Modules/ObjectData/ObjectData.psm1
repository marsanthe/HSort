
function New-TargetName {
    <#
    .INPUTS
        Array of sanitized and normalized tokens.
    .OUTPUTS
        TargetName (== Title)
    #>

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray
    )

    $TargetName = $NameArray[3]
    $TargetName = $TargetName.trim()

    return $TargetName
}

function New-Id {
    <#
    .INPUTS
        Array of sanitized and normalized tokens.
    .OUTPUTS
        Id as string
    .DESCRIPTION
        Converts the sanitized $NameArray back to a string.
        The SHA1 of this string is the key of each object in UserLibrary
        [Collection -> Title -> SHA1(Id)]
    #>

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray
    )

    $Id = [system.String]::Join(" ", $NameArray)

    $Id = $Id.trim()

    return $Id
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
        [array]$NameArray
    )

    if ($NameArray[0] -eq "Anthology") {
        $Collection = "Anthology"
    }
    
    elseif ($NameArray[0] -eq "Manga") {
        $Collection = "$($NameArray[2])"
    }
    
    elseif ($NameArray[0] -eq "Doujinshi") {

        $Collection = "$($NameArray[1])"

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

        [array]$NameArray,
        
        [hashtable]$TagsHt

    )

    $TagList = [List[string]]::new()

    ### Select Tags from Title ###

    $Title = $NameArray[3]
    
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

    $MetaTokenString = $NameArray[4]

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
        [array]$NameArray,

        [Parameter(Mandatory)]
        [string]$CreationDate,

        [Parameter(Mandatory)]
        [hashtable]$TagsHT
    )

    $PublishingType = $NameArray[0]
    $Convention     = $NameArray[1] # Empty string if Object is not Doujinshi.
    $Artist         = $NameArray[2] # Artist is "Anthology" for Anthologies.
    $Title          = $NameArray[3]

    $DateArray = $CreationDate.split("-")
    $yyyy = $DateArray[0]; $MM = $DateArray[1]

    $TagsCSV = Select-Tags -NameArray $NameArray -TagsHt $TagsHt

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
        [string]$TargetName,

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
        TargetName = $TargetName;
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
        [string]$HashedID,

        [Parameter(Mandatory)]
        [string]$TargetName
    )


    process {

        $PublishingType = $NameArray[0]
        $Artist         = $NameArray[2]
        $Convention     = $NameArray[1]
        $Title          = $NameArray[3]
        
        return @{
            "PublishingType"    = $PublishingType
            "Artist"            = $Artist
            "Convention"        = $Convention
            "Title"             = $Title
            "TargetName"        = $TargetName
            "HashedID"             = $HashedID
        }

    }
}

Export-ModuleMember -Function New-Id,New-TargetName, Select-Collection, Convert-ListToString, Import-Tags, Select-Tags, Write-Meta, Write-Properties, New-Selector -Variable Id,TargetName, Collection, String, TagsCSV,TagsHt