function Assert-ValidString {

    Param(
        [Parameter(Mandatory, Position = 0)]
        $string
    )

    [bool]$Value

    if ($string -match "^[\w\+\-]+$" -and $string.length -le 20) {
        $Value = $true
    }
    else {
        $Value = $false
    }

    # Write-Host "Value: $Value"

    return $Value
}

function Get-TargetName {
    <#
    .SYNOPSIS
    Selects a part of a string-array and returns it.
    .DESCRIPTION
    Selects the title portion of Object-NameArray
    and returns it.
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
    .SYNOPSIS
    Converts a string-array to a single string
    and returns it.
    .DESCRIPTION
    Converts the sanitized and normalized Object-NameArray
    to a single string and returns it.
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
    .SYNOPSIS
    Selects an element from a string-array
    and returns it.
    .DESCRIPTION
    Selects the Colection from Object-NameArray
    and returns it.
    .NOTES

    [Collection] is set to:

    - the Convention's name for Doujinshi

    - the Artist's name for Manga

    - the literal string "Anthology" for Anthologies

    .OUTPUTS
    String
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
    <#
    .DESCRIPTION
    Converts a [List[String]] object to a single CSV string.
    #>
    Param(
        [Parameter(Mandatory)]
        [List[String]]$List
    )

    $String = ""                                
 
    for ($i = 0; $i -le ($List.Count - 1); $i++) {
        if ($i -le ($List.Count - 2)) {
            $String += "$($List[$i])," # < Comma
        }
        else {
            $String += "$($List[$i])"
        }
    }

    return $String
} 

function Read-TagsFromFile {
    <# 
    .SYNOPSIS
    Reads tags from a text-file, stores them in a hashtable,
    and returns it.
    #>

    Param(
        [Parameter(Mandatory)]
        [string]$Path 
    )

    <# 
        26/06/2024
        MAKE THIS FUNCTION SAFER

        e.g.
        - limit length of tags
        - remove control characters
        - etc.
    #>
    $TagsHt = @{}
    $TagSet = [HashSet[string]]::new() # < What did I do here ???

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        
        if ($line -ne "" -and -not $line.StartsWith("#")) {

            if ($line -match "^\[(?<Category>.+)\]") {

                if ($TagSet.Count -gt 0) {
                    if ((Assert-ValidString -string $Category) -eq $true){

                        try{
                            $null = $TagsHt.Add($Category, $TagSet)
                        }
                        catch{
                            Write-Debug "Category already exists."
                        }
                    }
                    else{
                        Write-Debug "Invalid category name."
                    }
                    # $TagSet = [List[string]]::new() # Changed to HashSet 06/07/2024 >>>
                    $TagSet = [HashSet[string]]::new()  
                }

                $Category = $Matches.Category # <<< This _must_ come after the if-statement. 

            }
            else {
                if ((Assert-ValidString -string $line) -eq $true) {
                    $null = $TagSet.Add(($line.Trim()))
                }
                else{
                    Write-Debug "Invalid tag name."
                }
            }
        }
    }

    # Add last Key-Value pair to $TagsHt
    if(-not $TagsHt.ContainsKey($Category)){
        $TagsHt.Add($Category, $TagSet)
    }

    #$TagsHt | Export-Clixml -Path "$($PathsLibrary.Logs)\TagsHt.xml" -Force

    return $TagsHt
}


function Import-Tags {
    <#
    .SYNOPSIS
    Imports tags from a text-file and returns them as a hashtable.
    .DESCRIPTION
    Imports tags from a text-file returns a hashtable of
    grouped tags. 
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
    # use default tags.
    if ($ImportFlag -gt 0) {

        $TagSetMeta = [System.Collections.Generic.HashSet[String]] @("english", "eng", "japanese", "digital", "censored", "uncensored", "decensored", "full color")

        $TagsHt = @{
            "Meta" = $TagSetMeta
        }
    }

    return $TagsHt
}

function Get-Tags {
    <#
    .SYNOPSIS
    Returns a CSV-string
    .DESCRIPTION
    Checks the Title and Meta portion of the NameArray of an object
    for tags that are listed in TagsHt.
    They are converted to a CSV-string and returned.
    #>

    Param(
        [array]$NameArray,
        [hashtable]$TagsHt
    )

    $TagsHt = Import-Tags

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

function Get-Meta {
    <# 
    .SYNOPSIS
    Creates and returns a hashtable.
    .DESCRIPTION
    Calculates the Meta-Properties of an object,
    stores them in a hashtable and returns it.
    #>    

    Param(
        [Parameter(Mandatory)]
        [array]$NameArray,

        [Parameter(Mandatory)]
        [string]$CreationDate
    )

    $PublishingType = $NameArray[0]
    $Convention     = $NameArray[1] # Empty string if Object is not Doujinshi.
    $Artist         = $NameArray[2] # Artist is "Anthology" for Anthologies.
    $Title          = $NameArray[3]

    $DateArray = $CreationDate.split("-")
    $yyyy = $DateArray[0]; $MM = $DateArray[1]

    $TagsCSV = Get-Tags -NameArray $NameArray

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

function Get-Properties {

    <# 
    .SYNOPSIS
    Creates and returns a hashtable
    .DESCRIPTION
    Calculates properties of an object,
    stores them in a hashtable and returns it.
    #>

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

        $TargetID = $TargetName + $NewExtension

        if($SafeCopyFlag -eq 0){

            $SourceHash = (Get-FileHash -LiteralPath $Object.FullName -Algorithm MD5).hash
        }
        else{
            $SourceHash = 0
        }

    }
    else {

        $NewExtension = ".cbz"

        $TargetID = $TargetName + $NewExtension

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
        ObjectParent  = (Split-Path -Parent $Object.FullName);
        ObjectSource  = ($Object.FullName);
        ObjectTarget  = $ObjectTarget;
        ObjectName    = $Object.Name;
        TargetName    = $TargetName;
        TargetID      = $TargetID;
        Extension     = $Ext; # Required to check if Object is file or folder in COPY
        NewExtension  = $NewExtension;
        SourceHash    = $SourceHash;
        TargetHash    = $Null
    }

}

function New-Selector {
    <# 
    .SYNOPSIS
    Creates and returns a hashtable
    .DESCRIPTION
    Creates and returns a hashtable that makes it possible to remove 
    a title from the UserLibrary during the copy part of [HSort.ps1]
    
    If an object fails to copy, it must be removed from the UserLibrary.
    .OUTPUTS
    Hashtable
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
            "HashedID"          = $HashedID
        }

    }
}

Export-ModuleMember -Function New-Id,Get-TargetName, Select-Collection, Convert-ListToString, Import-Tags, Get-Tags, Get-Meta, Get-Properties, New-Selector -Variable Id,TargetName, Collection, String, TagsCSV,TagsHt