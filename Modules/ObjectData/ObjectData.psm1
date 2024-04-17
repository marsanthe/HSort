
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
function Read-Creator {
    <# 
    .OUTPUTS
    String
    #>

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

    if ($ObjectNameArray[0] -eq "Anthology") {
        $Creator = "Anthology"
    }
    
    elseif ($ObjectNameArray[0] -eq "Manga") {
        $Creator = "$($ObjectNameArray[2])"
    }
    
    elseif ($ObjectNameArray[0] -eq "Doujinshi") {

        $Creator = "$($ObjectNameArray[1])"

    }

    return $Creator
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


function Find-Tags {
    <#
    .DESCRIPTION
        Select valid Tokens from Meta-Section of the object name.
        Return string of comma seperated tags
        to be stored in ObjectMeta.
    #>

    Param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$MetaString,

        [string]$Title,

        [hashtable]$TagsHt

    )

    $TagList = [List[string]]::new()
    
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

    if ($MetaString -ne "") {
        
        # Array
        $MetaTokenArray = $MetaString.Split(",")

        for ($i = 0; $i -le ($MetaTokenArray.length - 1); $i++) {

            # Remove all non-word characters.
            # 02/04/2024 Use string-invariants
            $Token = $MetaTokenArray[$i] -replace '[^a-zA-Z]', ''

            if ($TagsHt.Meta.Contains($Token)) {

                $TagList.Add($Token)

            }
        }
 
    }

    if ($TagList.Count -gt 0) {
        $MetaTags = Convert-ListToString -List $TagList
    }
    else{
        $MetaTags = ""
    }

    return $MetaTags

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
        [AllowEmptyString()]
        [string]$TagsCSV,

        [Parameter(Mandatory)]
        [string]$CreationDate
    )

    $PublishingType = $ObjectNameArray[0]

    # Empty string if Object is not Doujinshi.
    $Convention = $ObjectNameArray[1]

    # Artist is "Anthology" for Anthologies.
    $Artist = $ObjectNameArray[2]
    $Title = $ObjectNameArray[3]

    $DateArray = $CreationDate.split("-")
    $yyyy = $DateArray[0]
    $MM = $DateArray[1]

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
        [string]$NameNEX,

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

        $SourceHash = (Get-FileHash -LiteralPath $Object.FullName -Algorithm MD5).hash

    }
    else {
        $NewExtension = ""
        $SourceHash = (($Object | Get-ChildItem) | Measure-Object -Sum Length).sum
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
        ObjectNameNEX = $NameNEX;
        Extension     = $Ext;
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

Export-ModuleMember -Function New-ObjectName,Read-Creator,Convert-ListToString,Find-Tags,Write-Meta,Write-Properties,New-Selector -Variable Name,Creator,String,MetaTags