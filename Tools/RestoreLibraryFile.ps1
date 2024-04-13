using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"


function Restore-LibraryFile{

    [CmdletBinding()]

    Param(
        [Parameter(Mandatory)]
        [string]$LibraryPath,

        [Parameter(Mandatory)]
        [string]$TargetPath
    )

    $RestoredLibrary = @{}

    $LibraryPath -match ".+\\(?<LibraryName>.+)"
    $LibraryName = $Matches.LibraryName

    $Artists = Get-ChildItem -Path "$LibraryPath\Artists"

    foreach($Artist in $Artists){

        # Artist is directory object
        $ArtistName = $Artist.Name
        $RestoredLibrary[$ArtistName] = @{}
        $Titles = Get-ChildItem -LiteralPath $Artist.FullName

        <# 
            TitleFile is .cbr or .cbz
            since HSort only creates .cbz or .cbr.
        #>
        foreach($TitleFile in $Titles){
            <# 
                $TitleFile is folder- or file-object,
                so $TitleFile.Name <=> ObjectName
                (See $ObjectProperties in HSort)
            #>

            $TitleFile.Name -match "(\(.+\))* *(?<Title>[^(|)]+) *(\(.+\))*(\..*)"
            
            $Title = $Matches.Title
            $Title = $Title.trim()

            $RestoredLibrary.$ArtistName[$Title] = @{
                ObjectSourcePath = "missing";
                Extension = "missing";
                FirstDiscovered = "missing"
                Duplicates = [List[string]]::new()  
            }
        }
    }

    $Artists = @()

    ##################################################################

    $Conventions = Get-ChildItem -Path "$LibraryPath\Conventions"
    
    foreach($Convention in $Conventions){

        $ConventionName = $Convention.Name
        $RestoredLibrary[$ConventionName] = @{}
        $Titles = Get-ChildItem -LiteralPath $Convention.FullName

        foreach($TitleFile in $Titles){

            $TitleFile.Name -match "(\(.+\))* *(?<Title>[^(|)]+) *(\(.+\))*(\..*)"

            $Title = $Matches.Title

            $Title = $Title.trim()

            $RestoredLibrary.$ConventionName[$Title] = @{
                ObjectSourcePath = "missing";
                Extension = "missing";
                FirstDiscovered = "missing"
                Duplicates = [List[string]]::new()  
            }
        }
    }

    $Conventions = @()

    ##################################################################

    $Anthologies = Get-ChildItem -Path "$LibraryPath\Anthologies"

    # If folder not empty.
    if($Anthologies.length -gt 0){

        $RestoredLibrary["Anthologies"] = @{}

        foreach($AnthologyFile in $Anthologies){

            $AnthologyFile.Name -match "(\(.+\))* *(?<Title>[^(|)]+) *(\(.+\))*(\..*)"

            $Anthology = $Matches.Title

            $Anthology = $Anthology.trim()

            $RestoredLibrary.Anthologies[$Anthology] = @{
                ObjectSourcePath = "missing";
                Extension = "missing";
                FirstDiscovered = "missing"
                Duplicates = [List[string]]::new()  
            }
        
        }
    }
    
    $Anthologies = @()

    $RestoredLibrary | Export-Clixml -Path "$TargetPath\Library $LibraryName.xml"
}