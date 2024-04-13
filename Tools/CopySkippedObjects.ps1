function Copy-SkippedObjects{
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
            To copy the files that were skipped during the creation of this library.
        ParentDir
            Directory to create the folder containing the skipped files in. 
    #>

    [CmdletBinding()]

    Param(
        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [string]$ParentDir
    )

    Begin{
        [string]$SourceDir = "$HOME\AppData\Roaming\HSort\SkippedObjects"
    }

    Process{

        "This script will move all Manga that were skipped
        to a folder of your choice.`n" | Write-Output
        
        try{
            $SkippedXML = Import-Clixml -Path "$SourceDir\Skipped $LibraryName.xml"
        }
        catch{
            Write-Output "Skipped $Libraryname.xml not found. Exiting"
            exit
        }
        
        $PathsSkipped = [ordered]@{
            "Base" = "$ParentDir\Skipped $LibraryName";
            "NoMatch" = "$ParentDir\Skipped $LibraryName\NoMatch";
            "Duplicates" = "$ParentDir\Skipped $LibraryName\Duplicates";
            "WrongExtension" = "$ParentDir\Skipped $LibraryName\WrongExtension";
            "XMLinsertionError" = "$ParentDir\Skipped $LibraryName\XMLinsertionError";
            "HashMismatch" = "$ParentDir\Skipped $LibraryName\HashMismatch";
            "UnsupportedOrMixedContent" = "$ParentDir\Skipped $LibraryName\UnsupportedOrMixedContent"
        }

        foreach($Path in $PathsSkipped.Keys){
            if(!(Test-Path $PathsSkipped.$Path)){
                $null = New-Item -ItemType "directory" -Path $PathsSkipped.$Path
            }
        }


        foreach($Object in $SkippedXML.Keys){
            $Source = $SkippedXML.$Object.Path
            $Name = $SkippedXML.$Object.ObjectName
            $ParentDir = $SkippedXML.$Object.ObjectParent

            if($SkippedXML.$Object.Reason -eq "Duplicate"){
                $H = $Source.GetHashCode()
            }

            if($SkippedXML.$Object.Extension -eq "Folder") {

                switch(($SkippedXML.$Object.Reason)){

                    "NoMatch" {
                                try{$null = robocopy $Source "$($PathsSkipped.NoMatch)\$Name" /E /DCOPY:DAT /move}
                                catch {Write-Output "Copy error"};
                                break }
                    

                    "Duplicate" {
                        try{$null = robocopy $Source "$($PathsSkipped.Duplicates)\$Name $H" /E /DCOPY:DAT}
                        catch {Write-Output "Copy error"};
                        break }

                    "HashMismatch" {
                                try{$null = robocopy $Source "$($PathsSkipped.HashMismatch)\$Name" /E /DCOPY:DAT /move}
                                catch {Write-Output "Copy error"};
                                break }

                    # Defines in GetObjects.psm1
                    "UnsupportedOrMixedContent" { Write-Information -MessageData "$Name" -InformationAction Continue
                                try{$null = robocopy $Source "$($PathsSkipped.UnsupportedOrMixedContent)\$Name" /E /DCOPY:DAT /move}
                                catch {Write-Output "Copy error"};
                                break }
                }
            }
            else{
                switch(($SkippedXML.$Object.Reason)){

                    "NoMatch" {try{$null = robocopy $ParentDir $PathsSkipped.NoMatch $Name /move}
                    catch {Write-Output "Copy error"};
                    break }

                    "WrongExtension" {try{$null = robocopy $ParentDir $PathsSkipped.WrongExtension $Name /move}
                    catch {Write-Output "Copy error"};
                    break }

                    "Duplicate" {try{$null = robocopy $ParentDir $PathsSkipped.Duplicates $Name /move;
                        Rename-Item -LiteralPath "$($PathsSkipped.Duplicates)\$Name" -NewName "$H $Name"}
                    catch {Write-Output "Copy error"};
                    break }

                    "XMLinsertionError" {try{$null = robocopy $ParentDir $PathsSkipped.XMLinsertionError $Name /move}
                    catch {Write-Output "Copy error"};
                    break }

                    "HashMismatch" {try{$null = robocopy $ParentDir $PathsSkipped.HashMismatch $Name /move}
                    catch {Write-Output "Copy error"};
                    break }

                }
            }
        }
    }
}