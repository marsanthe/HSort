
function Invoke-ExcludedObjectsHandler {
    <#
    .SYNOPSIS 
    Adds [ExcludedObjects] (from current session) to ExcludedObjects_LibraryName
    Then exports ExcludedObjects_Library as [ExcludedObjects $LibraryName.xml] 
    
    [ExcludedObjects $LibraryName.xml] is required by <MoveExcludedObjects.ps1> in Tools.
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$ExcludedObjects_Session,

        [Parameter(Mandatory)]
        [string]$LibraryFiles,

        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [string]$LibSrc,

        [Parameter(Mandatory)]
        [hashtable]$SrcDirTree

    )

    # Check if a "Excluded" file for this library already exists.
    if (Test-Path "$($LibraryFiles)\$LibraryName\Excluded $LibraryName.xml") {
    
        $ExcludedObjects_Library = Import-Clixml -Path "$($LibraryFiles)\$LibraryName\Excluded $LibraryName.xml"
    
        # Check if this "Excluded" file has a section for the library's current source.
        if ($ExcludedObjects_Library.ContainsKey($LibSrc)) {

            # Update (overwrite) SrcDirTree with SrcDirTree of the source in its current state.
            $ExcludedObjects_Library.$LibSrc.SrcDirTree = $SrcDirTree

            $Excluded = $ExcludedObjects_Library.$LibSrc.Excluded
            
            foreach ($Object in $Excluded.Keys) {

                if( -not $Excluded.ContainsKey($Object) ){

                    $Excluded.Add($Object, $Excluded.$Object)
                }

            }
        }
        else { # Create a section for Excluded files from the library's current source.
            

        <# 
        .NOTES

        [SrcDirTree] in the hashtable below is currently not used.
        It is left in the hashtable for testing purposes only.

        Explanation:
        Exporting [SrcDirTree] - which is an ordered dictionary - is currently pointless,
        since it loses it's ordered state when we import [ExcludedObjects_Library] back
        via <Import-Clixml>.
        This is a powershell-bug only fixed in a later version of PS 7.X
        This script aims to be compatible with PS 5.1 upwards.

        To handle this, we export SrcDirTree in [GetObjects.ps1] as a text-file,
        which is later parsed by [MoveExcludedObjects.ps1] to recreate
        [SrcDirTree] as an ordered hashtable.

        #>
            $ExcludedObjects_Library[$LibSrc] = @{
                SrcDirTree = $SrcDirTree;
                Excluded    = $ExcludedObjects_Session
            }
        }
    
        $ExcludedObjects_Library | Export-Clixml -Path "$($LibraryFiles)\$LibraryName\Excluded $LibraryName.xml" -Force
    }
    else { # Create a Excluded file for this library.
        
        $ExcludedObjects_Library = @{}

        $ExcludedObjects_Library[$LibSrc] = @{
                SrcDirTree = $SrcDirTree;
                Excluded = $ExcludedObjects_Session
        }
        
        $ExcludedObjects_Library | Export-Clixml -Path "$($LibraryFiles)\$LibraryName\Excluded $LibraryName.xml" -Force
    }
}

Export-ModuleMember -Function Invoke-ExcludedObjectsHandler