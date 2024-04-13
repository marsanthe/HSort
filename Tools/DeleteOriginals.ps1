
function Remove-SourceObjects{

    [CmdletBinding()]

    Param(
        [Parameter(Mandatory)]
        [string]$LibraryName,

        [string]$CopiedObjectsDir = "$HOME\AppData\Roaming\HSort\CopiedObjects"
    )

    Process{

        $Timestamp = Get-Date -Format "dd_MM_yyyy HH_mm_ss"

        $StringArray = ("WARNING",
        "==========================",
        " ",
        "This script will DELETE all Manga",
        "that were successfully copied to your Library",
        "from their original location.",
        "Are you sure you want to proceed?")

        for($i = 0; $i -le ($StringArray.length -1); $i++){
            Write-Information -MessageData $StringArray[$i] -InformationAction Continue
        }

        While($true){

            $ConfirmDeletion = Read-Host "Enter [y] to continue or [n] to abort."

            if($ConfirmDeletion -eq "y"){
                Write-Output "Deleting files..."
                break
            }
            elseif($ConfirmDeletion -eq "n"){
                Write-Output "Script aborted. No files were deleted."
                exit
            }
            else{
                Write-Output "Please enter [y] for deletion or [n] to abort."
            }
        }

        try{
            $CopiedObjects = Import-Clixml -Path "$CopiedObjectsDir\CopiedObjects $LibraryName.xml"
        }
        catch{
            Write-Information -MessageData "CopiedObjects.xml not found.
            Exiting." -InformationAction Continue
        }
        
        foreach($Object in $CopiedObjects.Keys){

            try{
                Write-Output "Deleting $($CopiedObjects.$Object.ObjectName)"
                Remove-Item -LiteralPath $CopiedObjects.$Object.ObjectSource -Recurse -Force -WhatIf
                Out-File -FilePath "$CopiedObjectsDir\CopiedObjects $Timestamp.txt"
            }
            catch{
                # Just to be absolutely on the save side.
                if(!(Test-Path -LiteralPath $CopiedObjects.$Object.ObjectSource)){

                    "$($CopiedObjects.$Object.ObjectSource) - OBJECT NOT FOUND" |
                    Out-File -FilePath "$CopiedObjectsDir\UncopiedObjects $Timestamp.txt" -Encoding unicode -Append
                }
                else{

                    "$($CopiedObjects.$Object.ObjectSource) - UNKONWN ERROR" |
                    Out-File -FilePath "$CopiedObjectsDir\UncopiedObjects $Timestamp.txt" -Encoding unicode -Append
                }
            }   
        }
    }
}

