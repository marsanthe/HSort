using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"

function Read-Settings{

    Param(
        [Parameter(Mandatory)]
        [string]$Path # to Settings.txt
    )

    $ParsedSettings_Hst = [ordered]@{

        "ScriptVersion" = $ScriptVersion;

        "LibraryName" = "";

        "Source" = ""
        
        "Target" = "";

    }

    foreach($line in [System.IO.File]::ReadLines($Path)) {

        if($line -ne ""){

            # We need ^ as StartOfNewLine-marker, since the regex
            # would otherwise match a part of the description!
            if($line -match "(?<Argument>^LibraryName) += +(?<Value>.+)"){

                $ParsedSettings_Hst[$Matches.Argument] = $Matches.Value
            }

            elseif($line -match "(?<Argument>^Target) += +(?<Value>.+)"){

                $ParsedSettings_Hst[$Matches.Argument] = $Matches.Value
            }

            elseif($line -match "(?<Argument>^Source) += +(?<Value>.+)"){

                $ParsedSettings_Hst[$Matches.Argument] = $Matches.Value
            }

            elseif($line -match "(?<Argument>^\w+) += *$"){

                Show-Information -InformationText ("Warning",
                                                    "==========================",
                                                    " ",
                                                    "Settings.txt is incomplete.",
                                                    "Please fill out Settings.txt completely and run the script again.")

                $ParsedSettings_Hst.Add("Error","3")

                break
            }

        }
    
    }
    return $ParsedSettings_Hst
}

function Write-Settings{

    Param(
        [Parameter(Mandatory)]
        [string]$ScriptVersion,

        [Parameter(Mandatory)]
        [string]$Path # Path to Settings.txt
    )

    $SettingsTemplate_Hst = [ordered]@{

        "TxtHead"             = ("# H-Sort Settings","");

        "TxtInfo"             = ("# NOTE: DO NOT USE QUOTATION MARKS","")

        "ScriptVersion"       = $ScriptVersion;

        "TxtLibraryName"      = ("# Please enter a name for your library folder.",
                                 "# For example:  LibraryName = MyMangaLibrary",
                                 "# Allowed characters are: [a-zA-Z0-9_+-]`r`n");

        "LibraryName"         = "Hlib0";

        "TxtSource"        = ("# Please enter the path to the folder where your Manga are.",
                                 "# For example, if you have your Manga in a folder named  Manga  on your Desktop,",
                                 "# enter  C:\Users\YourUserName\Desktop\Manga`r`n");

        "Source"           = "C:\Users\Username\Desktop\Tests\Testinstanzen_reduziert";

        "TxtTarget" = ("# Please enter a path to a folder to create the library at.",
        "# For example  C:\Users\YourUserName\Desktop  will create the library on your Desktop.",
        "# All Manga will then be sorted and saved in the library-folder on your Desktop.`r`n");

        "Target"    = "C:\Users\Username\Desktop\HSortOutput"
        
    }

    # Create initial Settings.txt file
    foreach($Key in $SettingsTemplate_Hst.Keys){

        if(!($Key.StartsWith("Txt"))){
            "$Key = $($SettingsTemplate_Hst.$Key)`r`n" | Out-File -FilePath $Path -Encoding unicode -Append
        }
        else{
            for($i = 0; $i -le ($SettingsTemplate_Hst.$Key.length -1); $i++){
                $SettingsTemplate_Hst.$Key[$i] | Out-File -FilePath $Path -Encoding unicode -Append
            }
        }

    }
}

function Show-Information{

    Param(
        [Parameter(Mandatory)]
        [string[]]$InformationText
    )

    for($i = 0; $i -le ($InformationText.length -1); $i++){
        Write-Information -MessageData $InformationText[$i] -InformationAction Continue
    }

}

function Backup-Settings{
    <# 
    .NOTES
        Source: Settings
        Target: TempDir
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Target
    )

    $null =  robocopy $Source $Target *.xml  /njh /njs
}

function Restore-Settings{
    # Source: TempDir
    # Target: Settings
    
    Param(
        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [string]$Source
    )

    # Remove .XML files from "\Settings" if User aborted the script.
    Remove-Item -Path $Target -Include *.xml -Recurse -Force

    # Copy .XML files from Temp to "\Settings".
    $null =  robocopy $Source $Target *.xml  /njh /njs

    # Remove .XML files from Temp after copying.
    Remove-Item -Path $Source -Include *.xml -Recurse -Force
}

function Initialize-Script{ 

    Param(
        [Parameter(Mandatory)]
        [string]$ScriptVersion,

        [Parameter(Mandatory)]
        # Import path variables
        [System.Collections.Specialized.OrderedDictionary]$PathsProgram

    )

    $InitializeScript_ExitCode = 999

    <# 
    .Path Variables >>> from HSort.ps1

        $PathsProgram = [ordered]@{

            "Parent" = "$HOME\AppData\Roaming";
            "Base" = "$HOME\AppData\Roaming\HSort";
            "AppData" = "$HOME\AppData\Roaming\HSort\ApplicationData";
            "Libs" = "$HOME\AppData\Roaming\HSort\Libraries";
            "Settings" = "$HOME\AppData\Roaming\HSort\Settings";
            "TxtSettings" = "$HOME\AppData\Roaming\HSort\Settings\Settings.txt";
            "Copied" = "$HOME\AppData\Roaming\HSort\CopiedObjects";
            "Skipped" = "$HOME\AppData\Roaming\HSort\SkippedObjects"

        }
    #>

    ###############################################################################################################
                                                ### Script States ###
    ###############################################################################################################
    
    ### CONDITION
    ### No ProgramFolder (HSort) in \AppData\Roaming <=> 
    ### Script is run for the very first time

    if(!(Test-Path $PathsProgram.Base)){

        $InitializeScript_ExitCode = 1

        # Create folder structure in \AppData\Roaming

        foreach($Path in $PathsProgram.Keys){
            if($Path -ne "Parent"){
                if($Path -eq "TxtSettings"){
                    $null = New-Item -ItemType "file" -Path $PathsProgram.$Path
                }
                else{
                    $null = New-Item -ItemType "directory" -Path $PathsProgram.$Path  
                }  
            }
        }

        # Populate Settings-file
        Write-Settings -ScriptVersion $ScriptVersion -Path $PathsProgram.TxtSettings

        Show-Information -InformationText  ("INFORMATION",
        "==========================",
        " ",
        "This seems to be the first time you run this script.",
        "A Settings file (Settings.txt) was created here:",
        "$($PathsProgram.TxtSettings)",
        " ",
        "Please edit it to your liking and rerun the script.`n")

        Start-Sleep -Seconds 1.0
        Invoke-Item $PathsProgram.TxtSettings
        
    }

    ### CONDITION
    ### ProgramFolder (HSort) Exists

    else{

        # Validate folder structure 
        foreach($Path in $PathsPrograms.Keys){

            # Base and Parent were asserted to exist.
            if(!(Test-Path $PathsProgram.$Path)){

                if($Path -ne "TxtSettings"){
                    $null = New-Item -ItemType "directory" -Path $PathsProgram.$Path 
                } 
            }
        }
        
        # Check for Settings.txt 
        if(!(Test-Path $PathsProgram.TxtSettings)){

            $InitializeScript_ExitCode = 2
    
            Show-Information -InformationText ("Warning",
            "==========================",
            " ",
            "Settings file (Settings.txt) not found.",
            "New Settings file (Settings.txt) created here:  $($PathsProgram.Settings)",
            "Please edit it to your liking and rerun the script.")
    
            $null = New-Item -ItemType "file" -Path $PathsProgram.TxtSettings
    
            Write-Settings -ScriptVersion $ScriptVersion -Path $PathsProgram.TxtSettings
            Start-Sleep -Seconds 1.0
            Invoke-Item $PathsProgram.TxtSettings
        }
        
        ### CONDITION
        ### Settings.txt exists

        else{
    
            # Settings-file found. Waiting for User input.
            Write-Information -MessageData "Settings file found!`n`n" -InformationAction Continue
            
            # Read Settings.txt and return Hashtable
            $CurrentSettings = Read-Settings -Path $PathsProgram.TxtSettings
    
            if($CurrentSettings.Contains("Error")){
                $InitializeScript_ExitCode = 3
            }
            else{

                $CurrentParentDir = $CurrentSettings.Target
                $CurrentName = $CurrentSettings.LibraryName
                $CurrentSource = $CurrentSettings.Source

                # Display current settings.
                Show-Information -InformationText ("YOUR SETTINGS",
                "==========================",
                " ",
                "ScriptVersion: $($CurrentSettings.ScriptVersion)",
                " ",
                "LibraryName: $($CurrentSettings.LibraryName)",
                " ",
                "Target: $($CurrentSettings.Target)",
                " ",
                "Source: $($CurrentSettings.Source)",
                " ",
                "==========================",
                " ") 
    
                ### CONDITION: CurrentSettings.xml exists.
                ### IMPLIES: Settings.txt was read before at least once.
                ### SHOULD IMPLY: SettingsArchive.xml exists.

                if(Test-Path "$($PathsProgram.Settings)\CurrentSettings.xml"){

                    # In case User answers [n] in the Start-Script function.
                    Backup-Settings -Source $PathsProgram.Settings -Target $PathsProgram.Tmp
    
                    try{
                        $SettingsArchive = Import-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml"
    
                        Write-Information -MessageData "SettingsArchive.xml exists: TRUE"-InformationAction Continue
                    }
                    catch{
                        $InitializeScript_ExitCode = 4

                        Write-Information -MessageData "Error: Couldn't import SettingsArchive.xml"
                        
                    }
                    
                    ### CONDITION: Known LibraryName

                    if($SettingsArchive.ContainsKey($CurrentName)){
                        
                        Show-Information -InformationText ("SettingsArchive.xml contains CurrentName: $CurrentName`n")
    
                        ### SUB-CONDITION: Different Target.
                        ### ACTION: Special

                        if($SettingsArchive.$CurrentName.Target -ne $CurrentParentDir){
    
                            Show-Information -InformationText ("WARNING",
                            "==========================",
                             " ",
                            "A library of this name already exists in a different location.")
    
                            while($true){

                                $MoveLibrary = Read-Host "If you'd like to move the library, press [y]. If not, press [n] and change the library name."

                                if($MoveLibrary -eq "y"){
                                    
                                    $InitializeScript_ExitCode = 5

                                    Show-Information -InformationText ( "Please copy or move the library folder to the desired location.",
                                    "Then restart the script.`n")
    
                                    # Update library-information
                                    $SettingsArchive.$CurrentName.Target = $CurrentParentDir
    
                                    # Serialize updated library-information
                                    $SettingsArchive | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml" -Force
    
                                    # Serialize updated Settings.xml (Overwrite old $CurrentSettings with -Force)
                                    $CurrentSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml" -Force
                                        
                                    break
                                }
                                elseif($MoveLibrary -eq "n"){
                                    
                                    $InitializeScript_ExitCode = 6

                                    Restore-Settings -Target $PathsProgram.Settings -Source $PathsProgram.Tmp

                                    Write-Information -MessageData "Script aborted."
                                    
                                    break
                                }
                                else{
                                    Write-Information -MessageData "Please enter [y] or [n] ." -InformationAction Continue
                                }
                            }
                        }
                        
                        ### SUB-CONDITION: Same Target but different Source.
                        ### ACTION: Update Library.

                        elseif($SettingsArchive.$CurrentName.Source -ne $CurrentSource){

                            # ExitCode = -2 <=> Library is in SettingsArchive <?> LibraryContent_LibraryName.xml exists
                            $InitializeScript_ExitCode = -2
                            
                            Show-Information -InformationText ("Updating Library: $CurrentName`n", "From new Source: $CurrentSource")

                            $SettingsArchive.$CurrentName.Source = $CurrentSource
    
                            # Serialize updated library-information.
                            $SettingsArchive | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml" -Force
    
                            # Serialize updated Settings.xml
                            $CurrentSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml" -Force
                        }
    
                        ### SUB-CONDITION: Settings unchanged.
                        ### ACTION: Update Library.

                        elseif($SettingsArchive.$CurrentName.Source -eq $CurrentSource){

                            $InitializeScript_ExitCode = -2

                            Show-Information -InformationText ("Updating Library: $CurrentName`n", "From Source: $CurrentSource")

                            # The only thing that could have changed is the script version.
                            $SettingsArchive | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml" -Force
    
                            # Serialize updated Settings.xml
                            $CurrentSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml" -Force
                        }
                    }

                    ### SUB-CONDITION: New/Unknown LibraryName.
                    ### ACTION: Create new library.

                    <# 
                        The script was run before and libraries were created.
                        (This implies SettingsArchive.xml exists.)

                        Now the User wants to create a new library with
                        a different name.
                    #>
                    else{

                        if(Test-Path "$($CurrentSettings.Target)\$CurrentName"){

                            $InitializeScript_ExitCode = 7
                            
                            Show-Information ("A folder of this name already exists.",
                            "This folder is no known library folder.","Please change LibraryName in Settings.txt.")

                        }
                        else{

                            $InitializeScript_ExitCode = -1

                            # Add $UserSettings to $SettingsArchive.
                            $SettingsArchive.Add($CurrentName,$CurrentSettings)
    
                            $SettingsArchive | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml" -Force
        
                            # Serialize CurrentSettings to allow main-script access.
                            $CurrentSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml" -Force
        
                            Write-Information -MessageData "Creating new library: $CurrentName"-InformationAction Continue
                        }
                    }
                }

                ### SUB-CONDITIONS:  $CurrentSettings.xml doesn't exist ( => SettingsArchive.xml doesn't exist)
                ### ACTION: Create new Library.

                <#
                    Initial creation of SettingsArchive.xml

                    This should happen when the user runs the script for the SECOND time.
                    The first time the script is run, only the ProgramFolder structure is created,
                    including Settings.txt.
                    
                    CurrentSettings.xml doesn't exist yet -> Create and Serialize.
                    SettingsArchive.xml does NOT exist yet either -> Create and Serialize.

                #>
                else{
                    if(Test-Path "$($CurrentSettings.Target)\$CurrentName"){

                        Show-Information ("A folder of this name already exists.",
                        "This folder is no known library folder.","Please change LibraryName in Settings.txt.")

                        $InitializeScript_ExitCode = 7
                    }
                    else{

                        $InitializeScript_ExitCode = 0
    
                        $SettingsArchive = @{}
                        $SettingsArchive.Add($CurrentName,$CurrentSettings)
                        $SettingsArchive | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsArchive.xml" 
        
                        # Serialize $Settings_Hst to allow main-script access.
                        $CurrentSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\CurrentSettings.xml" -Force
        
                        Write-Information -MessageData "SettingsArchive.xml exists: FALSE`n" -InformationAction Continue
                        Write-Information -MessageData "Creating: SettingsArchive.xml`n" -InformationAction Continue
                        Write-Information -MessageData "Creating New Library File: $CurrentName`n" -InformationAction Continue

                    }
                }
            }
        }
    }

    return $InitializeScript_ExitCode
}

function Confirm-Settings{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$CurrentSettings
    )

    $ConfirmSettings_ExitCode = 0

    if($CurrentSettings.LibraryName -match "[\w\+\-]+"){
        Write-Information -MessageData "LibraryName $($CurrentSettings.LibraryName): OK`n" -InformationAction Continue

        if(Test-Path -Path $CurrentSettings.Target){
            Write-Information -MessageData "Target $($CurrentSettings.Target): OK`n" -InformationAction Continue

            if(Test-Path -Path $CurrentSettings.Source){
                Write-Information -MessageData "LibrarySource $($CurrentSettings.Source): OK`n" -InformationAction Continue
            }
            else{

                $ConfirmSettings_ExitCode = 1

                Show-Information -InformationText("Error",
                "==========================",
                " ",
                "Invalid Source",
                "The directory $($CurrentSettings.Source)",
                "doesn't exist.",
                "Please change Source in Settings.txt.")
            }
        }
        else{

            $ConfirmSettings_ExitCode = 1

            Show-Information -InformationText ("Error",
            "==========================",
            " ",
            "Invalid Target",
            "The directory $($CurrentSettings.Target)",
            "doesn't exist.",
            "Please change Target in Settings.txt.")
        }
    }
    else{

        $ConfirmSettings_ExitCode = 1

        Show-Information -InformationText("Error",
            "==========================",
            " ",
            "Bad LibraryName",
            "LibraryName contains illegal characters.",
            "Allowed characters are: [a-zA-Z0-9_+-]",
            "Please change LibraryName in Settings.txt."
            )

    }

    return $ConfirmSettings_ExitCode
}

function Start-Script{

    $StartScript_ExitCode = 999

    while($true){

        $ConfirmStartScript = Read-Host  "Press [y] to confirm and start the script. Press [n] to abort."

        if($ConfirmStartScript -eq "y"){
            Write-Information -MessageData "Starting script. Please wait..." -InformationAction Continue
            $StartScript_ExitCode = 0
            break
        }
        elseif($ConfirmStartScript -eq "n"){
            Write-Information -MessageData "Script aborted." -InformationAction Continue
            $StartScript_ExitCode = 1
            break
        }
        else{
            Write-Information -MessageData "Please enter [y] to start the script or [n] to abort." -InformationAction Continue
        }
    }
    return $StartScript_ExitCode
}

Export-ModuleMember -Function Initialize-Script,Confirm-Settings,Start-Script,Restore-Settings,Resume-OnError -Variable InitializeScript_ExitCode,ConfirmSettings_ExitCode,StartScript_ExitCode,ResumeExitCode
