
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"
$DebugPreference = "Continue"

#===========================================================================
# Public functions
#===========================================================================
function Write-Head($text) {
    <#
    .SYNOPSIS
    Prints a text block surrounded by a section divider for enhanced output readability.

    .DESCRIPTION
    This function takes a string input and prints it to the console, surrounded by a section divider made of hash characters.
    It is designed to enhance the readability of console output.

    .NOTES
    Inspired by <Write-Section> from Wintuil on Github
    #>
    Write-Information -MessageData " " -InformationAction Continue 
    Write-Information -MessageData ("=" * ($text.Length + 4)) -InformationAction Continue 
    Write-Information -MessageData "= $text =" -InformationAction Continue 
    Write-Information -MessageData ("=" * ($text.Length + 4)) -InformationAction Continue 
    Write-Information -MessageData " " -InformationAction Continue 
}

function Write-Paragraph {
    <#
    .SYNOPSIS
    Displays one or more lines of text.
    #>
    Param(
        [Parameter(Mandatory)]
        #[AllowEmptyString()]
        [string[]]$InformationArray
    )

    if($InformationArray){
        for ($i = 0; $i -le ($InformationArray.length - 1); $i++) {
            Write-Information -MessageData $InformationArray[$i] -InformationAction Continue
        }
    }
    else{
        Write-Information -MessageData " " -InformationAction Continue
    }

}

function Get-UserInput {
    <# 
    .DESCRIPTION
    A function based on Read-Host.
    The user is asked to answer a Yes/No question.
    The function returns "y"  for a "yes"-response
    and "n" for a "no"-response.

    The function expects a hashtable of this format as argument:

    $UserInput = Get-UserInput -Dialog @{
        "Question" = "The question";
        "YesResponse" = "Message after "yes"-response";
        "NoResponse" = Message after "no"-response
    }
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$Dialog
    )

    [string]$UserResponse

    if ($Dialog.Intro) {
        Write-Paragraph -InformationArray $Dialog.Intro
    } 

    while ($true) {

        $Answer = Read-Host $Dialog.Question

        if ($Answer -eq "y") {
            Write-Information -MessageData $Dialog.YesResponse -InformationAction Continue
            $UserResponse = "y"
            break
        }
        elseif ($Answer -eq "n") {
            Write-Information -MessageData $Dialog.NoResponse -InformationAction Continue
            $UserResponse = "n"
            break
        }
        else {
            Write-Information -MessageData "`nPlease enter [y] or [n].`n" -InformationAction Continue
        }

    }

    return $UserResponse
}

#===========================================================================
# Private functions
#===========================================================================

function Read-Settings{
    <# 
    .SYNOPSIS
    Parses Settings.txt from $HOME/AppData/Roaming/HSort
    to extract the settings for this invocation of <Hsort.ps1>
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$Path # to Settings.txt
    )

    Write-Debug "[Read-Settings]"

    $ParsedSettings = [ordered]@{

        "ScriptVersion" = $ScriptVersion;

        "SafeCopy"      = "";

        "LibraryName" = "";

        "Source" = ""
        
        "Target" = "";

    }
    
    $lc = 0
    foreach($line in [System.IO.File]::ReadLines($Path)) {

        $lc++

        if ($line -ne "" -and -not $line.StartsWith("#")) {

            # We need ^ as StartOfNewLine-marker, since the regex
            # would otherwise match a part of the description!
            if($line -match "(?<Argument>^LibraryName) += +(?<Value>.+)"){

                $ParsedSettings[$Matches.Argument] = $Matches.Value
            }

            elseif ($line -match "(?<Argument>^SafeCopy) += +(?<Value>.+)") {

                $ParsedSettings[$Matches.Argument] = $Matches.Value
            }

            elseif($line -match "(?<Argument>^Target) += +(?<Value>.+)"){

                $ParsedSettings[$Matches.Argument] = $Matches.Value
            }

            elseif($line -match "(?<Argument>^Source) += +(?<Value>.+)"){

                $ParsedSettings[$Matches.Argument] = $Matches.Value
            }

        }

        if($lc -gt 50){
            throw "Malformatted Settings.txt`nLine count is greater than 50.`nExiting..."
        }
    
    }

    return $ParsedSettings
}
#"C:\Users\M. Thedja\AppData\Roaming\HSort\Temp"
function Write-Settings{
    <# 
    .SYNOPSIS
    Create and populate initial Settings.txt file.

    .DESCRIPTION
    Creates Settings.txt in $HOME/AppData/Roaming/HSort
    and writes the default settings/examples to it. 
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$Path # Path to Settings.txt
    )

    Write-Debug "[Write-Settings]"

    $SettingsTemplate_Hst = [ordered]@{

        "TxtHead"             = ("# H-Sort Settings","");

        "ScriptVersion"       = $ScriptVersion;

        "TxtSafeCopy"         = ("# Calculate FileHash/FolderSize.",
                                 "# before and after copying to detect copy errors.",
                                 "# Recommended setting: True `r`n");

        "SafeCopy"            = "True";

        "TxtLibraryName"      = ("# Please enter a name for your library folder.",
                                 "# For example:  LibraryName = MyMangaLibrary",
                                 "Spaces are not allowed.",
                                 "# Allowed characters are: [a-zA-Z0-9_+-]`r`n");

        "LibraryName"         = "YourLibraryName";

        "TxtSource"           = ("# Please enter the path to a source folder.",
                                 "# For example, if you have your Manga in a folder named [Manga] on your Desktop,",
                                 "# enter:  C:\Users\YourUserName\Desktop\Manga`r`n");

        "Source"              = "C:\Users\YourUserName\Desktop\Manga";

        "TxtTarget"           = ("# Where to create the library.",
                                 "# For example:  C:\Users\YourUserName\Desktop  will create the library on your Desktop.`r`n");

        "Target"              = "C:\Users\YourUserName\Desktop\Manga"
        
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


function Backup-Settings{
    <#
    .DESCRIPTION
    Required in case the user decides to end the script in <Start-Script>
    Create copy of [ActiveSettings.xml] and [SettingsHistory.xml] in $HOME/AppData/Roaming/HSort/Tmp 
    We need a copy of [SettingsHistory.xml] too, since [ActiveSettings.xml] will be written to it later.

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
        
    Write-Debug "[Backup-Settings]"
    
    $null =  robocopy $Source $Target *.xml  /njh /njs
}

function Restore-Settings{
    <#
    .SYNOPSIS
    Restore [ActiveSettings.xml] and [SettingsHistory.xml] to $HOME/AppData/Roaming/HSort/Settings 

    .DESCRIPTION
    Required in case the user decides to end the script in <Start-Script>
    Remove current [ActiveSettings.xml] and current [SettingsHistory.xml] from /HSort/Settings
    Copy backup [ActiveSettings.xml] and [SettingsHistory.xml] from /HSort/Tmp to /HSort/Settings
    Remove backup [ActiveSettings.xml] and [SettingsHistory.xml] from /HSort/Tmp

    .NOTES 
    Source: TempDir
    Target: Settings
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$Target,
        
        [Parameter(Mandatory)]
        [string]$Source
    )
        
    Write-Debug "[Restore-Settings]"

    # Remove .XML files from "\Settings" if User aborted the script.
    Remove-Item -Path $Target -Include *.xml -Recurse -Force

    # Copy .XML files from Temp to "\Settings".
    $null =  robocopy $Source $Target *.xml  /njh /njs

    # Remove .XML files from Temp after copying.
    Remove-Item -Path $Source -Include *.xml -Recurse -Force
}

function New-ProgramFolder{
    <# 
    .SYNOPSIS
    Create HSort-folder in $HOME\AppData\Roaming\

    .DESCRIPTION
    Create initial HSort-folder in $HOME\AppData\Roaming\

    = Initial HSort-folder layout =

    $HOME\AppData\Roaming\HSort\
    --- LibraryFiles\
    --- Settings\
    --- Temp\
    --- Settings.txt 
    #>
    Param(
        [Parameter(Mandatory)]
        [System.Collections.Specialized.OrderedDictionary]$PathsProgram

    )

    Write-Debug "[New-ProgramFolder]"

    # Create empty program folders/files in $HOME\AppData\Roaming\HSort
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
}

function Assert-ProgramFolder{
    <# 
    .DESCRIPTION
    Check if HSort-folder in $HOME\AppData\Roaming\ is complete.
    If not, create missing folders/files.
    #>

    Write-Debug "[Assert-ProgramFolder]"

    foreach($Path in $PathsPrograms.Keys){
        # Restore missing folders
        if(! (Test-Path $PathsProgram.$Path) ){

            Write-Information -MessageData "[WARNING] $Path missing" -InformationAction Continue

            if($Path -ne "TxtSettings"){
                Write-Information -MessageData "[CREATING] $Path " -InformationAction Continue
                $null = New-Item -ItemType "directory" -Path $PathsProgram.$Path 
            }
        }
    }
}

function Test-SettingsTxt{
    <# 
    .SYNOPSIS
    Check if Settings.txt in $HOME\AppData\Roaming\HSort\ exists.

    .DESCRIPTION 
    Check if Settings.txt in $HOME\AppData\Roaming\HSort\ exists.
    
    Since we check its existence only after we've checked that the HSort-folder exists,
    we can assume that this is not the first time the user runs the script. 
    So it definitely _should_ exist at this point, if it wasn't deleted accidentially.
    
    If it does not exist, we display a warning and create a new default Settings.txt file.
    #>

    Write-Debug "[Test-SettingsTxt]"

    if(!(Test-Path $PathsProgram.TxtSettings)){

        $Init_ExitCode = 2

        Write-Head("WARNING")
        Write-Paragraph -InformationArray (
        "Settings file (Settings.txt) not found.",
        "New Settings file (Settings.txt) created here:  $($PathsProgram.Settings)",
        "Please edit it to your liking and run the script again.",
        " ")

        $null = New-Item -ItemType "file" -Path $PathsProgram.TxtSettings

        Write-Settings -Path $PathsProgram.TxtSettings

    }
    else{
        Write-Debug "Settings file exists."
        $Init_ExitCode = 0
    }

    return $Init_ExitCode

}

function Assert-SettingsValid{
    <# 
    .SYNOPSIS
    We assert that Settings.txt contains valid settings for <HSort.ps1> 
    #>

    Param(

        [Parameter(Mandatory)]
        [hashtable]$ActiveSettings
    )

    Write-Debug "[Assert-SettingsValid]"

    $Init_ExitCode = 0

    Write-Debug "[Checking Settings.txt]" 
    
    foreach($Key in $CurrentSettgins.Keys){

        if($ActiveSettings.Key -eq ""){

            Write-Head("WARNING")
            Write-Paragraph -InformationArray (" ",
                "Settings.txt is incomplete.",
                "Please fill out Settings.txt completely and run the script again.")

            $Init_ExitCode = 3
            break
        }
    }

    if($Init_ExitCode -eq 0){
        
        # Check if only allowed caracters are used: same as [a-zA-Z0-9_+-]
        # StartOfLine "^" and EndOfLine "$" markers are required.
        if ($ActiveSettings.LibraryName -match "^[\w\+\-]+$") {
    
            Write-Debug "LibraryName: [OK]"
    
            if (Test-Path -Path $ActiveSettings.Target) {
    
                Write-Debug "Target: [OK]"
    
                if (Test-Path -Path $ActiveSettings.Source) {
    
                    if ($ActiveSettings.Target -ne $ActiveSettings.Source) {
    
                        Write-Debug "Source: [OK]"
    
                    }
                    else {

                        $Init_ExitCode = 1
    
                        Write-Head("ERROR")
                        Write-Paragraph -InformationArray(
                            "Source and target must be different directories.",
                            " ")
                    }
                }
                else {
    
                    $Init_ExitCode = 1
    
                    Write-Head("ERROR")
                    Write-Paragraph -InformationArray(
                        "Invalid Source",
                        "This directory doesn't exist:",
                        "$($ActiveSettings.Source)",
                        "Please change [Source] in Settings.txt.",
                        " ")
                }
            }
            else {
    
                $Init_ExitCode = 1
    
                Write-Head("ERROR")
                Write-Paragraph -InformationArray (
                    "Invalid Target",
                    "This directory doesn't exist.",
                    "$($ActiveSettings.Target)",
                    "Please change [Target] in Settings.txt.",
                    " ")
            }
        }
        else {
    
            $Init_ExitCode = 1
    
            Write-Head("ERROR")
            Write-Paragraph -InformationArray(
                "Bad LibraryName",
                "LibraryName contains illegal characters.",
                "Allowed characters are: [a-zA-Z0-9_+-]",
                "Please change LibraryName in Settings.txt.",
                " ")
        }
    }

    return $Init_ExitCode
}

function Test-ProgramFolder{
    <# 
    .SYNOPSIS
    We check if the HSort-folder exists in $HOME\AppData\Roaming
    If it doesn't we can assume that this is the first time
    the user runs the script and call <New-ProgramFolder> to create it.

    .NOTES
    $PathsProgram is defined as global in <HSort.ps1> 
    #>

    Write-Debug "[Test-ProgramFolder]"

    $Init_ExitCode = 0

    if ( !(Test-Path $PathsProgram.Base) ) {

        $Init_ExitCode = 1

        New-ProgramFolder -PathsProgram $PathsProgram

        # Populate Settings-file
        Write-Settings -Path $PathsProgram.TxtSettings

        Write-Head("INFORMATION")
        Write-Paragraph -InformationArray  (
            "This seems to be the first time you run this script.",
            "A Settings file (Settings.txt) was created here:",
            "$($PathsProgram.TxtSettings)",
            " ",
            "Please edit it to your liking and run the script again.",
            " ")

        Start-Sleep -Seconds 1.0
        #Invoke-Item $PathsProgram.TxtSettings

    }

    return $Init_ExitCode
}

function Get-DriveSerial{
    <#
    .DESCRIPTION
    Get the serial number of the drive that contains the source folder.
    The serial number allows us to find the right library
    even when the drive letter of the drive it is stored on changes.
    
    This might happen if the library is stored on an external drive
    that gets disconnected and later reconnected.
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )

    Write-Debug "[Get-DriveSerial]"

    try{
        $DriveLetter = (Get-item -LiteralPath $SourcePath).PSDrive.Name
        $Drive = get-partition -DriveLetter $DriveLetter | get-disk
    }
    catch{
        throw "Drive is not accessible."
    }
    $SerialNumber = $Drive.SerialNumber

    return $SerialNumber
}

function Assert-LibrariesAccessible{
    <# 
    .SYNOPSIS
    Assert that the library-folder is accessible.

    .DESCRIPTION
    If the user wants to update an existsing library,
    we have to make sure that the library-folder is accessible.
    #>

    Param(
        [Parameter(Mandatory)]
        [hashtable]$SettingsHistory,

        [Parameter(Mandatory)]
        [string]$ActiveName
    )

    Write-Debug "[Assert-LibrariesAccessible]"
    $NewDriveLetter = ""

    if( Test-Path $SettingsHistory.$ActiveName.Source){
        Write-Debug "Library found"
    }
    else{
        $SrcDriveSerial = $SettingsHistory.$ActiveName.SrcDriveSerial

        $Volumes = Get-Volume

        foreach ($Volume in $Volumes) {

            $DriveLetter = ($Volume.DriveLetter)

            # Some (hidden) volumes don't have dirve letters
            if ($DriveLetter) {

                $Disk = get-partition -DriveLetter $DriveLetter | get-disk
                $SerialNumber = $Disk.SerialNumber

                if($SrcDriveSerial -eq $SerialNumber){

                    Write-Debug "$LibName found."
                    Write-Debug "New drive letter [$DriveLetter]"

                    $NewDriveLetter = $DriveLetter
                    
                    $NewSource = $NewDriveLetter + ($SettingsHistory.$ActiveName.Source).Substring(1) 

                    $SettingsHistory.$ActiveName.Source = $NewSource
                    
                }
                else{

                    Write-Debug "Library folder $LibName not found."
                    Write-Debug "The library folder doesn't exist in the specified location."
                    Write-Debug "It might have been moved or deleted."

                    Throw "Library folder $LibName not found."
                }
            }
        }
    }
}

function Move-Library{
    <# 
    .SYNOPSIS
    Allows the user to move an existsing library-folder to a new location.
    This requires updating [ActiveSettings.xml] and [SettingsHistory.xml]
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SettingsHistory,

        [Parameter(Mandatory)]
        [hashtable]$ActiveSettings
    )

    Write-Debug "[Move-Library]"

    Write-Head("WARNING")
    Write-Paragraph -InformationArray (
    "A library of this name already exists in a different location.",
    "Creating two libraries with the same name is not possible.",
    " ")

    while($true){

        $MoveLibrary = Read-Host "If you'd like to move the library, press [y]. If not, press [n] and change the library name."

        if($MoveLibrary -eq "y"){
            
            $Init_ExitCode = 5

            Write-Paragraph -InformationArray ( "Please copy or move the library folder to the desired location.",
            "Then run the script again.`n")

            # Update library-information
            $SettingsHistory.$ActiveName.Target = $ActiveTarget

            # Serialize updated library-information
            $SettingsHistory | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsHistory.xml" -Force

            # Serialize updated Settings.xml (Overwrite old $ActiveSettings with -Force)
            $ActiveSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\ActiveSettings.xml" -Force
                
            break
        }
        elseif($MoveLibrary -eq "n"){
            
            $Init_ExitCode = 6

            Restore-Settings -Target $PathsProgram.Settings -Source $PathsProgram.Tmp
            Write-Information -MessageData "Exiting..." -InformationAction Continue
            
            break
        }
        else{
            Write-Information -MessageData "Please enter [y] or [n] ." -InformationAction Continue
        }
    }

    return $Init_ExitCode
}

function Update-Library{
    <# 
    .SYNOPSIS
    Tells <HSort.ps1> to update an existing library.
    Either from the same source-folder or a new source-folder.
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SettingsHistory,

        [Parameter(Mandatory)]
        [hashtable]$ActiveSettings,

        [Parameter(Mandatory)]
        [string]$ActiveSource
    )

    Write-Debug "[Update-Library]"

    if($SettingsHistory.$ActiveName.Source -ne $ActiveSource){

        # ExitCode = -2 <=> Library is in SettingsHistory <?> UserLibrary_[LibraryName].xml exists
        $Init_ExitCode = -2
        
        Write-Paragraph -InformationArray ("Updating Library: $ActiveName`n", "From new Source: $ActiveSource")

        $SettingsHistory.$ActiveName.Source = $ActiveSource

        # Serialize updated library-information.
        $SettingsHistory | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsHistory.xml" -Force

        # Serialize updated Settings.xml
        $ActiveSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\ActiveSettings.xml" -Force
    }

    ### SUB-CONDITION: Settings unchanged.
    ### ACTION: Update Library.

    elseif($SettingsHistory.$ActiveName.Source -eq $ActiveSource){

        $Init_ExitCode = -3
        Write-Paragraph -InformationArray ("Updating Library: $ActiveName`n", "From Source: $ActiveSource")

    }

    return $Init_ExitCode
}

function New-Library{
    <# 
    .SYNOPSIS
    Tells <HSoprt.ps1> to create a new library.
    Creates a new folder in $HOME\AppData\Roaming\HSort\LibraryFiles
    named after the library (library name is [$ActiveName])
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SettingsHistory,

        [Parameter(Mandatory)]
        [hashtable]$ActiveSettings,

        [Parameter(Mandatory)]
        [string]$ActiveName
    )

    Write-Debug "[New-Library]"

    # Assert that no other "LibraryName" folder exists in the TargetDir
    if(-not (Test-Path "$($ActiveSettings.Target)\$ActiveName")){

        $Init_ExitCode = -1

        # Add $UserSettings to $SettingsHistory.
        $SettingsHistory.Add($ActiveName, $ActiveSettings)

        # Serialize ActiveSettings to allow main-script access.
        $ActiveSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\ActiveSettings.xml" -Force
        $SettingsHistory | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsHistory.xml" -Force

        Write-Information -MessageData "Creating new library: $ActiveName"-InformationAction Continue

        ### NEW 23/06/2024
        ### Create folder for library specific data in Roaming\HSort\LibraryFiles
        $null = New-Item -Path "$($PathsProgram.LibFiles)" -ItemType Directory -Name $ActiveName
    }
    else{
        $Init_ExitCode = 7
        
        Write-Head("WARNING")
        Write-Paragraph ("A folder of this name already exists.",
        "This folder is no known library folder.","Please change LibraryName in Settings.txt.")
    }

    return $Init_ExitCode
}

function Invoke-CreateFirstLibrary{
    <# 
    .SYNOPSIS
    Tells <HSort.ps1> that this is the first time the user creates a library.
    [$SettingsHistory.xml] is created here in $HOME\AppData\Roaming\HSort\Settings
    #>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$ActiveSettings,

        [Parameter(Mandatory)]
        [string]$ActiveName
    )

    Write-Debug "[Invoke-CreateFirstLibrary]"

    # Assert that no other "LibraryName" folder exists in the TargetDir
    if(-not (Test-Path "$($ActiveSettings.Target)\$ActiveName")){

        $Init_ExitCode = 0

        $SettingsHistory = @{}
        $SettingsHistory.Add($ActiveName, $ActiveSettings)
        $SettingsHistory | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsHistory.xml" 

        # Serialize $Settings_Hst to allow main-script access.
        $ActiveSettings | Export-Clixml -LiteralPath "$($PathsProgram.Settings)\ActiveSettings.xml" -Force

        ### NEW 23/06/2024
        ### Create folder for library specific data in Roaming\HSort\LibraryFiles
        try{
            $null = New-Item -Path "$($PathsProgram.LibFiles)" -ItemType Directory -Name $ActiveName
        }
        catch{
            throw "LibraryName not in SettingsHistory but LibraryFolder exists in LibraryFiles."
        }

        Write-Debug "SettingsHistory.xml exists: False"
        Write-Debug "[Creating] SettingsHistory.xml"
        Write-Debug "[Creating] $ActiveName`n"

    }
    else{
        Write-Head("WARNING")
        Write-Paragraph ("A folder of this name already exists.",
        "This folder is no known library folder.",
        "Please change LibraryName in Settings.txt.",
        " ")

        $Init_ExitCode = 7
    }

    return $Init_ExitCode
}

function Initialize-Script{ 

    $Init_ExitCode = 999

    #===========================================================================
    # Check Program State
    #===========================================================================

    # === ProgramFolder (HSort) does NOT exists ===
    #
    # We assume that this means that this
    # is the very first time the script is run (Init_ExitCode = 1)

    $Init_ExitCode = Test-ProgramFolder

    # === ProgramFolder (HSort) exists ===

    if($Init_ExitCode -eq 0){

        # Validate ProgramFolder
        Assert-ProgramFolder

        # Test if [Settings.txt] exists
        $Init_ExitCode = Test-SettingsTxt

        # === Settings.txt exists ===

        if($Init_ExitCode -eq 0){
    
            # Read Settings.txt
            # We receive the settgins as a hashtable
            $ActiveSettings = Read-Settings -Path $PathsProgram.TxtSettings

            # === Save drive serial number ===
            # Add the serial number of the drive that contains the source folder to [ActiveSettings]
            # If the library is created on an external drive its drive letter might change at some point.
            # Saving the serial number allows us to find the right library even when the drive letter has changed
            $SrcDriveSerial = Get-DriveSerial -SourcePath $ActiveSettings.Source
            $ActiveSettings.Add("DriveSerial",$SrcDriveSerial)

            $Init_ExitCode = Assert-SettingsValid -ActiveSettings $ActiveSettings
    
            if ($Init_ExitCode -eq 0){

                $ActiveTarget = $ActiveSettings.Target
                $ActiveName   = $ActiveSettings.LibraryName
                $ActiveSource = $ActiveSettings.Source

                # Display current settings.
                Write-Paragraph -InformationArray ("YOUR SETTINGS",
                "[ScriptVersion: $($ActiveSettings.ScriptVersion)]",
                "=======================================",
                " ",
                "LibraryName: $($ActiveSettings.LibraryName)",
                " ",
                "Source:      $($ActiveSettings.Source)",
                " ",
                "Target:      $($ActiveSettings.Target)",
                " ",
                "=======================================",
                " ") 
    
                #===========================================================================
                # One or more libraries exist 
                #===========================================================================

                # If SettingsHistory.xml exists we assume that
                # one or more libraries were created in the past.

                if (Test-Path "$($PathsProgram.Settings)\SettingsHistory.xml"){

                    Write-Debug "SettingsHistory.xml exists: TRUE"
                    
                    # === Backup Settings === 
                    #
                    # In case the user answers [n] in the <Start-Script> function.
                    Backup-Settings -Source $PathsProgram.Settings -Target $PathsProgram.Tmp
                    
                    # === Import SettingsHistory.xml ===

                    try{
                        $SettingsHistory = Import-Clixml -LiteralPath "$($PathsProgram.Settings)\SettingsHistory.xml"
                    }
                    catch{
                        #$Init_ExitCode = 4
                        #Write-Information -MessageData "Error: Couldn't import SettingsHistory.xml" -InformationAction Continue
                        throw "Error: Couldn't import SettingsHistory.xml"
                    }
                    
                    # === Case 1 ===
                    # === Library exists ===
                    #
                    # Since [SettingsHistory] contains [ActiveName] we assume
                    # that [ActiveName] is an existing library.   

                    if($SettingsHistory.ContainsKey($ActiveName)){

                        Write-Debug "$ActiveName is an existing library.`n"
                        
                        Assert-LibrariesAccessible -SettingsHistory $SettingsHistory -ActiveName $ActiveName
                        
                        # === CASE 1-a ===
                        #
                        # The library target is not the same as in [SettingsHistory].
                        # We assume the user wants to move the library to a different location.
                        if($SettingsHistory.$ActiveName.Target -ne $ActiveTarget){
                            
                            $Init_ExitCode = Move-Library -SettingsHistory $SettingsHistory -ActiveSettings $ActiveSettings

                        }
                        # === CASE 1-b ====
                        #
                        # Only the library source was changed.
                        # We assume that the user wants to update the library from a new source.
                        else{
                            $Init_ExitCode = Update-Library -SettingsHistory $SettingsHistory -ActiveSettings $ActiveSettings -ActiveSource $ActiveSource
                        }
                    }
                    # === Case 2 ===
                    # === Library does not exists ===
                    #
                    # Since SettingsHistory.xml exists, we assume 
                    # that the script was run before and libraries were created.
                    # And since [SettingsHistory] does not [ActiveName] we assume
                    # that the user want to create a new library with a different name [ActiveName]   

                    else{
                        $Init_ExitCode = New-Library -SettingsHistory $SettingsHistory -ActiveSettings $ActiveSettings -ActiveName $ActiveName
                    }
                }

                #===========================================================================
                # No libraries exist 
                #===========================================================================

                # Since [SettingsHistory.xml]  doesn't exist we assume
                # that no libraries were created before.
                # We create a new and therefore first library.

                <#
                    Initial creation of SettingsHistory.xml

                    This should happen when the user runs the script for the SECOND time.
                    The first time the script is run, only the ProgramFolder structure is created,
                    including Settings.txt.
                    
                    ActiveSettings.xml doesn't exist yet -> Create and Serialize.
                    SettingsHistory.xml does NOT exist yet either -> Create and Serialize.

                #>
                else{

                    $Init_ExitCode = Invoke-CreateFirstLibrary -ActiveSettings $ActiveSettings -ActiveName $ActiveName
                }
            }
        }
    }

    return $Init_ExitCode
}

function Start-Script{

    $StartScript_ExitCode = 999

    while($true){

        $ConfirmStartScript = Read-Host  "Press [y] to confirm and start the script. Press [n] to abort."

        if($ConfirmStartScript -eq "y"){
            Write-Information -MessageData "`nStarting script. Please wait...`n" -InformationAction Continue
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

Export-ModuleMember -Function Initialize-Script,Confirm-Settings,Start-Script,Restore-Settings,Resume-OnError,Write-Paragraph,Get-UserInput,Write-Head -Variable Init_ExitCode,ConfirmSettings_ExitCode,StartScript_ExitCode,ResumeExitCode,UserResponse
