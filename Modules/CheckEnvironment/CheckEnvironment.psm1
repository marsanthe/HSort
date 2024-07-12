function Assert-PSVersion{

    <# 
    .SYNOPSIS
    Asserts that PSVersion is at least 5.1
    #>
    Param(
        [Parameter(Mandatory)]
        [int32]$Major,

        [Parameter(Mandatory)]
        [int32]$Minor
    )

    $PSVersion_ExitCode = 0

    if ($Major -lt 5) {
        Write-Information -MessageData "This script requires at least Powershell Version 5.1`nExiting Script." -InformationAction Continue
        PSVersion_ExitCode = 1
        
    }
    elseif (($Major -eq 5) -and ($Minor -lt 1)) {
        Write-Information -MessageData "This script requires at least Powershell Version 5.1`nExiting Script." -InformationAction Continue
        PSVersion_ExitCode = 1
        
    }
    else {
        Write-Information -MessageData "PSVersion: [OK]" -InformationAction Continue
    }

    return $PSVersion_ExitCode


}


function Confirm-SevenZip{

    <# 
    .SYNOPSIS
    Asserts that 7-Zip is installed
    .DESCRIPTION
    Asserts that 7-Zip is installed by checking if
    CurrentVersion\Uninstall\ contains an app by Igor Pavlov.
    ...
    CurrentVersion\Uninstall\ contains an app with
    App.DisplayName equal to 7-Zip
    #>

    ### [BEGIN] Assert-7ZipInstalled ###
    #
    $64BitPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $32BitPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    
    $App64Found = $False
    $App32Found = $False


    # If 64-bit OS
    # Check if 64-bit 7Zip is installed 
    # if not check if 32-bit 7Zip is installed (on 64-bit OS)
    if ([System.Environment]::Is64BitOperatingSystem) {
        $64BitApps = Get-ItemProperty "HKLM:\$64BitPath"
        foreach ($App in $64BitApps) {
            if ($App.Publisher -eq "Igor Pavlov") {
                $AppName = "$($App.DisplayName)"
                if ($AppName.StartsWith("7-Zip")){
                    $SevenZip64 = $App
                    $App64Found = $True
                }
            }
        }
        if ($App64Found -eq $False) {
            $32BitApps = Get-ItemProperty "HKLM:\$32BitPath"
            foreach ($App in $32BitApps) {
                if ($App.Publisher -eq "Igor Pavlov") {
                    $AppName = "$($App.DisplayName)"
                    if ($AppName.StartsWith("7-Zip")) {
                        $SevenZip64 = $App
                        $App64Found = $True
                    }
                }
            }      
        }
    }
    # If 32-bit OS
    # Check if 32-bit 7Zip is installed 
    else {
        $32BitApps = Get-ItemProperty "HKLM:\$32BitPath"
        foreach ($App in $32BitApps) {
            if ($App.Publisher -eq "Igor Pavlov") {
                $AppName = "$($App.DisplayName)"
                if ($AppName.StartsWith("7-Zip")) {
                    $SevenZip64 = $App
                    $App64Found = $True
                }
            }
        }        
    }

    if ($App64Found -eq $true) {
        Write-Information -MessageData "7-Zip: [OK]" -InformationAction Continue
        $SevenZipPath = $SevenZip64.InstallLocation
        $7zip = "$SevenZipPath\7z.exe"
    }
    elseif ($App32Found -eq $true) {
        Write-Information -MessageData "7-Zip: [OK]" -InformationAction Continue
        $SevenZipPath = $SevenZip32.InstallLocation
        $7zip = "$SevenZipPath\7z.exe"
    }
    else {
        Write-Information -MessageData "7-Zip: [Error]" -InformationAction Continue
        Write-Information -MessageData "7-Zip not found. Exiting" -InformationAction Continue
        $7zip = ""

    }

    return $7zip

}


Export-ModuleMember -Function Assert-PSVersion, Confirm-SevenZip -Variable PSVersion_ExitCode,7zip
