function Assert-SufficientDiskSpace{
    <# 
    .SYNOPSIS
    Checks if diskspace is sufficient.
    .DESCRIPTION
    Checks if diskspace is sufficient and
    returns DiskSpace_ExitCode.
    #>

    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir,

        [Parameter(Mandatory)]
        [string]$LibraryParentFolder,

        [switch]$Talkative
    )

    Write-Information -MessageData "Asserting that disk space is sufficient. Please wait..." -InformationAction Continue

    $DiskSpace_ExitCode = 999

    <# 
        Consider space required for ComicInfo.xml and other metadata.
    #>

    $Buffer = 100
    $SourceSize = (($SourceDir | Get-ChildItem -Recurse | Measure-Object Length -Sum).sum)
    $SourceSize = [int][math]::Ceiling(($SourceSize)/1MB) + $Buffer

    $TargetDrive = (Get-Item $LibraryParentFolder).PSDrive.Name

    $AvailableSpace = (Get-Volume -DriveLetter $TargetDrive).SizeRemaining
    $AvailableSpace = [int][math]::Floor(($AvailableSpace)/1MB)

    if($AvailableSpace -le $SourceSize){
        $DiskSpace_ExitCode = 1
        Write-Information -MessageData "Insufficient disk space." -InformationAction Continue
        Write-Information -MessageData "Please select a different LibraryParentFolder in Settings.txt" -InformationAction Continue
        Write-Information -MessageData "Exiting`n" -InformationAction Continue
    }
    else{
        if($Talkative){
            Write-Information -MessageData "Sufficient disk space. Continuing.`n" -InformationAction Continue            
        }
        $DiskSpace_ExitCode = 0
    }

    return $DiskSpace_ExitCode
}

Export-ModuleMember -Function Assert-SufficientDiskSpace -Variable DiskSpace_ExitCode



