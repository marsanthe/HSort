function New-Header{
    Param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$FilePath,

        [switch]$Subheader
    )

    if(! $Subheader){

$Header = @"


$Name [$Timestamp]
=================================================

"@
    }
    else{
$Header = @"


$Name
=================================================

"@        
    }

    $Header | Out-File -FilePath $FilePath -Encoding unicode -Width 400 -Append -Force

}

function Update-CopiedLog{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$ObjectsToCopy,

        [Parameter(Mandatory)]
        [string]$TargetDir

    )

    $FilePath = "$TargetDir\ObjectLog.txt"

    New-Header -Name "Copied Objects" -FilePath $FilePath
    
    foreach ($HashedID in $ObjectsToCopy.Keys) {

        $ObjectProperties = $ObjectsToCopy.$HashedID

        # We don't do "$Item_Log = [pscustomobject]$ObjectProperties"
        # since we want to preserve the order of the properties.
    
        $Item_Log = [PSCustomObject]@{
    
            Source = "{0}" -f ($ObjectProperties.ObjectSource)
            Target           = "{0}" -f "$($ObjectProperties.ObjectTarget)\$($ObjectProperties.TargetName)$($ObjectProperties.NewExtension)"
            OriginalFileHash = "{0}" -f ($ObjectProperties.SourceHash)
            CopiedFileHash = "{0}" -f ($ObjectProperties.TargetHash)
    
        }
    
        $Property0 = @{
        expression = "Source"
        width = 100}
    
        $Property1 = @{
        expression = "Target"
        width = 100}
    
        $Property2 = @{
        expression = "OriginalFileHash"
        width = 40}
    
        $Property3 = @{
        expression = "CopiedFileHash"
        width = 40}
    
        $Item_Log | Format-Table -Property $Property0,
        $Property1,
        $Property2,
        $Property3 -Wrap | Out-File -FilePath $FilePath -Encoding unicode -Width 400 -Append -Force
    }
}

function Update-SkippedLog{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$SkippedObjects,

        [Parameter(Mandatory)]
        [string]$TargetDir

    )

    $FilePath = "$TargetDir\SkippedObjects.txt"

    New-Header -Name "Skipped Objects" -FilePath $FilePath

    foreach ($Parent in $SkippedObjects.Keys){

        New-Header -Name "Parent Dir: $Parent" -FilePath $FilePath -Subheader

        foreach($Object in $SkippedObjects.$Parent.Keys){

            $SkippedObjectProperties = $SkippedObjects.$Parent.$Object

            $SkippedFilesLog = [PSCustomObject]@{
        
                Path = "{0}" -f ($SkippedObjectProperties.Path)
                Reason = "{0}" -f ($SkippedObjectProperties.Reason)

            }
        
            $Property0 = @{
            expression = "Path"
            width = 160}
        
            $Property1 = @{
            expression = "Reason"
            width = 60}
        
            $SkippedFilesLog | Format-Table -Property $Property0,
            $Property1 -Wrap | Out-File -FilePath $FilePath -Encoding unicode -Width 400 -Append -Force
        }
    }
}

function Add-TimingLogEntry{

    Param(

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [hashtable]$ObjectTimer,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$FileName

        )

    $TimingTable = [PSCustomObject]@{

        Number = "{0}" -f $ObjectTimer.Number
        Object = "{0}" -f ($Name)
        Time = "{0}" -f ($ObjectTimer.Time)

    }

    $Property0 = @{
    expression = "Number"
    width = 20}

    $Property1 = @{
    expression = "Object"
    width = 180}

    $Property2 = @{
    expression = "Time"
    width = 40}

    $TimingTable | Format-Table -Property $Property0,
    $Property1,
    $Property2 -Wrap | Out-File -FilePath "$Path\$Timestamp $FileName.txt" -Encoding unicode -Width 260 -Append -Force
}

Export-ModuleMember -Function Update-CopiedLog, Update-SkippedLog