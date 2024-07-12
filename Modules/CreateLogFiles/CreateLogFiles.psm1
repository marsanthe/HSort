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

    $FilePath = "$TargetDir\ProcessedObjects.txt"

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

        Export-Csv -InputObject $Item_Log -Path "$TargetDir\ProcessedObjects.csv" -NoTypeInformation -Append
    
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

function Update-ExcludedLog{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$ExcludedObjects,

        [Parameter(Mandatory)]
        [string]$TargetDir

    )

    $FilePath = "$TargetDir\ExcludedObjects.txt"

    foreach($Obj in $ExcludedObjects.Keys){

        $CustomObject = [PSCustomObject]@{
            
            ParentPath = "{0}" -f ($ExcludedObjects.$Obj.ParentPath)
            # Path       = "{0}" -f ($ExcludedObjects.$Obj.Path)
            Name       = "{0}" -f ($ExcludedObjects.$Obj.ObjectName)
            # Extension  = "{0}" -f ($ExcludedObjects.$Obj.Extension)
            Reason     = "{0}" -f ($ExcludedObjects.$Obj.Reason)

        }

        $Property0 = @{
            expression = "ParentPath"
            width = 160
        }
        
        $Property1 = @{
            expression = "Name"
            width = 160
        }
        $Property2 = @{
            expression = "Reason"
            width      = 60
        }
        
        $CustomObject | Format-Table -Property $Property0,
        $Property1,$Property2 -Wrap | Out-File -FilePath $FilePath -Encoding unicode -Width 400 -Append -Force

        Export-Csv -InputObject $CustomObject -Path "$TargetDir\ExcludedObjects.csv" -NoTypeInformation -Append
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

Export-ModuleMember -Function Update-CopiedLog, Update-ExcludedLog