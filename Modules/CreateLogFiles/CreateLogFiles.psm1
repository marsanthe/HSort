$LogModuleTimestamp = Get-Date -Format "dd_MM_yyyy"

function Add-LogEntry{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$ObjectProperties,

        [Parameter(Mandatory)]
        [string]$Path

    )

    # We don't do "$Item_Log = [pscustomobject]$ObjectProperties"
    # since we want to preserve the order of the properties.

    $Item_Log = [PSCustomObject]@{

        #Extension = "{0}" -f ($ObjectProperties.Extension)
        Source = "{0}" -f ($ObjectProperties.ObjectSource)
        Target           = "{0}" -f "$($ObjectProperties.ObjectTarget)\$($ObjectProperties.ObjectNewName)$($ObjectProperties.NewExtension)"
        OriginalFileHash = "{0}" -f ($ObjectProperties.SourceHash)
        CopiedFileHash = "{0}" -f ($ObjectProperties.TargetHash)

    }

    #Calculated properties of an Item_Log object

    <#
        .source
        https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_calculated_properties?view=powershell-7.4

        expression - A string or script block used to calculate the value of the new property.

        >>>If the expression is a string, the value is interpreted as a property name on the input object.<<<

        This is a shorter option than expression = { $_.<PropertyName> }.
    #>

    # $Property0 = @{
    # expression = "Extension"
    # width = 30}

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
    $Property3 -Wrap | Out-File -FilePath "$Path\ObjectLog $LogModuleTimestamp.txt" -Encoding unicode -Width 400 -Append -Force


}

function Add-SkippedLogEntry{

    Param(
        [Parameter(Mandatory)]
        [hashtable]$SkippedObjectProperties,

        [Parameter(Mandatory)]
        [string]$Path

        )

    $SkippedFilesLog = [PSCustomObject]@{

        Path = "{0}" -f ($SkippedObjectProperties.Path)
        Reason = "{0}" -f ($SkippedObjectProperties.Reason)
        Test = "Test"
    }

    $Property0 = @{
    expression = "Path"
    width = 160}

    $Property1 = @{
    expression = "Reason"
    width = 60}

    $SkippedFilesLog | Format-Table -Property $Property0,
    $Property1 -Wrap | Out-File -FilePath "$Path\SkippedObjects $LogModuleTimestamp.txt" -Encoding unicode -Width 400 -Append -Force
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
    $Property2 -Wrap | Out-File -FilePath "$Path\$FileName $LogModuleTimestamp.txt" -Encoding unicode -Width 260 -Append -Force
}