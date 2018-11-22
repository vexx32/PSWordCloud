function Convert-ToRadians {
    <#
    .SYNOPSIS
    Converts a value from degrees to radians.

    .DESCRIPTION
    Uses [Math]::PI to calculate the corresponding radian value from the input degrees value.

    .PARAMETER Degrees
    Angle to convert to radians.

    .EXAMPLE
    180 | ConvertTo-Radians

    3.14159265358979
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline)]
        [double]
        $Degrees
    )
    process {
        ([Math]::PI / 180) * $Degrees
    }
}