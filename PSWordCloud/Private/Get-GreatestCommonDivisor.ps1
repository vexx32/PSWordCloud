using namespace System.Management.Automation

function Get-GreatestCommonDivisor {
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]
        $Numerator,

        [Parameter()]
        [int]
        $Denominator
    )
    process {
        while ($Denominator -gt 0 -and $Numerator -ne $Denominator) {
            $Numerator, $Denominator = $Denominator, ($Numerator % $Denominator)
        }

        return $Numerator
    }
}