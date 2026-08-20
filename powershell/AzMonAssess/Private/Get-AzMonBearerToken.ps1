#requires -Version 7.0

function Get-AzMonBearerToken {
    <#
    .SYNOPSIS
        Returns a plain-string bearer token for a resource URL, tolerating
        both the legacy (string) and current (SecureString) return shapes
        of Get-AzAccessToken across Az.Accounts versions.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $ResourceUrl)

    $tokenObj = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop
    $raw = $tokenObj.Token
    if ($raw -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($raw)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        } finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    return [string]$raw
}
