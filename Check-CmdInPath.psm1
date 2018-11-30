function Check-CmdInPath
{
[CmdletBinding()]
param
(
[Parameter(Position=0, Mandatory=$true)] [string]$Cmd,
[Parameter(Mandatory=$false)] [string]$Name = $Cmd
)

    if(!(Get-Command $Cmd -ErrorAction SilentlyContinue))
    {
        Write-Host "Fatal Error: missing $Cmd. Make sure $Name is installed and in your PATH environment variable" -ForegroundColor Red
        break
    }

}
