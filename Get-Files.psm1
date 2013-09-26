#requires -version 3

function Get-Files
{
[CmdletBinding()]
param
(
[Parameter(Position=0, Mandatory=$true)] [string[]]$inputs,
[Parameter(Mandatory=$false)] [string]$match = ".*",
[Parameter(Mandatory=$false)] [string]$matchDesc = "",
[Parameter(Mandatory=$false)] [switch]$acceptFiles = $true,
[Parameter(Mandatory=$false)] [switch]$acceptFolders = $false
)

$inFilesAll = @()
    foreach ($input in $inputs)
    {
        # Workaround for Get-Childitem bug
        if(!($input -match '`')) 
        { $inputEsc = [System.Management.Automation.WildcardPattern]::Escape($input) } else { $inputEsc = $input }
        
        try 
        { 
            $inFiles = Get-ChildItem -LiteralPath ([System.Management.Automation.WildcardPattern]::Unescape($inputEsc)) -ErrorAction Stop `
                     | ?{ $_ -match $match }
        }
        catch
        {
            $msg = "Error: Failed processing $input"
            if($_.Exception.GetType().Name -eq "ItemNotFoundException")
            {
                $msg += ": File or Directory not found"
            }
            else { $msg +=  ".`nError Message: $($_.Exception.Message)" } 
            throw $msg
        }

        if (!$inFiles)
        {
            if(Test-Path $input -pathType container)
            { Write-Host "Notice: No $matchDesc files found in $input." -ForegroundColor Gray}
            else 
            { 
                Write-Host "Error: $input is not a $matchDesc file."
                exit
            }
        }
        
        $inFilesAll += $inFiles
    }
    return $inFilesAll
}

Export-ModuleMember Get-Files