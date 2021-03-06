<#

.ForwardHelpTargetName Write-Host
.ForwardHelpCategory Cmdlet

#>
function Write-HostEx {
    [CmdletBinding(HelpUri='http://go.microsoft.com/fwlink/?LinkID=113426', RemotingCapability='None')]
    param(
        [Parameter(Position=0, ValueFromPipeline=$true, ValueFromRemainingArguments=$true)]
        [System.Object]
        ${Object},

        [switch]
        ${NoNewline},

        [System.Object]
        ${Separator},

        [System.ConsoleColor]
        ${ForegroundColor},

        [System.ConsoleColor]
        ${BackgroundColor},

        [bool]
        ${If}=$true)

    begin
    {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Host', [System.Management.Automation.CommandTypes]::Cmdlet)
            if ($If)
            {
                [Void]$PSBoundParameters.Remove("If")
            }
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            if($if -or $if -eq $null) { $steppablePipeline.Begin($PSCmdlet) }
        } catch {
            throw
        }
    }

    process
    {
        try {
            if($if) {$steppablePipeline.Process($_) }
        } catch {
            throw
        }
    }

    end
    {
        try {
            if($if) {$steppablePipeline.End() }
        } catch {
            throw
        }
    }
}

Export-ModuleMember Write-HostEx