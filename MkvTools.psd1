@{
RootModule = 'MkvTools.psm1'
ModuleVersion = '0.0.1'
GUID = 'afd66be9-872e-4316-96ab-eee551f7fa5f'
Author = 'line0'
Description = 'A small set of PowerShell modules that enable batch processing of Matroska files with mkvtoolnix.'

PowerShellVersion = '3.0'

NestedModules = @('Extract-Mkv.psm1', 'Get-Files.psm1', 'Get-Mkvinfo.psm1', 'Write-HostEx.psm1')
FunctionsToExport = @('Extract-Mkv','Format-MkvInfo-Table', 'Get-MkvInfo', 'Get-Files', 'Write-HostEx')
CmdletsToExport = ''
VariablesToExport = '*'
AliasesToExport = '*'


# HelpInfo URI of this module
# HelpInfoURI = ''
}