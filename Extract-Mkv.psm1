function Extract-Mkv
{
[CmdletBinding()]
param
(
[Parameter(Position=0, Mandatory=$true, HelpMessage='Files or Directories to process (comma-delimited). Can take mixed.')]
[alias("i")]
[string[]]$Inputs,
[Parameter(Mandatory=$false, HelpMessage='Extract tracks. Possible values: none, all (Default), video, audio, subtitles, [Track IDs]')]
[alias("t")]
[string[]]$Tracks = @("all"),
[Parameter(Mandatory=$false, HelpMessage='Extract attachments. Possible values: none (Default), all, fonts')]
[alias("a")]
[string[]]$Attachments = @(),
[Parameter(Mandatory=$false, HelpMessage='Extract chapters. Possible values: none (Default), xml, simple')]
[alias("c")]
[string]$Chapters,
[Parameter(Mandatory=$false, HelpMessage='Extract v2 timecodes for video tracks. NOT IMPLEMENTED, YET.')]
[alias("tc")]
[switch]$Timecodes,
[Parameter(Mandatory=$false, HelpMessage='Filename pattern for extracted tracks.')]
[alias("tp")]
[string]$TrackPattern = '$f_$i',
[Parameter(Mandatory=$false, HelpMessage='Filename pattern for extracted attachments.')]
[alias("ap")]
[string]$AttachmentPattern = '$f_Attachments\$n',
[Parameter(Mandatory=$false, HelpMessage='Output directory. Defaults to parent directory of the inputs')]
[alias("o")]
[string]$OutDir,
[Parameter(Mandatory=$false, HelpMessage='Increase verbosity.')]
[alias("v")]
[switch]$Verb,
[Parameter(Mandatory=$false, HelpMessage='Parse the whole file instead of relying on the index.')]
[alias("f")]
[switch]$ParseFully = $false,
[Parameter(Mandatory=$false, HelpMessage='Suppress status output.')]
[alias("q")]
[switch]$Quiet = $false,
[Parameter(Mandatory=$false, HelpMessage='Also try to extract the CUE sheet from the chapter information and tags for tracks.')]
[switch]$Cuesheet = $false,
[Parameter(Mandatory=$false, HelpMessage='Extract track data to a raw file.')]
[switch]$Raw = $false,
[Parameter(Mandatory=$false, HelpMessage='Extract track data to a raw file including the CodecPrivate as a header.')]
[switch]$FullRaw = $false
)
    if($verb) { $VerbosePreference = "Continue" }

    #CheckMissingCommands -commands "mkvinfo.exe"
   
    try { $mkvs = Get-Files $inputs -match '.mk[v|a|s]$' -matchDesc Matroska -acceptFolders }
    catch
    {
        if($_.Exception.WasThrownFromThrowStatement -eq $true)
        {
            Write-Host $_.Exception.Message -ForegroundColor Red
            break
        }
        else {throw $_.Exception}
    }

    [PSObject[]]$extractData = @()
    
    foreach($mkv in $mkvs)
    {
        $mkvInfo = Get-MkvInfo -file $mkv.FullName
        $tracks | ?{($_ -eq "all")} | %{$mkvInfo.tracks | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}
        $tracks | ?{($_ -match '[0-9]+')} | %{$mkvInfo.GetTracksByID($_)} | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force
        $tracks | ?{($_ -match 'subtitles|audio|video')} | %{$mkvInfo.GetTracksByType($_)} | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force
        
        $attachments | ?{($_ -eq "all")} | %{$mkvInfo.Attachments | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}
        $attachments | ?{($_ -eq "fonts")} | %{$mkvInfo.GetAttachmentsByExtension("ttf|ttc|otf|fon") | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}

        if(!$Quiet)
        {
            Write-Host "$($mkv.Name)" -ForegroundColor White -NoNewline

            if($mkvInfo.Title) { Write-Host " (Title: $($mkvInfo.Title))" -ForegroundColor Gray}
            else { Write-Host "`n`n" }
            Write-Host "(Tracks marked yellow will be extracted)`n" -ForegroundColor Gray 

            $tables = @{ video     = "Video tracks"
                         audio     = "Audio tracks"
                         subtitles = "Subtitle tracks"
                         attachments = "Attachments"
                       }
            $tables.GetEnumerator() | %{
               Write-Host $_.Value -NoNewline -ForegroundColor DarkBlue -BackgroundColor DarkGray
               $mkvInfo | Format-MkvInfo-Table -type $_.Key
            }
        }
        
        $cmnFlags = @{ "verbose" = $Verb
                       "quiet" = $Quiet
                       "parse-fully" = $ParseFully
                     }
        $trackFlags = $cmnFlags + @{ "cuesheet" = $Cuesheet
                                     "raw"      = $Raw
                                     "fullraw"  = $FullRaw 
                                   }

        $mkvInfo = ExtractTracks -MkvInfo $mkvInfo -Pattern $trackPattern -OutDir $outdir -flags $trackFlags
        $mkvInfo = ExtractAttachments -MkvInfo $mkvInfo -pattern $attachmentPattern -OutDir $outdir -flags $cmnFlags
        
    }
}

function ExtractTracks([PSCustomObject]$mkvInfo, [regex]$pattern, [string]$outDir, [hashtable]$flags)
{
    if (!$outdir) {$outdir = $mkvInfo.Path | Split-Path -Parent }
    $mkvInfo.Tracks | ?{$_._toExtract -eq $true -and $_.CodecExt} | %{$extArgs=@(); $extCnt=0} {
        $patternVars = @{
        '$f' = [System.IO.Path]::GetFileNameWithoutExtension($mkvInfo.Path)
        '$i' = [string]$_.ID
        '$t' = [string]$_.Type
        '$n' = [string]$_.Name
        '$l' = [string]$_.Language
        } 

        $patternVars.GetEnumerator() | ForEach-Object {$outFile=$pattern} { 
            $outFile = $outFile -replace [regex]::escape($_.Key), $_.Value
        }

        $outFile = "$(Join-Path $outDir $outFile).$($_.CodecExt)"
        $mkvInfo.GetTracksByID($_.ID) | Add-Member -NotePropertyName _ExtractPath -NotePropertyValue $outFile -Force

        $extArgs += "$($_.ID):$outFile "
        $extCnt++
    }

    if($extCnt -gt 0)
    {
        &mkvextract --ui-language en tracks $mkvInfo.Path $extArgs `        ($flags.GetEnumerator()| ?{$_.Value -eq $true -and $_.Key -ne "quiet"} | %{"--$($_.Key)"}) `        | Tee-Object -Variable dbgLog | %{             if($_ -match '(?:Progress: )([0-9]{1,2})(?:%)')            {                $extPercent = $matches[1]                Write-Progress -Activity $mkvInfo.Path -Status "Extracting $extCnt tracks..." -PercentComplete $extPercent -CurrentOperation "$extPercent% complete"            }            elseif($_ -match 'Error: .*')            {                Write-HostEx $matches[0] -ForegroundColor Red -If (!$flags.quiet)                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _ExtractPath            }            elseif($_ -match '(?:Progress: )(100)(?:%)')            {                $extPercent = $matches[1]                Write-Progress -Activity $mkvInfo.Path -Status "Extraction complete." -Complete                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _toExtract                Write-HostEx "`nDone.`n" -ForegroundColor Green -If (!$flags.quiet)            }            elseif($_ -match '^\(mkvextract\)')            {                Write-Verbose $_            }            elseif($_.Trim() -and (!$flags.quiet))            { Write-HostEx $_ -ForegroundColor Gray -If (!$flags.quiet) }           }
    } else { Write-Verbose "No tracks to extract" }
    return $mkvInfo
}

function ExtractAttachments([PSCustomObject]$mkvInfo, [regex]$pattern, [string]$outDir, [hashtable]$flags)
{
    if (!$outdir) {$outdir = $mkvInfo.Path | Split-Path -Parent }
    $mkvInfo.Attachments | ?{$_._toExtract -eq $true} | %{$extArgs=@(); $extCnt=0} {
        $extCnt++
        $patternVars = @{
        '$f' = [System.IO.Path]::GetFileNameWithoutExtension($mkvInfo.Path)
        '$i' = [string]$_.UID
        '$n' = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        } 

        $patternVars.GetEnumerator() | ForEach-Object {$outFile=$pattern} { 
            $outFile = $outFile -replace [regex]::escape($_.Key), $_.Value
        }

        $outFile = "$(Join-Path $outDir $outFile)$([System.IO.Path]::GetExtension($_.Name))"
        $mkvInfo.GetAttachmentsByUID($_.UID) | Add-Member -NotePropertyName _ExtractPath -NotePropertyValue $outFile -Force

        $extArgs += "$extCnt`:$outFile "
    }

    if($extCnt -gt 0)
    {
        &mkvextract --ui-language en attachments $mkvInfo.Path $extArgs `        ($flags.GetEnumerator()| ?{$_.Value -eq $true -and $_.Key -ne "quiet"} | %{"--$($_.Key)"}) `        | Tee-Object -Variable dbgLog | %{ $doneCnt=0 } {            if($_ -match "^The attachment (#[0-9]+), ID ([0-9]+), MIME type (.*?), size ([0-9]+), is written to '(.*?) '\.$")            {                $doneCnt++                $extPercent = ($doneCnt / $extCnt) * 100                Write-Progress -Activity $mkvInfo.Path -Status "Extracting $extCnt attachments..." -PercentComplete $extPercent -CurrentOperation ($matches[5] | Split-Path -Leaf)                $mkvInfo.Attachments = $mkvInfo.Attachments | ?{$_.UID -eq [uint64]$matches[2]}| Select-Object * -ExcludeProperty _toExtract            }            elseif($_ -match 'Error: .*')            {                Write-HostEx $matches[0] -ForegroundColor Red -If (!$flags.quiet)                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _ExtractPath                $err = $true            }            elseif($_ -match '^\(mkvextract\)')            {                Write-Verbose $_            }            elseif($_.Trim() -and (!$flags.quiet))            { Write-HostEx $_ -ForegroundColor Gray -If (!$flags.quiet) }         }

        Write-HostEx "`nDone.`n" -ForegroundColor Green -If (!$flags.quiet -and !$err)

    } else { Write-Verbose "No attachments to extract" }
    return $mkvInfo   
}

Export-ModuleMember Extract-Mkv