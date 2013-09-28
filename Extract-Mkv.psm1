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
[string[]]$Chapters = @(),
[Parameter(Mandatory=$false, HelpMessage='Extract v2 timecodes for video tracks. NOT IMPLEMENTED, YET.')]
[alias("tc")]
[switch]$Timecodes,
[Parameter(Mandatory=$false, HelpMessage='Filename pattern for extracted tracks.')]
[alias("tp")]
[string]$TrackPattern = '$f_$i',
[Parameter(Mandatory=$false, HelpMessage='Filename pattern for extracted attachments.')]
[alias("ap")]
[string]$AttachmentPattern = '$f_Attachments\$n',
[Parameter(Mandatory=$false, HelpMessage='Filename pattern for extracted chapters.')]
[alias("cp")]
[string]$ChapterPattern = '$f_Chapters',
[Parameter(Mandatory=$false, HelpMessage='Output directory. Defaults to parent directory of the inputs')]
[alias("o")]
[string]$OutDir,
[Parameter(Mandatory=$false, HelpMessage="Verbosity level: `n 0: Don't display tables`n 1: Default`n 2: Show additonal mkvextract output`n 3: Pass --verbose to mkvextract")]
[alias("v")]
[int]$Verbosity=1,
[Parameter(Mandatory=$false, HelpMessage='Parse the whole file instead of relying on the index.')]
[alias("f")]
[switch]$ParseFully = $false,
[Parameter(Mandatory=$false, HelpMessage='Suppress status output.')]
[alias("q")]
[switch]$Quiet = $false,
[Parameter(Mandatory=$false, HelpMessage='Recurse directories.')]
[alias("r")]
[switch]$Recurse = $false,
[Parameter(Mandatory=$false, HelpMessage='Also try to extract the CUE sheet from the chapter information and tags for tracks.')]
[switch]$Cuesheet = $false,
[Parameter(Mandatory=$false, HelpMessage='Extract track data to a raw file.')]
[switch]$Raw = $false,
[Parameter(Mandatory=$false, HelpMessage='Extract track data to a raw file including the CodecPrivate as a header.')]
[switch]$FullRaw = $false
)

    if ($Quiet) { $Verbosity=-1 }
    elseif($Verbosity -eq 3) { $VerbosePreference = "Continue" }

    #CheckMissingCommands -commands "mkvinfo.exe"
   
    try { $mkvs = Get-Files $inputs -match '.mk[v|a|s]$' -matchDesc Matroska -acceptFolders -recurse:$Recurse }
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
    
    $doneCnt = 0
    $activityMsg = "Extracting from $($mkvs.Count) files..."
    Write-Progress -Activity $activityMsg -Id 0 -PercentComplete 0 -Status "File 1/$($mkvs.Count)"

    foreach($mkv in $mkvs)
    {
        $mkvInfo = Get-MkvInfo -file $mkv.FullName
        $tracks | ?{($_ -eq "all")} | %{$mkvInfo.tracks | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}
        $tracks | ?{($_ -match '[0-9]+')} | %{$mkvInfo.GetTracksByID($_)} | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force
        $tracks | ?{($_ -match 'subtitles|audio|video')} | %{$mkvInfo.GetTracksByType($_)} | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force
        
        $attachments | ?{($_ -eq "all")} | %{$mkvInfo.Attachments | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}
        $attachments | ?{($_ -eq "fonts")} | %{$mkvInfo.GetAttachmentsByExtension("ttf|ttc|otf|fon") | Add-Member -NotePropertyName _toExtract -NotePropertyValue $true -Force}
 
        Write-HostEx "$($mkv.Name)" -ForegroundColor White -NoNewline:($Verbosity -ge 1) -If ($Verbosity -ge 0)

        if($mkvInfo.Title) 
        { 
            Write-HostEx " (Title: $($mkvInfo.Title))" -ForegroundColor Gray -If ($Verbosity -ge 1)
        }
        else { Write-HostEx "`n`n" -If ($Verbosity -ge 1)}

        Write-Verbose "(Tracks marked yellow will be extracted)`n"
        
        if($Verbosity -ge 1)
        {
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
        
        $cmnFlags = @{ "vb" = $Verbosity
                       "verbose" = ($Verbosity -ge 3)
                       "parse-fully" = $ParseFully
                     }
        $trackFlags = $cmnFlags + @{ "cuesheet" = $Cuesheet
                                     "raw"      = $Raw
                                     "fullraw"  = $FullRaw 
                                   }

        $mkvInfo = ExtractTracks -MkvInfo $mkvInfo -Pattern $trackPattern -OutDir $OutDir -flags $trackFlags
        $mkvInfo = ExtractAttachments -MkvInfo $mkvInfo -pattern $attachmentPattern -OutDir $OutDir -flags $cmnFlags
        $mkvInfo = ExtractChapters -MkvInfo $mkvInfo -Pattern $ChapterPattern -OutDir $OutDir -vb $Verbosity -types $Chapters
        
        $doneCnt++
        Write-Progress -Activity $activityMsg -Id 0 -PercentComplete (100*$doneCnt/$mkvs.Count) -Status "File $($doneCnt+1)/$($mkvs.Count)"
    }
     Write-Progress -Activity $activityMsg -Id 0 -Completed 
}

function ExtractTracks([PSCustomObject]$mkvInfo, [string]$pattern, [string]$outDir, [hashtable]$flags)
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
        Write-HostEx "Extracting $extCnt tracks..." -If ($flags.vb -ge 0)
        &mkvextract --ui-language en tracks $mkvInfo.Path $extArgs `        ($flags.GetEnumerator()| ?{$_.Value -eq $true -and $_.Key -ne "vb"} | %{"--$($_.Key)"}) `        | Tee-Object -Variable dbgLog | %{             if($_ -match '(?:Progress: )([0-9]{1,2})(?:%)')            {                $extPercent = $matches[1]                Write-Progress -Activity "   $($mkvInfo.Path | Split-Path -Leaf)" -Status "Extracting $extCnt tracks..." -PercentComplete $extPercent -CurrentOperation "$extPercent% complete" -Id 1 -ParentId 0            }            elseif($_ -match 'Error: .*')            {                Write-HostEx $matches[0] -ForegroundColor Red -If ($flags.vb -ge 0)                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _ExtractPath            }            elseif($_ -match "^Extracting track ([0-9]+) with the CodecID '(.*?)' to the file '(.*?) '\. Container format: (.*?)$")            {                Write-HostEx "#$($matches[1]): $($matches[3]) ($($matches[4]))" -If ($flags.vb -ge 2) -ForegroundColor Gray            }            elseif($_ -match '(?:Progress: )(100)(?:%)')            {                Write-HostEx "Done.`n" -ForegroundColor Green -If ($flags.vb -ge 0 -and $extPercent -ne 100)                $extPercent = $matches[1]                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _toExtract            }            elseif($_ -match '^\(mkvextract\)')            {                Write-Verbose $_            }            elseif($_.Trim())            { Write-HostEx $_ -ForegroundColor Gray -If ($flags.vb -ge 0) }           }
    } else { Write-HostEx "No tracks to extract" -ForegroundColor Gray -If ($flags.vb -ge 2)}
    return $mkvInfo
}

function ExtractAttachments([PSCustomObject]$mkvInfo, [string]$pattern, [string]$outDir, [hashtable]$flags)
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
        Write-HostEx "Extracting $extCnt attachments..." -If ($flags.vb -ge 0)
        &mkvextract --ui-language en attachments $mkvInfo.Path $extArgs `        ($flags.GetEnumerator()| ?{$_.Value -eq $true -and $_.Key -ne "vb"} | %{"--$($_.Key)"}) `        | Tee-Object -Variable dbgLog | %{ $doneCnt=0 } {            if($_ -match "^The attachment (#[0-9]+), ID (?:-)?([0-9]+), MIME type (.*?), size ([0-9]+), is written to '(.*?) '\.$")            {                $doneCnt++                $extPercent = ($doneCnt / $extCnt) * 100                Write-Progress -Activity "   $($mkvInfo.Path | Split-Path -Leaf)" -Status "Extracting $extCnt attachments..." -PercentComplete $extPercent -CurrentOperation ($matches[5] | Split-Path -Leaf) -Id 1 -ParentId 0                $mkvInfo.Attachments = $mkvInfo.Attachments | ?{$_.UID -eq [uint64]$matches[2]}| Select-Object * -ExcludeProperty _toExtract                Write-HostEx "$($matches[1]): $($matches[5])" -If ($flags.vb -ge 2) -ForegroundColor Gray            }            elseif($_ -match 'Error: .*')            {                Write-HostEx $matches[0] -ForegroundColor Red -If ($flags.vb -ge 0)                $mkvInfo.Tracks = $mkvInfo.Tracks | Select-Object * -ExcludeProperty _ExtractPath                $err = $true            }            elseif($_ -match '^\(mkvextract\)')            {                Write-Verbose $_            }            elseif($_.Trim())            { Write-HostEx $_ -ForegroundColor Gray -If ($flags.vb -ge 0) }         }

        Write-HostEx "Done.`n" -ForegroundColor Green -If ($flags.vb -ge 2 -and !$err)

    } else { Write-HostEx "No attachments to extract" -ForegroundColor Gray -If ($flags.vb -ge -1) }
    return $mkvInfo
}

function ExtractChapters([PSCustomObject]$mkvInfo, [string]$pattern, [string]$outDir, [string[]]$types, [int]$vb)
{
    if (!$outDir) {$outDir = $mkvInfo.Path | Split-Path -Parent }
    $types | %{
    
        $patternVars = @{
        '$f' = [System.IO.Path]::GetFileNameWithoutExtension($mkvInfo.Path)
        '$n' = [string]$mkvInfo.Title
        }
      
        $patternVars.GetEnumerator() | ForEach-Object {$outFile=$pattern} { 
            $outFile = $outFile -replace [regex]::escape($_.Key), $_.Value
        }

        if($_ -eq "xml" -or $_ -eq "simple")
        {
            $outFile = "$(Join-Path $outDir $outFile).$(if($_ -eq "simple"){"txt"} else {"xml"})"
            Write-HostEx "Extracting $_ chapters into `'$outFile`'" -ForegroundColor Gray -If ($vb -ge 0)
            $out = &mkvextract --ui-language en chapters $mkvInfo.Path $(if($_ -eq "simple") { "--simple" })                if ($out[0] -match "terminate called after throwing an instance")            {                Write-HostEx "Error: mkvextract terminated in an unusual way. Make sure the input file exists and is readable." -ForegroundColor Red -If ($vb -ge 0)                $err = $true            }            elseif ($out[0] -match '\<\?xml version="[0-9].[0-9]"\?\>')            {                ([xml]$out).Save($outFile)            }            elseif ($out[0] -match 'CHAPTER[0-9]+=')            {                    [string[]]$out = $out | ?{$_.Trim()}                    Set-Content -LiteralPath $outFile -Encoding UTF8 $out            }            elseif($out.Trim())            { Write-HostEx $_ -ForegroundColor Gray -If ($vb -ge 0) }            Write-HostEx "Done.`n" -ForegroundColor Green -If ($vb -ge 0 -and !$err)            } elseif ($_ -ne "none") {               Write-HostEx "Error: unsupported chapter type `"$_`"." -ForegroundColor Red -If ($vb -ge 0)        }    }
    return $mkvInfo
}

Export-ModuleMember Extract-Mkv