﻿<#

.SYNOPSIS
Get-MkvInfo runs mkvinfo to get information about the contents of a Matroska file and formats it into an object for further processing.

.DESCRIPTION

Get-MkvInfo takes a Matroska (*.mkv, *.mka, *.mks) as an input and returns a custom object containing
general information about the file as well as a list of tracks and a list of attachments.
The returned object also exposes a number of Methods to filter the track and attachment lists.

Get-MkvInfo uses CodecId.xml and FourCC.xml to provide user friendly video/audio/subtitle codec information.

The Module acts as a wrapper for the mkvtoolnix command line tools and therefore requires 
your PATH environment variable to point to mkvinfo.exe

.EXAMPLE
Get-MkvInfo X:\Video.mkv

.NOTES
Segment linking is not supported at this time.

.LINK
https://github.com/line0/MkvTools

#>
#requires -version 3

function Get-Mkvinfo
{
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)] [string]$file)
    
    Check-CmdInPath mkvinfo.exe -Name mkvtoolnix

    try { $file = Get-Files $file -match '.mk[v|a|s]$' -matchDesc Matroska }
    catch
    {
        if($_.Exception.WasThrownFromThrowStatement -eq $true)
        {
            Write-Host $_.Exception.Message -ForegroundColor Red
            break
        }
        else {throw $_.Exception}
    }

    $mkvInfo = (&mkvinfo  --ui-language en $file) | ? {$_.trim() -ne "" }
 
    $i=0
    $mkvInfoFilt = @()
    foreach ($line in $mkvInfo)
    {
        #Remove somes sizes and lengths from the nodes
        $line = $line -replace '^([| ]*\+ )(.*?)(?:[,:]? \(?(?:size|length):? [0-9]+\)?)(?: \((.*?)\))?','$1$2$3'
        $regex = '^([| ]*)\+ (.*?)(?:\: (.*?))?(?:$)'
        $matches = select-string -InputObject $line -pattern $regex  | Select -ExpandProperty Matches
         
        [int]$depth = $matches.Groups[1].Length #get tree depth from indentation
        $prop = $matches.Groups[2].Value
        $val = $matches.Groups[3].Value

        $mkvInfoFilt += ,@($i, $depth, $prop, $val)
        $i++
    }

    $maxDepth = $mkvInfoFilt | %{$_[1]} | Measure-Object -Maximum
    [array]::reverse($mkvInfoFilt)

    for([int]$i=$maxDepth.Maximum; $i -ge 0; $i--)
    {
        $lines = @($mkvInfoFilt | ?{$_[1] -eq $i}) + ,@(-1,-1,-1) # additional dummy line required to parse last line

        Remove-Variable lastLine -ErrorAction SilentlyContinue
        foreach($line in $lines)
        {
            if(!$lastLine) { $arr = @() }

            elseif(($lastLine[0]-1) -ne $line[0]) #non-consecutive line indexes mean we've found our parent node
            {
                $idx = [array]::IndexOf(($mkvInfoFilt | %{$_[0]}),$lastLine[0]-1)
                [array]::reverse($arr)
                $mkvInfoFilt[$idx][3] = $arr
                $arr = @()
            }

            if($line[0] -ne -1) #skip dummy entries
            {
                if (!$arr.($line[2])) #if a sibling node with the same name doesn't exist...
                { 
                    $hash = @{$line[2] = ,@($line[3])} # create new hashtable and add array with value of the current line as first element
                    $arr += $hash

                }
                else
                {
                    $arr[-1].Set_Item($line[2],(,@($line[3]))+$arr[-1].($line[2])) # else add value as new array item to the existing hashtable
                }
                $lastLine = $line
            }

        }
        $mkvInfoFilt = $mkvInfoFilt | ?{$_[1] -eq 0 -or $_[1] -ne $i}    # remove processed lines except the parent nodes 
        0..($mkvInfoFilt.Length-1) |%{$mkvInfoFilt[$_][0]=$mkvInfoFilt.Length-1-$_} # make line indexes continuous again
    }

    $segments = $mkvInfoFilt | ?{ $_[2] -eq "Segment"} | %{$_[3]}
    $tracks = if ($segments.("Segment tracks")) { $segments.("Segment tracks")[0]["A track"] } else { $segments.("Tracks")[0]["Track"] }

    $segmentInfo = [PSCustomObject]@{
        UID = [byte[]]$(,@($segments.("Segment information").("Segment UID")[0] -split "\s" | %{[byte]$_}))
        Duration = [TimeSpan]($segments.("Segment information").("Duration")[0] -creplace "(?:.*?s \()?([0-9]*:[0-9]{2}:[0-9]{2}.[0-9]{3,7})[0-9]*\)?","`$1")
        TrackCount = $tracks.Length
        Path = $file
    }

    if($segments.("Segment information").("Title")) 
        { Add-Member -InputObject $segmentInfo -Name Title -Value $segments.("Segment information").("Title")[0] -MemberType NoteProperty }

    $codecIDs = [xml](Get-Content (Join-Path (Split-Path -parent $PSCommandPath) "CodecID.xml"))`
              | %{$_.codecs.codec}
    $fourCCs = [xml](Get-Content (Join-Path (Split-Path -parent $PSCommandPath) "FourCC.xml"))`
              | %{$_.FourCCs.FourCC}

    $tracksInfo = @()
    foreach ($track in $tracks)
    {
        $trackId = [int]($track.("Track number")[0] -creplace "[0-9]+ \(track ID for mkvmerge \& mkvextract\: ([0-9]*)\)","`$1")
        $trackType = $track.("Track type")[0]
        $codecInfo = $codecIDs | ?{$_.id -eq $track.("Codec ID")[0]}
        if ($track.("CodecPrivateFourCC")) 
            { $fourCC = $fourCCs | ?{$_.code -eq ($track.("CodecPrivateFourCC")[0]).Substring(0,4)} }

        $trackInfo = [PSCustomObject]@{
            ID = $trackId
            Type = $trackType
            CodecID = $track.("Codec ID")[0]
        }

        $trackInfo = AddNoteProperties -InputObject $trackInfo -Properties `                     @(
                        @{Name="CodecName"; Value=$codecInfo.name},
                        @{Name="CodecDesc"; Value=$codecInfo.desc},
                        @{Name="Name";      Value=$track.("Name")},
                        @{Name="Language";  Value=$track.("Language")},
                        @{Name="CodecExt";  Value=$codecInfo.ext},
                        @{Name="Enabled";   Value=$track.("Enabled"); Type=[type]"bool"},
                        @{Name="Forced";    Value=if ($track.("Forced flag")) { $track.("Forced flag") } else { $track.("Forced track flag") } ; Type=[type]"bool"},
                        @{Name="Lacing";    Value=$track.("Lacing flag"); Type=[type]"bool"}
                        @{Name="Default";   Value=if ($track.("Default flag")) { $track.("Default flag") } else { $track.("Default track flag") }; Type=[type]"bool"}
                      )

        if($trackType -eq "video")
        {
            $trackInfo = AddNoteProperties -InputObject $trackInfo -Properties `                @(
                    @{Name="dResX";       Value=$track.("Video track").("Display width"); Type=[type]"int"},
                    @{Name="dResY";       Value=$track.("Video track").("Display height"); Type=[type]"int"},
                    @{Name="pResX";       Value=$track.("Video track").("Pixel width"); Type=[type]"int"},
                    @{Name="pResY";       Value=$track.("Video track").("Pixel height"); Type=[type]"int"},
                    @{Name="Interlaced";  Value=$track.("Video track").("Interlaced"); Type=[type]"bool"},
                    @{Name="Framerate";   Value=$track.("Default duration"); RepFrom='(?:[0-9]*.[0-9]*ms|(?:[0-9]{2}:)*[0-9]{2}.[0-9]*) \(([0-9]*.[0-9]*) frames.*?\)'; RepTo='$1'; Type=[type]"float"},
                    @{Name="FourCC";      Value=$fourCC.code},
                    @{Name="FourCCName";  Value=$fourCC.name},
                    @{Name="FourCCDesc";  Value=$fourCC.desc},
                    @{Name="Profile";     Value=$track.("CodecPrivateh.264 profile")}
                 )
        }

        if($trackType -eq "audio")
        {
            $trackInfo = AddNoteProperties -InputObject $trackInfo -Properties `                @(
                    @{Name="SampleRate";   Value=$track.("Audio track").("Sampling Frequency"); Type=[type]"int"},
                    @{Name="ChannelCount"; Value=$track.("Audio track").("Channels"); Type=[type]"int"},
                    @{Name="BitDepth";     Value=$track.("Audio track").("Bit Depth"); Type=[type]"int"},
                    @{Name="FormatTag";    Value=$track.("CodecPrivateformat tag")}
                )
        }

        $tracksInfo += $trackInfo
    }
    $segmentInfo = $segmentInfo | Add-Member -MemberType NoteProperty -Name Tracks -Value $tracksInfo -PassThru

    
    if($segments.("Attachments")) 
    { 
        $attsInfo = @()
        #$atts = $segments.("Attachments").("Attached")
        $atts = $segments.("Attachments")[0]["Attached"] # because the above line apparently resolves array too far when there's only one element
   
        foreach ($att in $atts)
        {
            $attInfo = [PSCustomObject]@{
                UID = [uint64]$att.("File UID")[0]
                MimeType = $att.("Mime type")[0]
                Name = $att.("File Name")[0]
                }
            $attsInfo += $attInfo
        }
        $segmentInfo = $segmentInfo | Add-Member -MemberType NoteProperty -Name Attachments -Value $attsInfo -PassThru
    }


    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][string]$type) 
        return $this.Tracks | ?{ $_.Type -match $type }
    } -Name GetTracksByType -PassThru

    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][int]$id) 
        return $this.Tracks | ?{ $_.ID -eq $id }
    } -Name GetTracksById -PassThru

    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][hashtable]$props) 

    $props.GetEnumerator() | %{$tracksSel=$this.Tracks}{
        $key=$_.Key
        $val=$_.Value
        if ($val.GetType().Name -eq [type]"Regex") { $tracksSel = $tracksSel | ?{$_.$key -match $val} }
        else { $tracksSel = $tracksSel | ?{$_.$key -eq $val} }
    }
    return $tracksSel

    } -Name GetTracksByProperties -PassThru

    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][string]$ext) 
        return $this.Attachments | ?{ $_.Name -match ".($ext)`$"}
    } -Name GetAttachmentsByExtension -PassThru

    $segmentInfo = $segmentInfo | Add-Member -MemberType ScriptMethod -Value `
    { param([Parameter(Mandatory=$true)][uint64]$uid) 
        return $this.Attachments | ?{ $_.UID -eq $uid }
    } -Name GetAttachmentsByUID -PassThru

    return $segmentInfo
}

function AddNoteProperties([object]$InputObject, $Properties)
{
    $Properties | ?{ $_.Value }| %{
        $type = if($_.Type) { $_.Type } else { [type]"string" }
        $val = if($_.RepFrom -and $_.RepTo) {$_.Value -replace $_.RepFrom,$_.RepTo} else {$_.Value}
        if($_.Type -eq [type]"bool") {$val = [int]$val} # Convert "0" -> $false
        $val = $val -as $type
        Add-Member -InputObject $InputObject -NotePropertyName $_.Name  -NotePropertyValue $val             
    }
    return $InputObject
}

Export-ModuleMember Get-Mkvinfo