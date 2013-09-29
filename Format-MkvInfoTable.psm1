<#
.SYNOPSIS
Format-MkvInfoTable outputs a formatted table with important information about 
the tracks and attachments of a Matroska file.

.DESCRIPTION

Format-MkvInfoTable takes objects returned by Get-MkvInfo and lets you outputs a configurable set of tables.
Being a filter it is designed to receive objects from the pipe
and will stream formatted tables as the objects are coming in.

Setting _ExtractStateTracks (tracks) or _ExtractState (attachments) to [int]1 will set the EX flag
which highlights the row of the flagged track.
Type "Get-Help Extract-Mkv -full" and refer to the -ReturnMkvInfo parameter for more information on extraction flags.

.EXAMPLE

Get-MkvInfo X:\Video.mkv | Format-MkvInfoTable
Displays all tables (video, audio, subtitles, attachments).

.EXAMPLE

Get-MkvInfo X:\Video.mkv | Format-MkvInfoTable -Tables attachments, subtitles
Displays only the attachments and subtitle tracks in the specified order.

.LINK
https://github.com/line0/MkvTools

#>
#requires -version 3

filter Format-MkvInfoTable([string[]]$Tables=@("video","audio","subtitles","attachments"))
{
    if(!$_ -or $_ -isnot [PSCustomObject]) { throw "Input missing or not an Get-MkvInfo object" } 

    $color = [PSCustomObject]@{
        ForegroundDefault = $Host.UI.RawUI.ForegroundColor
        BackgroundDefault = $Host.UI.RawUI.BackgroundColor
        ForegroundHighlight = [ConsoleColor]"Yellow"
        BackgroundHighlight = $Host.UI.RawUI.BackgroundColor
    } | Add-Member -MemberType ScriptMethod -Name Highlight -PassThru -Value `
            {   
                param([bool]$enable=$true,[string]$retValHighlight,[string]$retValDefault) 
                if ($enable) {
                    $Host.UI.RawUI.ForegroundColor=$this.ForegroundHighlight 
                    $Host.UI.RawUI.BackgroundColor=$this.BackgroundHighlight
                    return $retValHighlight
                } else { 
                    $Host.UI.RawUI.ForegroundColor=$this.ForegroundDefault
                    $Host.UI.RawUI.BackgroundColor=$this.BackgroundDefault
                    return $retValDefault
                } 
            }    

    $mkvInfo = $_

    $Tables | %{

        # Keep these inside the loop to avoid having to clone all the hashtables to not permanently lose the order key
        $tblCommon=@(
        @{label="EX"; Expression={ $color.Highlight(($_._ExtractStateTrack -eq 1),"X") }; Width=2; Order=0},
        @{label="ID"; Expression={ $_.ID }; Width=2; Order=1},
        @{label="Track Name"; Expression={ $_.Name }; Width=35; Order=2 },
        @{label="Lang"; Expression={ $_.Language }; Width=4; Order=3},
        @{label="Flags"; Expression={ if($_.Enabled){"Enabled"}; if($_.Forced){"Forced"}; if($_.Default){"Default"} }; Width=18; Order=99}
        )
    
        $tblVideo=@(
        @{label="Codec"; Expression={ if($_.FourCCName) {"$($_.FourCCName) (VfW)"} else {"$($_.CodecName)$(if($_.Profile){", $($_.Profile)"})"} }; Width=30; Order=4 },
        @{label="Resolution"; Expression={"$($_.pResX)x$($_.pResY)$(if($_.Interlaced) {"i"} else {"p"})$($_.Framerate)"}; Width=18; Order=5}
        )

        $tblAudio=@(
        @{label="Codec"; Expression={ if($_.FormatTag) {"$($_.CodecName) (ACM)"} else {"$($_.CodecName)"} }; Width=20; Order=4 },
        @{label="SRate"; Expression={"$([float]$_.SampleRate/1000) kHz"}; Width=10; Order=5},
        @{label="Chan"; Expression={$_.ChannelCount}; Width=4; Order=6},
        @{label="Depth"; Expression={if($_.BitDepth){"$($_.BitDepth)bit"}}; Width=6; Order=7}
        )

        $tblSubs=@(
        @{label="Codec"; Expression={ $_.CodecName }; Width=20; Order=4 }
        )

        $tblAtt=@(
        @{label="EX"; Expression={ $color.Highlight(($_._ExtractState -eq 1),"X") }; Width=2; Order=0},
        @{label="UID"; Expression={ $_.UID }; Width=20; Order=1 },
        @{label="Mime Type"; Expression={$_.MimeType}; Width=30; Order=2},
        @{label="Name"; Expression={$_.Name}; Width=40; Order=3}
        )

        $tblHeader = switch ($_) { 
            "video"       { "Video tracks" }
            "audio"       { "Audio tracks" }
            "subtitles"   { "Subtitle tracks" }
            "attachments" { "Attachments" }
            }
        Write-Host $tblHeader -NoNewline -ForegroundColor DarkBlue -BackgroundColor DarkGray

        $tbl = switch ($_) {
            "video" {$tblCommon + $tblVideo}
            "audio" {$tblCommon + $tblAudio}
            "subtitles" {$tblCommon + $tblSubs}
            "attachments" {$tblAtt}
            }

        $tbl = $tbl | Sort-Object -Property @{Expression={$_.Order}}
        $tbl | %{$_.Remove("Order")}  # because Format-Table doesn't accept extra keys
    
        if($_ -eq "attachments") { $mkvInfo.Attachments | Format-Table -Property $tbl }
        else {$mkvInfo.GetTracksByType($_) | Format-Table -Property $tbl }

        $color.Highlight($false)
    }
}

Export-ModuleMember Format-MkvInfoTable