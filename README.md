MkvTools
========

A small set of PowerShell modules that enable batch processing of Matroska files with  MKVToolNix.

_Note: Segment linking is not supported at this time._

Requirements
============

MkvTools requires PowerShell 3 and the following software to be available from your PATH environment variable:

* [MKVToolNix](http://www.bunkus.org/videotools/mkvtoolnix/) v6.3.0+

Installation
============

Unpack the MkvTools archive into: _%userprofile%\Documents\WindowsPowerShell\Modules_.

Usage
=====

Extract-Mkv
-----------

**Extract-Mkv batch extracts tracks, attachments, chapters and timecodes from Matroska files using the mkvtoolnix command line tools.**


Extract-Mkv accepts a comma-delimited list of input files and/or folders (recursing supported) and by default extracts all tracks from each input file using a configurable naming pattern.
It also allows you to specify which track types or track IDs to extract and lets you choose a custom output directory if you don't want it to extract into the parent directory of the input files.

Extract-Mkv can extract tracks, attachments, chapters and timecodes in one go, will indicate progress using status bars where possible and returns track/attachment tables that highlight what is being extracted. 

The Module acts as a wrapper for the mkvtoolnix command line tools and therefore **requires your PATH environment variable to point to mkvextract.exe and mkvinfo.exe**

For examples and information on command line parameters, run:

```
Get-Help Extract-Mkv -full
```

Get-MkvInfo
-----------

**Get-MkvInfo runs mkvinfo to get information about the contents of a Matroska file and formats it into an object for further processing.**

Get-MkvInfo takes a Matroska (*.mkv, *.mka, *.mks) as an input and returns a custom object containing general information about the file as well as a list of tracks and a list of attachments.
The returned object also exposes a number of Methods to filter the track and attachment lists.

Get-MkvInfo uses CodecId.xml and FourCC.xml to provide user friendly video/audio/subtitle codec information.

The Module acts as a wrapper for the mkvtoolnix command line tools and therefore **requires your PATH environment variable to point to mkvinfo.exe**


Format-MkvInfoTable
-------------------

**Format-MkvInfoTable outputs a formatted table with important information about the tracks and attachments of a Matroska file.**

Format-MkvInfoTable takes objects returned by Get-MkvInfo and lets you outputs a configurable set of tables.
Being a filter it is designed to receive objects from the pipe
and will stream formatted tables as the objects are coming in.

Setting _ExtractStateTracks (tracks) or _ExtractState (attachments) to [int]1 will set the EX flag
which highlights the row of the flagged track.
Type "Get-Help Extract-Mkv -full" and refer to the -ReturnMkvInfo parameter for more information on extraction flags.

For examples and information on command line parameters, run:

```
Get-Help Format-MkvInfoTable -full
```