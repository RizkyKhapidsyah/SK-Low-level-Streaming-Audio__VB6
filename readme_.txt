---------------------------------------------------------------------
The information in this article applies to:

 - Microsoft Visual Basic Learning, Professional, and Enterprise Editions
   for Windows, version 5.0
---------------------------------------------------------------------

SUMMARY
=======

LLStream.exe is a sample that contains a Visual Basic project that
demonstrates how to play streaming audio files by calling several low-level
functions in the Windows API. This sample helps you create a program that
plays audio files continually without slowing down the operating system.

MORE INFORMATION
================

The sample uses several low-level functions in the Windows API to work
around the Multimedia Control Interface. You must be familiar with using
Visual Basic to call the Windows API functions in order to understand how
this sample works.

The following file is available for download from the Microsoft Software
Library:

 ~ LLStream.exe


Release Date: June-01-1998

For more information about downloading files from the Microsoft Software
Library, please see the following article in the Microsoft Knowledge Base:

   ARTICLE-ID: Q119591
   TITLE : How to Obtain Microsoft Support Files from Online Services

The sample uses several low-level functions in the Windows API to work
around the Multimedia Control Interface. You must be familiar with using
Visual Basic to call the Windows API functions to understand how this
sample works.

When you run the self-extracting executable file, the following files are
expanded into the Low-Level Streaming Audio Sample Project directory:

 - Form1.frm (3Kb) - the main form
 - Module1.bas (12Kb) - the module containing the function declarations
 - Project1.vbp (1Kb) - the main project file
 - Project1.vbw (1Kb) - the project workspace
 - Readme.txt - overview of the sample and download information

The Multimedia Control Interface (MCI) is the one way to add multimedia
capabilities to your program in Visual Basic. However, when you use the MCI
to play audio files continuously, you can slow down or lock the operating
system.

This sample shows how to call the appropriate functions directly to work
around the limitations of the MCI. Because the sample hooks into a Windows
process, you should compile this sample and run the executable file rather
than running the sample from the Visual Basic design environment.

The sample uses the following functions in the Windows API:

   GlobalAlloc
   GlobalFree
   GlobalLock
   mmioAscend
   mmioClose
   mmioDescend
   mmioDescendParent
   mmioOpen
   mmioRead
   mmioReadString
   mmioSeek
   mmioStringToFOURCC
   PostWavMessage
   RtlMoveMemory
   waveOutClose
   waveOutGetDevCaps
   waveOutGetErrorText
   waveOutGetNumDevs
   waveOutGetPosition
   waveOutOpen
   waveOutPrepareHeader
   waveOutReset
   waveOutUnprepareHeader
   waveOutWrite
