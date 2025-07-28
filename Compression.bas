Attribute VB_Name = "Compression"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function z2Compress Lib "sxpcompq.dll" Alias "BZ2_bzBuffToBuffCompress" (Dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function z2Decompress Lib "sxpcompq.dll" Alias "BZ2_bzBuffToBuffDecompress" (Dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long

Public Function CompressData(TheData() As Byte, ByVal lCompressionLevel As Long) As Long
Dim BufferSize As Long, TempBuffer() As Byte, Result As Long, lSourceLen As Long

If lCompressionLevel > 9 Then lCompressionLevel = 9
If lCompressionLevel < 0 Then lCompressionLevel = 1
lSourceLen = UBound(TheData) + 1
BufferSize = lSourceLen + (lSourceLen * 0.01) + 600
ReDim TempBuffer(BufferSize)
Result = z2Compress(TempBuffer(0), BufferSize, TheData(0), lSourceLen, lCompressionLevel, 0, 0)
ReDim Preserve TheData(BufferSize - 1)
CopyMemory TheData(0), TempBuffer(0), BufferSize
Erase TempBuffer
If Result = 0 Then CompressData = UBound(TheData) + 1
CompressData = Result
End Function

Public Function DeCompressData(TheData() As Byte, lDestLen As Long) As Long
Dim TempBuffer() As Byte, Result As Long, lSourceLen As Long, lVerbosity As Long ' We want the DLL to shut up, so set it to 0
Dim lSmall As Long ' if <> 0 then use (s)low memory routines
lVerbosity = 0
lSmall = 0
lSourceLen = UBound(TheData) + 1
ReDim TempBuffer(lDestLen - 1)
Result = z2Decompress(TempBuffer(0), lDestLen, TheData(0), lSourceLen, lSmall, lVerbosity)
ReDim Preserve TheData(lDestLen - 1)
CopyMemory TheData(0), TempBuffer(0), lDestLen
Erase TempBuffer
If Result = 0 Then DeCompressData = UBound(TheData) + 1
DeCompressData = Result
End Function

