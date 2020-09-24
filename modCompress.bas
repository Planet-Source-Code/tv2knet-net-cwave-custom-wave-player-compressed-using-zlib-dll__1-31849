Attribute VB_Name = "modCompress"
'You can use this silly enumeration for Compresion Level
Public Enum CompressionLevel
    None = 0
    Poor = 1
    Fair = 2
    Average = 3
    Normal = 4
    Good = 5
    VeryGood = 6
    Best = 7
    SuperCompressed = 8
    MaxCompression = 9
End Enum

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'CompressBytes - Compress a Bytes Buffer
'IN  - Bytes - Bytes Array
'      Level - Compression Level to use
'OUT - Nothing
Public Sub CompressBytes(Bytes() As Byte, level As Integer)
    Dim BuffSize As Long
    Dim TBuff() As Byte
    
    BuffSize = UBound(Bytes) + 1
    BuffSize = BuffSize + (BuffSize * 1.01) + 12
    ReDim TBuff(BuffSize)
    
    compress2 TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1, level
    
    ReDim Bytes(BuffSize - 1)
    
    CopyMemory Bytes(0), TBuff(0), BuffSize
End Sub

'UnCompressBytes - Uncompresses a Byte Buffer to original size
'IN  - Bytes - Compressed Bytes Array
'      OriginalSize - Uncompressed size of Bytes Buffer
'Out - Nothing
Public Sub UnCompressBytes(Bytes() As Byte, OriginalSize As Long)
    Dim BuffSize As Long
    Dim TBuff() As Byte
    
    BuffSize = OriginalSize
    BuffSize = BuffSize + (BuffSize * 1.01) + 12
    ReDim TBuff(BuffSize)
    
    uncompress TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1
    
    ReDim Bytes(BuffSize - 1)
    
    CopyMemory Bytes(0), TBuff(0), BuffSize
End Sub

'CompressFile - Compresses a File using CompressBytes
'IN  - Src - Source File to compress
'      Dest - Compressed Destination File
'      Level - Compression Level To Use
'OUT - Nothing
Public Sub CompressFilef(src As String, dest As String, level As Integer)
    Open src For Binary Access Read As 1
    Open dest For Binary Access Read Write As 2
    
    Dim Srcs As Long
    Srcs = LOF(1)
    ReDim buff(Srcs - 1) As Byte
    Get 1, , buff
        
    CompressBytes buff, 9
        
    Put 2, , Srcs
    Put 2, , buff
    Close
End Sub

'UnCompressFile - UnCompresses a Compressed File (duh!) using UnCompressBytes
'IN  - Src - Source File to UnCompress
'      Dest - UnCompressed Destination File
'OUT - Nothing
Public Sub UnCompressFile(src As String, dest As String)
    Open src For Binary Access Read As 1
    Open dest For Binary Access Write As 2
    
    Dim Srcs As Long
    Get 1, , Srcs
    ReDim buff(Srcs - 1) As Byte
    Get 1, , buff
        
    UnCompressBytes buff, Srcs
        
    Put 2, , buff
    Close
End Sub

