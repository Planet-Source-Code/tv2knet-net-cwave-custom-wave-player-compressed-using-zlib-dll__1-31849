Attribute VB_Name = "modIDv2"
Public Sub SaveIDv2(File As String, Title As String, Comments As String)
Dim T(1 To 4) As String
Dim lf As String

Dim PaT As String
PaT = App.Path
If Right$(PaT, 1) <> "\" Then
PaT = PaT + "\"
End If
lf = ReadBinFile(File)
T(1) = ReadBinFile(PaT + "edit1.dat")
T(2) = ReadBinFile(PaT + "edit2.dat")
T(3) = ReadBinFile(PaT + "edit3.dat")
T(4) = T(1) + Comments + T(2) + Title + T(3)
lf = T(4) + lf
Dim k As String
k = lf
Open File + ".cwav" For Output As #1
Print #1, k
Close #1
Form1.ComDLG.DialogTitle = "Choose file to save to..."
Form1.ComDLG.Filter = "*.cwav | *.cwav"
Form1.ComDLG.DefaultExt = "*.cwav"
Form1.ComDLG.FileName = "*.cwav"
Form1.ComDLG.ShowSave
If Form1.ComDLG.FileName = "" Then Exit Sub
CompressFilef File + ".cwav", Form1.ComDLG.FileName, 9

End Sub
Public Function ReadBinFile(ByVal FileName As String)
    Dim strFromFile As String
    Dim lngFileSize As Long
    Dim FileNum As Integer
    FileNum = FreeFile
    lngFileSize = FileLen(FileName)
    strFromFile = String(lngFileSize, " ")
    Open FileName For Binary As FileNum
    Get FileNum, , strFromFile
    Close FileNum
    ReadBinFile = strFromFile
End Function
