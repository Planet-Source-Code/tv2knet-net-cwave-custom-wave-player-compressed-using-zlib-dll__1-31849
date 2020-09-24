VERSION 5.00
Begin VB.UserControl Media 
   BackColor       =   &H00CB565F&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FontTransparent =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not all the code is created by ftp software... The binarie file read is from another author. Credit must go to him"
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim X As String
Event WavErr(ErrorNum As Integer)
'Default Property Values:
Const m_def_Title = ""
Const m_def_Comments = ""
'Const m_def_Title = ""
'Const m_def_Comments = ""
'Property Variables:
Dim m_Title As String
Dim m_Comments As String
'Dim m_Title As String
'Dim m_Comments As String





'Sub AddWaveToMem(FileName As String, MemId As String)
'Dim SBuf As String 'Sound Buffer
'Dim chunk As String 'Amount to be read at a time
'Dim FF As Integer
'On Error GoTo WavMemErr
'chunk = Space$(1024)
'SBuf = ""
'FF = FreeFile
'Open FileName For Binary As FF
'Do While Not EOF(FF)
'    Get FF, , chunk 'Load in chunk Interval
'    SBuf = SBuf & chunk
'Loop
'Close FF
'MemId = Trim$(SBuf) 'Assign file to string
'Exit Sub
'WavMemErr:
'    RaiseEvent WavErr(1)
'End Sub

Public Sub TvMPlay(File As String)
LoadIntoMEM File
Dim i As Long
'If Len(Replace(X, "RIFF", " ", 1, Len(X), vbBinaryCompare)) <> Len(X) Then
'Exit Sub
'End If


For i = 1 To Len(X)

If Mid$(X, i, 4) = "RIFF" Then
i = i
GoTo 10
End If

Next
10
X = Mid$(X, i, Len(X) - i + 1)
Dim PaT As String
PaT = App.Path
If Right$(PaT, 1) <> "\" Then PaT = PaT + "\"
SaveTemp X, PaT + "temp.wav"
sndPlaySound PaT + "temp.wav", 1

End Sub

Public Sub SaveTemp(Data As String, File As String)
Open File For Output As #1
Print #1, Data
Close #1
End Sub

Sub PlayWav(FileName As String, Optional PlayType As Integer = 1, Optional NoDef As Boolean = True)
On Error GoTo WavPlayErr
Select Case PlayType
    Case 1
        PlayType = SND_async
    Case 2
        PlayType = SND_Sync
    Case 3
        PlayType = SND_nodefault
    Case 4
        PlayType = SND_loop
    Case 5
        PlayType = SND_nostop
End Select
If NoDef = True Then
    sndPlaySound FileName, PlayType Or SND_nodefault
Else
    sndPlaySound FileName, PlayType
End If
WavPlayErr:
    RaiseEvent WavErr(3)
End Sub


Sub StopWavs()
On Error GoTo StopErr
sndstopsound 0, 0
StopErr:
    RaiseEvent WavErr(4)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get Title() As String
'    Title = m_Title
'End Property
'
'Public Property Let Title(ByVal New_Title As String)
'    m_Title = New_Title
'    PropertyChanged "Title"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get Comments() As String
'    Comments = m_Comments
'End Property
'
'Public Property Let Comments(ByVal New_Comments As String)
'    m_Comments = New_Comments
'    PropertyChanged "Comments"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Title = m_def_Title
'    m_Comments = m_def_Comments
    m_Title = m_def_Title
    m_Comments = m_def_Comments
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Title = PropBag.ReadProperty("Title", m_def_Title)
'    m_Comments = PropBag.ReadProperty("Comments", m_def_Comments)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    m_Comments = PropBag.ReadProperty("Comments", m_def_Comments)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
'    Call PropBag.WriteProperty("Comments", m_Comments, m_def_Comments)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("Comments", m_Comments, m_def_Comments)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,1,
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Title As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Title = New_Title
    PropertyChanged "Title"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,1,
Public Property Get Comments() As String
    Comments = m_Comments
End Property

Public Property Let Comments(ByVal New_Comments As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Comments = New_Comments
    PropertyChanged "Comments"
End Property

'---------
'---------


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

Public Function GetMP3v2Comments() As String
Dim i As Long
Dim t As String
Dim l As Long
Dim y As String
Dim q As Long
For i = 1 To 300 'Don't know if this is correct
y = Mid$(X, i, 4)
If y = "COMM" Then
For q = i + 11 To 300
t = Mid$(X, q, 4)
l = l + 1
'Debug.Print t
If t = "TPE1" Then
y = Mid$(X, i + 11 + 4, l - 1 - 4) ' Str(i)
Debug.Print y + " <--" + vbCrLf

GoTo 10
End If
If t = "TCON" Then
y = Mid$(X, i + 11 + 4, l - 1 - 4) ' Str(i)
Debug.Print y + " <--" + vbCrLf

GoTo 10
End If
Next
'GoTo 10
End If

Next
GoTo 20
10
GetMP3v2Comments = y
Exit Function
20
GetMP3v2Comments = ""
End Function

Public Function GetMP3v2Title() As String
Dim i As Long
Dim y As String

For i = 1 To 300 'Don't know if this is correct
y = Mid$(X, i, 4)
If y = "TIT2" Then
y = Mid(X, i + 11, 1000)
y = GetToHeader(y)
' Str(i)
GoTo 10
End If

Next
GoTo 20
10
GetMP3v2Title = y
Exit Function
20
GetMP3v2Title = ""
End Function
Public Function GetToHeader(MP3Data As String) As String
Dim i As Long
For i = 1 To Len(MP3Data)
If Mid$(MP3Data, i, 4) = "RIFF" Then
GetToHeader = Mid$(MP3Data, 1, i - 3)
Exit Function
End If

Next
GetToHeader = MP3Data
End Function
Public Function GetMP3v2Year() As String
Dim i As Long
Dim t As String
Dim l As Long
Dim y As String
Dim q As Long
For i = 1 To 300 'Don't know if this is correct
y = Mid$(X, i, 4)
If y = "TYER" Then
For q = i + 11 To 300
t = Mid$(X, q, 4)
l = l + 1
If t = "TALB" Then
y = Mid$(X, i + 11, l - 1) ' Str(i)

GoTo 10
End If

Next
'GoTo 10
End If

Next
GoTo 20
10
GetMP3v2Year = y
Exit Function
20
GetMP3v2Year = ""
End Function

Public Sub LoadIntoMEM(MP3File As String)
X = ReadBinFile(MP3File)
If Left$(X, 4) = "RIFF" Then X = ""
m_Title = GetMP3v2Title

m_Comments = GetMP3v2Comments



End Sub

Public Sub ClearMEM()
X = ""
End Sub

