VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Wave Player V0.3"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Bugs"
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Restart"
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   3465
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Height          =   330
      Left            =   3780
      TabIndex        =   9
      Top             =   1155
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "CWAV IDv2 0.0.1"
      Height          =   2115
      Left            =   1260
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "Encode CWAV file"
         Height          =   330
         Left            =   525
         TabIndex        =   8
         Top             =   1575
         Width           =   1485
      End
      Begin VB.TextBox Com 
         Height          =   750
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "Form1.frx":0000
         Top             =   735
         Width           =   2325
      End
      Begin VB.TextBox Title 
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Text            =   "Title"
         Top             =   315
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select and play sound"
      Height          =   330
      Left            =   1470
      TabIndex        =   4
      Top             =   1155
      Width           =   2115
   End
   Begin MSComDlg.CommonDialog ComDLG 
      Left            =   105
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "cwav"
      DialogTitle     =   "Choose CWAV File To Play"
      Filter          =   "*.cwav"
   End
   Begin CWAV.Media Media1 
      Left            =   105
      Top             =   105
      _ExtentX        =   953
      _ExtentY        =   397
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Compression using zlib.dll"
      Height          =   435
      Left            =   3885
      TabIndex        =   10
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created by T-Virus Creations. "
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   4950
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the CWAV Player"
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   210
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   4110
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CWave Encoder By T-Virus Creations"
      ForeColor       =   &H00397A16&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   4095
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
ComDLG.DialogTitle = "Choose CWAV file to play..."
ComDLG.DefaultExt = "*.cwav"
ComDLG.Filter = "*.cwav files|*.cwav"
ComDLG.ShowOpen
On Error GoTo 10
If ComDLG.FileName = "" Then Exit Sub
UnCompressFile ComDLG.FileName, ComDLG.FileName + "x"
ModMedia.sndPlaySound 0, 2
ModMedia.sndPlaySound 0, 0
 
Media1.TvMPlay ComDLG.FileName + "x"
Title.Text = Media1.Title
Com.Text = Media1.Comments
Kill ComDLG.FileName + "x"
10
End Sub

Private Sub Command2_Click()

ComDLG.DialogTitle = "Choose WAV file to encode to CWAV..."
ComDLG.DefaultExt = "*.wav"
ComDLG.Filter = "*.wav files|*.wav"
ComDLG.ShowSave
If ComDLG.FileName = "" Then Exit Sub
SaveIDv2 ComDLG.FileName, Title.Text, Com.Text

End Sub

Private Sub Command3_Click()
ModMedia.sndPlaySound 0, 2
ModMedia.sndPlaySound 0, 0
 
End Sub

Private Sub Command4_Click()
Shell App.EXEName

Unload Me
End Sub

Private Sub Command5_Click()
Dim t As String
t = "Current known bugs:" + vbCrLf + vbCrLf + " "
t = t + "- Don't try to open none CWAVE files or you'll have to click reset"
t = t + vbCrLf + " - Currently does this program support IDv2 Comments and Title only!"
MsgBox t, vbInformation, "Bugs..."

End Sub

Private Sub Form_Unload(Cancel As Integer)
ModMedia.sndPlaySound 0, 2
ModMedia.sndPlaySound 0, 0
 
End Sub

