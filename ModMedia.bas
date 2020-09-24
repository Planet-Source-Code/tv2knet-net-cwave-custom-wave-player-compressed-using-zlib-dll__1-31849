Attribute VB_Name = "ModMedia"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndstopsound Lib "winmm.dll" Alias "sndplaysounda" (ByVal anull As Long, ByVal flags As Integer) As Integer
Public Const SND_Sync = &H0 'wait till sound ends
Public Const SND_async = &H1 'wait till sound starts
Public Const SND_nodefault = &H2 'if not found no default sound
Public Const SND_memory = &H4 'Sound from memory
Public Const SND_loop = &H8 'loop
Public Const SND_nostop = &H10 'don't stop to play another



