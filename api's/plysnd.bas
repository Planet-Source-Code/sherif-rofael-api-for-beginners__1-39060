Attribute VB_Name = "plysnd"
'this program is designed by sherif rofael
'the declarations and function are collected from vairous _
 places ,
' u can use any part of that program freely ,
' I hope u find the code easy to follow
'mailto: sherif@vbcode.tk
'website: www.vbcode.tk



'play sound
'******************************
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_SYNC = &H0         '  play synchronously (default)
'**********************************

'computer name
'***********************************
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'***********************************



'capture active window
'************************
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'**********************

