Attribute VB_Name = "exitwindows"
'this program is designed by sherif rofael
'the declarations and function are collected from vairous _
 places ,
' u can use any part of that program freely ,
' I hope u find the code easy to follow
'mailto: sherif@vbcode.tk
'website: www.vbcode.tk








' TO PERFORM SOME DELAY TO A CERTAIN ACTION
'***********************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'***********************************************

' SHUT DOWN . LOGS OFF, REBOOT YOUR WINDOWS
'***************************************************
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EXIT_LOGOFF = 0
Public Const EXIT_SHUTDOWN = 1
Public Const EXIT_REBOOT = 2
'***************************************************
'CHANGES YOUR WALLPAPER
'***************************************
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20
'***************************************

' STARTS YOUR SCREEN SAVER
'***************************************
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const comWM_SYSCOMMAND = &H112&
Const cSC_SCREENSAVE = &HF140&
'***************************************


'open web browser (internet explorer)
'******************************************
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
'*******************************************

Public Sub ShutDownWindows(ByVal uFlags As Long)
  Call ExitWindowsEx(uFlags, 0)
End Sub



'*******************************************
Public Sub RunBrowser(strURL As String, iWindowStyle As Integer, fH As Long)
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)
End Sub
'*******************************************

