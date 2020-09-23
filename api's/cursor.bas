Attribute VB_Name = "Some"
'this program is designed by sherif rofael
'the declarations and function are collected from vairous _
 places ,
' u can use any part of that program freely ,
' I hope u find the code easy to follow
'mailto: sherif@vbcode.tk
'website: www.vbcode.tk

' all this functions is collected from various places
' and i didn't invent any .


Option Explicit
'hide the mouse cursor
'******************************************
Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
'******************************************

'open & close cd tray
'***********************************
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'***********************************

'toggle the caps lock (turns it on when it's off and turns if off when it's on)
'********************************
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'********************************

'desktop transparent
'*******************************
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
'*******************************

' change mouse buttons configurations
'*******************************************
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'*********************************************

'delete any file
'*************************************
 Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'*************************************

'empty recycle bin
'******************************************
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'*************************************************


'change form's state
'***********

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'***************************


Public Sub MakeRecycleBinEmpty(Optional ByVal Drive As String, Optional NoConfirmation As Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
 Dim hwnd, Flags As Long
 On Error Resume Next
 hwnd = Screen.ActiveForm.hwnd
 If Len(Drive) > 0 Then _
  Drive = Left$(Drive, 1) & ":\"
 Flags = (NoConfirmation And &H1) Or (NoProgress And &H2) Or (NoSound And &H4)
 SHEmptyRecycleBin hwnd, Drive, Flags
End Sub
'******************************************



