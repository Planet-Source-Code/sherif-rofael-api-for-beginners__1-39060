Attribute VB_Name = "useful_paths"
'this program is designed by sherif rofael
'the declarations and function are collected from vairous _
 places ,
' u can use any part of that program freely ,
' I hope u find the code easy to follow
'mailto: sherif@vbcode.tk
'website: www.vbcode.tk





'TO GET THE SYSTEM FOLDER
'**********************************************
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'**********************************************

'TO GET THE WINDOWS DIRECTORY
'***************************************
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'***************************************

