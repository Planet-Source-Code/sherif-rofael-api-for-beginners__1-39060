VERSION 5.00
Begin VB.Form desktop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "desktop"
   ClientHeight    =   6105
   ClientLeft      =   450
   ClientTop       =   930
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6915
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   480
   End
   Begin VB.Menu closethis 
      Caption         =   "close this window"
   End
End
Attribute VB_Name = "desktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub closethis_Click()
Unload Me
End Sub
Private Sub Timer1_Timer()
PaintDesktop desktop.hdc
End Sub
