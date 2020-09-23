VERSION 5.00
Begin VB.Form mouse_chaser 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mouse chaser"
   ClientHeight    =   6390
   ClientLeft      =   360
   ClientTop       =   840
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6510
   Begin VB.Menu closefrm 
      Caption         =   "Close this window"
   End
End
Attribute VB_Name = "mouse_chaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub closefrm_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse_chaser.Cls
FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
FillStyle = 0
Circle (Int(X - 50), Int(Y - 50)), Int(200)
Sleep 125 'api responsible of the delay
End Sub

