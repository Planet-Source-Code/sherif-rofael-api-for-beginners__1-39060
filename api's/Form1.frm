VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Code:
Private Sub Command1_Click()
Call ShowWindow(kk, SW_HIDE)
End Sub
Private Sub Command2_Click()
Call ShowWindow(kk, SW_NORMAL)
End Sub

Private Sub Command3_Click()
f = FindWindow("Shell_TrayWnd", "")
Call ShowWindow(f, SW_NORMAL)

End Sub

Private Sub Command4_Click()
f = FindWindow("Shell_TrayWnd", "")
Call ShowWindow(f, SW_HIDE)

End Sub

Private Sub Form_Load()
kk = WindowFromPoint(42, 590)

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
xx = x
yy = y
End Sub
Private Sub Timer1_Timer()
Call GetCursorPos(l)
xx = l.x
yy = l.y
a = WindowFromPoint(xx, yy)

Call SetWindowPos(1900612, 0, 0, 0, 0, 0, 5)
'Label1.Caption = CStr(l.x)
End Sub

