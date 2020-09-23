VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   375
   ClientTop       =   855
   ClientWidth     =   7380
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7380
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11033
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "windows operation"
      TabPicture(0)   =   "frmmain.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cdd"
      Tab(0).Control(1)=   "winoperations(19)"
      Tab(0).Control(2)=   "winoperations(18)"
      Tab(0).Control(3)=   "winoperations(17)"
      Tab(0).Control(4)=   "winoperations(16)"
      Tab(0).Control(5)=   "winoperations(15)"
      Tab(0).Control(6)=   "winoperations(14)"
      Tab(0).Control(7)=   "winoperations(13)"
      Tab(0).Control(8)=   "winoperations(0)"
      Tab(0).Control(9)=   "winoperations(1)"
      Tab(0).Control(10)=   "winoperations(2)"
      Tab(0).Control(11)=   "winoperations(3)"
      Tab(0).Control(12)=   "winoperations(4)"
      Tab(0).Control(13)=   "winoperations(5)"
      Tab(0).Control(14)=   "winoperations(6)"
      Tab(0).Control(15)=   "winoperations(7)"
      Tab(0).Control(16)=   "winoperations(8)"
      Tab(0).Control(17)=   "winoperations(9)"
      Tab(0).Control(18)=   "winoperations(10)"
      Tab(0).Control(19)=   "winoperations(11)"
      Tab(0).Control(20)=   "winoperations(12)"
      Tab(0).Control(21)=   "url"
      Tab(0).Control(22)=   "webbrowser"
      Tab(0).Control(23)=   "winoperations(20)"
      Tab(0).Control(24)=   "mywebsite"
      Tab(0).Control(25)=   "winoperations(21)"
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Useful special paths"
      TabPicture(1)   =   "frmmain.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OPEN_FOLDER"
      Tab(1).Control(1)=   "opt4specialpath(0)"
      Tab(1).Control(2)=   "opt4specialpath(1)"
      Tab(1).Control(3)=   "opt4specialpath(2)"
      Tab(1).Control(4)=   "opt4specialpath(3)"
      Tab(1).Control(5)=   "opt4specialpath(4)"
      Tab(1).Control(6)=   "opt4specialpath(5)"
      Tab(1).Control(7)=   "opt4specialpath(6)"
      Tab(1).Control(8)=   "opt4specialpath(7)"
      Tab(1).Control(9)=   "opt4specialpath(8)"
      Tab(1).Control(10)=   "opt4specialpath(9)"
      Tab(1).Control(11)=   "opt4specialpath(29)"
      Tab(1).Control(12)=   "opt4specialpath(28)"
      Tab(1).Control(13)=   "opt4specialpath(27)"
      Tab(1).Control(14)=   "opt4specialpath(26)"
      Tab(1).Control(15)=   "opt4specialpath(25)"
      Tab(1).Control(16)=   "opt4specialpath(24)"
      Tab(1).Control(17)=   "opt4specialpath(23)"
      Tab(1).Control(18)=   "opt4specialpath(22)"
      Tab(1).Control(19)=   "opt4specialpath(21)"
      Tab(1).Control(20)=   "opt4specialpath(20)"
      Tab(1).Control(21)=   "opt4specialpath(19)"
      Tab(1).Control(22)=   "opt4specialpath(18)"
      Tab(1).Control(23)=   "opt4specialpath(17)"
      Tab(1).Control(24)=   "opt4specialpath(16)"
      Tab(1).Control(25)=   "opt4specialpath(15)"
      Tab(1).Control(26)=   "opt4specialpath(14)"
      Tab(1).Control(27)=   "opt4specialpath(13)"
      Tab(1).Control(28)=   "opt4specialpath(12)"
      Tab(1).Control(29)=   "opt4specialpath(11)"
      Tab(1).Control(30)=   "opt4specialpath(10)"
      Tab(1).Control(31)=   "thepath"
      Tab(1).Control(32)=   "Label1"
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Disk Information"
      TabPicture(2)   =   "frmmain.frx":0E7A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame5 
         Caption         =   "COMPUTER NAME"
         Height          =   855
         Left            =   3600
         TabIndex        =   73
         Top             =   2760
         Width           =   3135
         Begin VB.Label computername 
            Caption         =   "####################"
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "DISK INFORMATION"
         Height          =   2055
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   3255
         Begin VB.DriveListBox choosedrive 
            Height          =   315
            Left            =   360
            TabIndex        =   70
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label fre 
            Caption         =   "##########################"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label cap 
            Caption         =   "##########################"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "save active window's shot"
         Height          =   975
         Left            =   3600
         TabIndex        =   65
         Top             =   480
         Width           =   3255
         Begin VB.CommandButton saveactive 
            Caption         =   "save image to file"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   3015
         End
         Begin VB.CommandButton browseeforactive 
            Caption         =   "browse"
            Height          =   255
            Left            =   2400
            TabIndex        =   67
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox fileactiveshot 
            Height          =   285
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "vote"
         Height          =   1695
         Left            =   240
         TabIndex        =   63
         Top             =   2640
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "vote for the program"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "the above link will open the program's page on http://planet-source-code.com/"
            Height          =   495
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   2895
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "screen shot"
         Height          =   975
         Left            =   3600
         TabIndex        =   59
         Top             =   1560
         Width           =   3255
         Begin VB.CommandButton actionsave 
            Caption         =   "save the image to file"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   3015
         End
         Begin VB.CommandButton browseee 
            Caption         =   "browse"
            Height          =   255
            Left            =   2400
            TabIndex        =   61
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox filescreen 
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton OPEN_FOLDER 
         Caption         =   "OPEN FOLDER"
         Height          =   375
         Left            =   -72960
         TabIndex        =   58
         Top             =   5520
         Width           =   1935
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "SYSTEM "
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "WINDOWS"
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   55
         Top             =   880
         Width           =   1215
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "PROGRAMS"
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   54
         Top             =   1280
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "MY DOCUMENTS"
         Height          =   495
         Index           =   3
         Left            =   -74760
         TabIndex        =   53
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "FAVORITES"
         Height          =   495
         Index           =   4
         Left            =   -74760
         TabIndex        =   52
         Top             =   2080
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "STARTUP"
         Height          =   495
         Index           =   5
         Left            =   -74760
         TabIndex        =   51
         Top             =   2480
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "RECENT"
         Height          =   495
         Index           =   6
         Left            =   -74760
         TabIndex        =   50
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "SENDTO"
         Height          =   495
         Index           =   7
         Left            =   -74760
         TabIndex        =   49
         Top             =   3240
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "START MENU"
         Height          =   495
         Index           =   8
         Left            =   -74760
         TabIndex        =   48
         Top             =   3680
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "MYMUSIC"
         Height          =   495
         Index           =   9
         Left            =   -74760
         TabIndex        =   47
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "COMMON"
         Height          =   495
         Index           =   29
         Left            =   -73080
         TabIndex        =   46
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "WINDOWS SYSTEM "
         Height          =   495
         Index           =   28
         Left            =   -71520
         TabIndex        =   45
         Top             =   3480
         Width           =   2295
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "USER DIRECTORY"
         Height          =   495
         Index           =   27
         Left            =   -71520
         TabIndex        =   44
         Top             =   3120
         Width           =   2415
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "MY PICTURES"
         Height          =   495
         Index           =   26
         Left            =   -73080
         TabIndex        =   43
         Top             =   3720
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "PROGRAM FILES"
         Height          =   495
         Index           =   25
         Left            =   -71520
         TabIndex        =   42
         Top             =   2400
         Width           =   2775
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "WINDOWS SYSTEM"
         Height          =   495
         Index           =   24
         Left            =   -71520
         TabIndex        =   41
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "WINDOWS"
         Height          =   495
         Index           =   23
         Left            =   -71520
         TabIndex        =   40
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "APPLICATION DATA FOR ALL USERS"
         Height          =   495
         Index           =   22
         Left            =   -71520
         TabIndex        =   39
         Top             =   960
         Width           =   3135
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "HISTORY"
         Height          =   495
         Index           =   21
         Left            =   -71520
         TabIndex        =   38
         Top             =   2040
         Width           =   2415
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "COOKIES"
         Height          =   495
         Index           =   20
         Left            =   -73080
         TabIndex        =   37
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "TEMPORARY INTERNET FILES"
         Height          =   495
         Index           =   19
         Left            =   -71520
         TabIndex        =   36
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "LOCAL SETTING'S APP DATA."
         Height          =   495
         Index           =   18
         Left            =   -71520
         TabIndex        =   35
         Top             =   1320
         Width           =   3375
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "PRINT HOOD"
         Height          =   495
         Index           =   17
         Left            =   -73080
         TabIndex        =   34
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "APPLICATION DATA"
         Height          =   495
         Index           =   16
         Left            =   -71520
         TabIndex        =   33
         Top             =   3840
         Width           =   2895
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "ALL USERS DESKTOP"
         Height          =   495
         Index           =   15
         Left            =   -73080
         TabIndex        =   32
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "ALL USERS START UP"
         Height          =   495
         Index           =   14
         Left            =   -73080
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "TEMPLATES"
         Height          =   495
         Index           =   13
         Left            =   -73080
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "FONTS"
         Height          =   495
         Index           =   12
         Left            =   -73080
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "NETHOOD"
         Height          =   495
         Index           =   11
         Left            =   -73080
         TabIndex        =   28
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton opt4specialpath 
         Caption         =   "DESKTOP"
         Height          =   495
         Index           =   10
         Left            =   -73080
         TabIndex        =   27
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "save the image of the screen to a bmp file "
         Height          =   495
         Index           =   21
         Left            =   -71520
         TabIndex        =   26
         Top             =   5640
         Width           =   2655
      End
      Begin VB.TextBox thepath 
         Height          =   375
         Left            =   -74280
         TabIndex        =   25
         Top             =   4920
         Width           =   5415
      End
      Begin VB.CheckBox mywebsite 
         Caption         =   "VISIT MY WEBSITE."
         Height          =   495
         Left            =   -74400
         TabIndex        =   24
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "play sound"
         Height          =   495
         Index           =   20
         Left            =   -74400
         TabIndex        =   23
         Top             =   5640
         Width           =   2655
      End
      Begin VB.CommandButton webbrowser 
         Caption         =   "open the web page"
         Height          =   375
         Left            =   -69960
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox url 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74400
         TabIndex        =   21
         Text            =   "http://www.planet-source-code.com"
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "log off windows"
         Height          =   735
         Index           =   12
         Left            =   -74400
         TabIndex        =   13
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "normal size"
         Height          =   735
         Index           =   11
         Left            =   -70080
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "minimize form"
         Height          =   735
         Index           =   10
         Left            =   -71520
         TabIndex        =   11
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "maximize form"
         Height          =   735
         Index           =   9
         Left            =   -72960
         TabIndex        =   10
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "delete file"
         Height          =   735
         Index           =   8
         Left            =   -74400
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "left button for click,dblclick"
         Height          =   735
         Index           =   7
         Left            =   -70080
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "right button for click,dblclick"
         Height          =   735
         Index           =   6
         Left            =   -71520
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "dektop transparent"
         Height          =   735
         Index           =   5
         Left            =   -72960
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "toggle num lock"
         Height          =   735
         Index           =   4
         Left            =   -74400
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "toggle caps lock"
         Height          =   735
         Index           =   3
         Left            =   -70080
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "Empty Recycle bin"
         Height          =   735
         Index           =   2
         Left            =   -71520
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "show cursor"
         Height          =   735
         Index           =   1
         Left            =   -72960
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "hide cursor"
         Height          =   735
         Index           =   0
         Left            =   -74400
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "reboot"
         Height          =   735
         Index           =   13
         Left            =   -72960
         TabIndex        =   14
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "shuts down windows"
         Height          =   735
         Index           =   14
         Left            =   -71520
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "open cd tray"
         Height          =   735
         Index           =   15
         Left            =   -70080
         TabIndex        =   16
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "close cd tray"
         Height          =   495
         Index           =   16
         Left            =   -74400
         TabIndex        =   17
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "mouse chaser"
         Height          =   495
         Index           =   17
         Left            =   -72960
         TabIndex        =   18
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "change wallpaper"
         Height          =   495
         Index           =   18
         Left            =   -71520
         TabIndex        =   19
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton winoperations 
         Caption         =   "open web browser"
         Height          =   495
         Index           =   19
         Left            =   -70080
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog cdd 
         Left            =   -68640
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "PATH:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   5040
         Width           =   495
      End
   End
   Begin VB.Menu aboo 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngReturn As Long
Dim strReturn As Long

Private Sub aboo_Click()
MsgBox "this program is designed by sherif rofael,declarations and functions are collected from various places , you are allowed to take any part of that program in your code " & vbCrLf & "mailto: sherif@vbcode.com , website:http://www.vbcode.tk", vbInformation, "About the author"
End Sub

Private Sub actionsave_Click()
On Error GoTo choosefile:
Call keybd_event(vbKeySnapshot, 0, 0, 0)
SavePicture Clipboard.GetData(vbCFBitmap), filescreen.Text
Exit Sub
choosefile:
MsgBox "please choose a file to save the image to .", vbCritical, "Error "
End Sub

Private Sub browseee_Click()
On Error GoTo erroroccured:
cdd.Filter = "bmp files only (*.bmp)|*.bmp"
cdd.DialogTitle = "Save as ! "
cdd.ShowSave
filescreen.Text = cdd.FileName
Exit Sub
erroroccured:
MsgBox "Error Occured while saving file !", vbCritical, "savig file"
End Sub


Private Sub browseeforactive_Click()
On Error GoTo erroroccured:
cdd.Filter = "bmp files only (*.bmp)|*.bmp"
cdd.DialogTitle = "Save as ! "
cdd.ShowSave
fileactiveshot.Text = cdd.FileName
Exit Sub
erroroccured:
MsgBox "Error Occured while saving file !", vbCritical, "savig file"
End Sub

Private Sub choosedrive_Change()

thedrivechoosen = choosedrive.Drive & "\"
Call CalculateValues

End Sub

Private Sub Command1_Click()
Call RunBrowser("http://vbsherif.members.easyspace.com/psc.htm", 10, 1)
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then ShowCursor (True)
'Print KeyAscii
End Sub

Private Sub Form_Load()
Dim compname As String * 256
Call GetComputerName(compname, 256)
frmmain.Caption = "Welcome ," & Left(compname, InStr(compname, Chr(0)) - 1) & " To the Program."
computername.Caption = Left(compname, InStr(compname, Chr(0)) - 1)
thedrivechoosen = choosedrive.Drive & "\"
Call CalculateValues
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim aSKTOVOTE
aSKTOVOTE = MsgBox("Would you like to vote for this program?", vbYesNo + vbExclamation, "Please vote for this program")
If aSKTOVOTE = vbYes Then Call RunBrowser("http://vbsherif.members.easyspace.com/psc.htm", 10, 1): MsgBox "Thanks for spending time for voting to my program.", , "Thanks"
If aSKTOVOTE = vbNo Then MsgBox "Thanks for using my program . ", vbOKOnly, "Thanks": End
End Sub

Private Sub mywebsite_Click()
If mywebsite.Value = 1 Then url.Text = LCase$("HTTP://VBSHERIF.MEMBERS.EASYSPACE.COM")
If mywebsite.Value = 0 Then url.Text = "http://www.planet-source-code.com"
End Sub

Private Sub OPEN_FOLDER_Click()
Call RunBrowser(thepath.Text, 10, 1)
End Sub

Private Sub opt4specialpath_Click(Index As Integer)
Select Case Index
' TO GET THE SYSTEM FOLDER
'*************************************************
Case 0
Dim SYSTEMFOLDER As String * 256
GetSystemDirectory SYSTEMFOLDER, 256
thepath.Text = Left(SYSTEMFOLDER, InStr(SYSTEMFOLDER, Chr(0)) - 1)
'*************************************************


'TO GET THE WINDOWS FOLDER
'*************************************************
Case 1
Dim WINDOWSFOLDER As String * 256
GetWindowsDirectory WINDOWSFOLDER, 256
thepath.Text = Left(WINDOWSFOLDER, InStr(WINDOWSFOLDER, Chr(0)) - 1)
'*************************************************



'FROM THIS ITEM TO THE END OF THE CASE SELECT USING
' THE DECLARTION FUNCTION SHGetSpecialFolderLocation
' YOU CAN SEE THAT THE "WINDOWS" FOLDER CAN BE GOT _
  USING 2 DECLATION .
Case 2
thepath.Text = getSpecialFolder(&H2) 'PROGRAMS
Case 3
thepath.Text = getSpecialFolder(&H5) 'MY DOCUMENTS
Case 4
thepath.Text = getSpecialFolder(&H6) 'FAVORITES
Case 5
thepath.Text = getSpecialFolder(&H7) 'STARTUP
Case 6
thepath.Text = getSpecialFolder(&H8) 'RECENT
Case 7
thepath.Text = getSpecialFolder(&H9)  'SEND TO
Case 8
thepath.Text = getSpecialFolder(&HB) 'START MENU
Case 9
thepath.Text = getSpecialFolder(&HD)  ' MY MUSIC
Case 10
thepath.Text = getSpecialFolder(&H10)  'DESKTOP
Case 11
thepath.Text = getSpecialFolder(&H13)  'NETHOOD
Case 12
thepath.Text = getSpecialFolder(&H14) 'FONTS
Case 13
thepath.Text = getSpecialFolder(&H15)  'Templates
Case 14
thepath.Text = getSpecialFolder(&H18) 'ALL USERS START UP
Case 15
thepath.Text = getSpecialFolder(&H19) 'ALL USERS DESKTOP
Case 16
thepath.Text = getSpecialFolder(&H1A) 'APPLICATION DATA
Case 17
thepath.Text = getSpecialFolder(&H1B) 'PRINT HOOD
Case 18
thepath.Text = getSpecialFolder(&H1C) 'LOCAL SEETING APP'S DATA
Case 19
thepath.Text = getSpecialFolder(&H20) ' TEMPORARY INTERNET FILES
Case 20
thepath.Text = getSpecialFolder(&H21) 'COOKIES
Case 21
thepath.Text = getSpecialFolder(&H22) 'HISTORY
Case 22
thepath.Text = getSpecialFolder(&H23) 'APP DATA FOR ALL USERS
Case 23
thepath.Text = getSpecialFolder(&H24) 'WINDOWS
Case 24
thepath.Text = getSpecialFolder(&H25) 'WINSYSTEM
Case 25
thepath.Text = getSpecialFolder(&H26) 'PROGRAM FILES
Case 26
thepath.Text = getSpecialFolder(&H27) 'MY PICTURES
Case 27
thepath.Text = getSpecialFolder(&H28) 'USER DIRECTORY
Case 28
thepath.Text = getSpecialFolder(&H29) 'WINDOWS SYSTEM
Case 29
thepath.Text = getSpecialFolder(&H2B) 'COMMON FILES


End Select


End Sub



Private Sub saveactive_Click()
On Error GoTo choosefile:
Call keybd_event(vbKeySnapshot, 1, 0, 0)
SavePicture Clipboard.GetData(vbCFBitmap), fileactiveshot.Text
Exit Sub
choosefile:
MsgBox "please choose a file to save the image to .", vbCritical, "Error "
End Sub



Private Sub SSTab1_DblClick()

End Sub

Private Sub webbrowser_Click()
On Error GoTo WEBERROR:
Call RunBrowser(url.Text, 10, 1)
Exit Sub
WEBERROR:
MsgBox "PLEASE ENTER A VALID WEB ADDRESS WITH THE 'HTTP://'  i.e. 'HTTP://WWW.BBC.COM'   ", vbCritical, "Error !"
End Sub

Private Sub winoperations_Click(Index As Integer)
Select Case Index

Case 0
ShowCursor (False)
MsgBox "To show the cursor again press 'Esc'", vbInformation, "Cursor hidied"

Case 1
ShowCursor (True)

Case 2
Dim Drive
Call MakeRecycleBinEmpty(Drive, False, False, False)

Case 3
Call keybd_event(vbKeyCapital, 0, 0, 0)

Case 4
Call keybd_event(vbKeyNumlock, 0, 0, 0)

Case 5
desktop.Visible = True

Case 6
SwapMouseButton True 'Right button will do the work of normal selection, Clicking, DoubleClicking
MsgBox "THE RIGHT BUTTON IS NOW THE BOTTOM RESPONSIBLE OF THE SELECTION , CLICK , DOUBLE CLICK", , "API !!"
Case 7
SwapMouseButton False 'Left button will do the work of normal selection, Clicking, DoubleClicking
 MsgBox "THE RIGHT LEFT IS NOW THE BOTTOM RESPONSIBLE OF THE SELECTION , CLICK , DOUBLE CLICK", , "API !!"

Case 8
Call openfiletodelete

Case 9
ShowWindow frmmain.hwnd, 3
'1 = Normal
'2 = Minimize
'3 = Maximize
'4 = Show but not Focus
'6 = Show both will Minimize
Case 10
ShowWindow frmmain.hwnd, 2
Case 11
ShowWindow frmmain.hwnd, 1

Case 12
Call ShutDownWindows(0)

Case 13
Call ShutDownWindows(2)

Case 14
Call ShutDownWindows(1)

Case 15

'To open the CD door, use this code:
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)


Case 16
'To close the CD door, use this code:
lngReturn = mciSendString("set CDAudio door closed", strReturn, 127, 0)

Case 17
mouse_chaser.Visible = True

Case 18
Dim lngSuccess As Long
Dim strBitmapImage As String
Call openimage(strBitmapImage)
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)
MsgBox "The choosen Bitamp Image is set to your desktop ", vbInformation, "Done,..."


Case 19
Call RunBrowser("http://vbsherif.members.easyspace.com/psc.htm", 10, 1)

Case 20
Dim strsound
Call playthesound(strsound)
Call sndPlaySound(strsound, SND_SYNC)

Case 21
Call keybd_event(vbKeySnapshot, 0, 0, 0)
Call savetheactive

End Select
End Sub




Public Function openfiletodelete()
Dim CHECKTODELETEFILE
cdd.Filter = "All  Files (*.*)|*.*"
cdd.DialogTitle = "Choose the file to delete ! "
cdd.ShowOpen
If cdd.FileName = "" Then MsgBox "YOU CHOOSED NOTHING TO DELETE", vbInformation, "API ": Exit Function
CHECKTODELETEFILE = MsgBox("THIS FILE [" & cdd.FileName & "] WILL BE DELETED ," & vbCrLf & "ARE YOU SURE YOU WANT TO DELETE THE FILE?", vbYesNo + vbExclamation, "API")
If CHECKTODELETEFILE = vbYes Then
DeleteFile (cdd.FileName)
End If
End Function

Public Function openimage(strBitmapImage)
On Error GoTo erroroccured:
cdd.Filter = "bmp files only (*.bmp)|*.bmp"
cdd.DialogTitle = "Choose the image to set as your desktop wallpaper  ! "
cdd.ShowOpen
strBitmapImage = cdd.FileName
Exit Function
'MsgBox "this file will be deleted", , "API"
'DeleteFile (cdd.filename)
erroroccured:
MsgBox " A file with a wrong format is choosen", vbCritical, "Error !"
End Function

Public Function playthesound(strsound)
On Error GoTo erroroccured:
cdd.Filter = "wav files only (*.wav)|*.wav"
cdd.DialogTitle = "Choose the (.wav) sound file to play  ! "
cdd.ShowOpen
strsound = cdd.FileName
Exit Function
erroroccured:
MsgBox " A file with a wrong format is choosen", vbCritical, "Error !"
End Function
Public Sub savetheactive()
Dim bitampp
On Error GoTo erroroccured:
cdd.Filter = "bmp files only (*.bmp)|*.bmp"
cdd.DialogTitle = "Save as ! "
cdd.ShowSave
bitampp = cdd.FileName
SavePicture Clipboard.GetData(vbCFBitmap), bitampp
Exit Sub
erroroccured:
MsgBox "Error Occured while saving file !", vbCritical, "savig file"
End Sub

Private Sub winoperations_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then ShowCursor (True)
End Sub









'get infos about all drives , cd-roms , floopies
'****************************************
'****************************************
Public Function CalculateValues()
Dim TempTotalBytes, TempFreeBytes, TotalBytes, FreeBytes, BytesAvailableToCaller
Dim Status, FreeSpaceValue, TotalSpaceValue, cstart
TempTotalBytes = 0
TempFreeBytes = 0
TotalBytes = 0
FreeBytes = 0
On Error Resume Next
Status = GetDiskFreeSpaceEx(thedrivechoosen, BytesAvailableToCaller, TotalBytes, FreeBytes)
If FreeBytes = 0 Then FreeSpaceValue = "Bytes"
If FreeBytes = 0 Then GoTo Confirm
If (FreeBytes * 10000 / 1024) < 1 Then GoTo FreeBytes_Bytes
If (FreeBytes * 10000 / 1024 / 1024) < 1 Then GoTo FreeBytes_Kilo
If (FreeBytes * 10000 / 1024 / 1024 / 1024) < 1 Then GoTo FreeBytes_Mega Else GoTo Giga
Exit Function
FreeBytes_Bytes:
FreeSpaceValue = "Bytes"
TempFreeBytes = FreeBytes * 10000
GoTo Confirm
FreeBytes_Kilo:
FreeSpaceValue = "Kilobytes"
TempFreeBytes = FreeBytes * 10000 / 1024
GoTo Confirm
FreeBytes_Mega:
FreeSpaceValue = "Megabytes"
TempFreeBytes = FreeBytes * 10000 / 1024 / 1024
GoTo Confirm
Giga:
FreeSpaceValue = "Gigabytes"
TempFreeBytes = FreeBytes * 10000 / 1024 / 1024 / 1024
Confirm:
If (TotalBytes * 10000 / 1024) < 1 Then GoTo TotalBytes_Bytes
If (TotalBytes * 10000 / 1024 / 1024) < 1 Then GoTo TotalBytes_Kilo
If (TotalBytes * 10000 / 1024 / 1024 / 1024) < 1 Then GoTo TotalBytes_Mega Else GoTo TotalGiga
TotalBytes_Bytes:
TotalSpaceValue = "Bytes"
TempTotalBytes = TotalBytes * 10000
GoTo ReConfirm
TotalBytes_Kilo:
TotalSpaceValue = "Kilobytes"
TempTotalBytes = TotalBytes * 10000 / 1024
GoTo ReConfirm
TotalBytes_Mega:
TotalSpaceValue = "Megabytes"
TempTotalBytes = TotalBytes * 10000 / 1024 / 1024
GoTo ReConfirm
TotalGiga:
TotalSpaceValue = "Gigabytes"
TempTotalBytes = TotalBytes * 10000 / 1024 / 1024 / 1024
ReConfirm:
cstart = InStr(TempTotalBytes, ".")
TempTotalBytes = Left(TempTotalBytes, cstart - 1) & "." & Mid(TempTotalBytes, cstart + 1, 2)
cstart = InStr(TempFreeBytes, ".")
TempFreeBytes = Left(TempFreeBytes, cstart - 1) & "." & Mid(TempFreeBytes, cstart + 1, 2)
If TempTotalBytes = 0 And TempFreeBytes = 0 Then
MsgBox "The drive isn't ready", vbCritical, "Error occured!"
End If
cap.Caption = "Drive Capacity: " & TempTotalBytes & " " & TotalSpaceValue
fre.Caption = "Free Space: " & TempFreeBytes & " " & FreeSpaceValue
End Function

'*****************************************************
'*****************************************************
