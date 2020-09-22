VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H0000C0C0&
   Caption         =   "EYE Clinic Zone:- Shyam Singh Chandel"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   -1950
   ClientWidth     =   15240
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   0
      ScaleHeight     =   9495
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   1220
      Width           =   3015
      Begin Project1.USStyle USStyle5 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   7080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Picture         =   "FrmMain.frx":0CCA
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Project1.USStyle USStyle4 
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   5400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OT/PARI TYPE"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Picture         =   "FrmMain.frx":19A4
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Project1.USStyle USStyle3 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Doctors Entry"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Picture         =   "FrmMain.frx":1FC6
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Project1.USStyle USStyle2 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Test / Visits"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Picture         =   "FrmMain.frx":2736
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin Project1.USStyle USStyle1 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Registration"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         ForeColor       =   16711680
         Picture         =   "FrmMain.frx":3010
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   1
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   9255
         Left            =   15
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2880
      End
   End
   Begin MSComDlg.CommonDialog ComD 
      Left            =   480
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6600
      Top             =   9360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S O F T W A R E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "S O F T W A R E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
   Begin VB.Image Image5 
      Height          =   1215
      Left            =   -120
      Picture         =   "FrmMain.frx":3CEA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15375
   End
   Begin VB.Image Image4 
      Height          =   7905
      Left            =   3960
      Picture         =   "FrmMain.frx":974D4
      Top             =   1800
      Width           =   10515
   End
   Begin VB.Image Image2 
      Height          =   9675
      Left            =   0
      Picture         =   "FrmMain.frx":A23DC
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   15300
   End
   Begin VB.Menu logout 
      Caption         =   "Log Out"
   End
   Begin VB.Menu b 
      Caption         =   "Admin"
      Begin VB.Menu USER 
         Caption         =   "Creat User"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu a 
      Caption         =   ""
   End
   Begin VB.Menu entries 
      Caption         =   "Entries"
      Begin VB.Menu adentries 
         Caption         =   "Registration"
         Shortcut        =   ^E
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu billEntry 
         Caption         =   "Test and Visits"
         Shortcut        =   ^B
      End
      Begin VB.Menu M 
         Caption         =   "-"
      End
      Begin VB.Menu docentry 
         Caption         =   "Doctors Entry"
      End
      Begin VB.Menu n 
         Caption         =   "-"
      End
      Begin VB.Menu otentry 
         Caption         =   "OT Entry"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View By Date"
      Visible         =   0   'False
      Begin VB.Menu allrelorder 
         Caption         =   "All Relies Orders"
         Shortcut        =   {F9}
      End
      Begin VB.Menu withrono 
         Caption         =   "With R.O.No."
      End
      Begin VB.Menu withoutrono 
         Caption         =   "Without R.O.No"
      End
      Begin VB.Menu viewads 
         Caption         =   "View Ads Date Wise"
      End
      Begin VB.Menu todayads 
         Caption         =   "ToDay's Ads"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu viewalldatesrecord 
      Caption         =   "View of All Dates"
      Visible         =   0   'False
      Begin VB.Menu all 
         Caption         =   "All"
         Shortcut        =   {F11}
         Visible         =   0   'False
      End
      Begin VB.Menu withronoall 
         Caption         =   "With R.O.No"
      End
      Begin VB.Menu withoutronoall 
         Caption         =   "Without R.O.No"
      End
   End
   Begin VB.Menu settings 
      Caption         =   "Settings"
      Begin VB.Menu CompName 
         Caption         =   "Set Company Name"
      End
      Begin VB.Menu i 
         Caption         =   "-"
      End
      Begin VB.Menu systemSett 
         Caption         =   "System Settings"
      End
      Begin VB.Menu g 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Color Scheme"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help?"
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
   Begin VB.Menu e 
      Caption         =   ""
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'////                                                                    /////
'////     Developer: Shyam Singh Chandel                                 /////
'////     shyamschandel@ rediffmail.com                                  /////
'////     shyamschandel@developerssourcecode.com                         /////
'////     Programmer, System, Hardware and Electronic Engineer           /////
'////     URL http://www.developerssourcecode.com                        /////
'////                                                                    /////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////



Dim DB As Database, Rs As Recordset
Public logdb
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Sub HideWithAnim()
On Error Resume Next
Dim hw As Long, f As RECT, t As RECT
    hw = Me.hWnd
    Call GetWindowRect(hw, f)
    Call OffsetRect(f, (Screen.Width / Screen.TwipsPerPixelX - (f.Right - f.Left)) \ 2, (Screen.Height / Screen.TwipsPerPixelY - (f.Bottom - f.Top)) \ 2)
    With t
        .Left = 3 / 4 * (Screen.Width / Screen.TwipsPerPixelX)
        .Right = .Left + (f.Right - f.Left) \ 3
        .Top = 2 / 3 * (Screen.Height / Screen.TwipsPerPixelY)
        .Bottom = .Top + (f.Bottom - f.Top) \ 3
    End With
    Call DrawAnimatedRects(hw, 3, f, t)
    Call CopyRect(f, t)
    Call SetRectEmpty(t)
    Call DrawAnimatedRects(hw, 3, f, t)
    MyRs.Close
    MyDb.Close
    
End Sub
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub adentries_Click()
FrmRegistration.Show
End Sub

Private Sub all_Click()
FrmAll.Show
End Sub

Private Sub allrelorder_Click()
FrmDispROAll.Show
End Sub

Private Sub billEntry_Click()
FrmTestVisit.Show
End Sub

Private Sub color_Click()
FrmColorSet.Show vbModal
'ComD.ShowColor
'Call SaveSetting("USEYE", "Settings", "Back Color", ComD.color)
End Sub

Private Sub CompName_Click()
Dim lst, ast
lst = InputBox("Enter Company Name", "Enter Company Name")
If lst <> "" Then
Call SaveSetting("USEYE", "Settings", "Comp_Name", lst)
Else
Exit Sub
End If
ast = InputBox("Enter Company Address", "Enter Company Address")
If ast <> "" Then
Call SaveSetting("USEYE", "Settings", "Comp_Add", ast)
Else
Exit Sub
End If
End Sub

Private Sub docentry_Click()
'inp = InputBox("Your ID Please")
'If Not inp = "bansara" Then
'MsgBox "Wrong ID"
'Exit Sub
'Else
'inc = InputBox("Your Password Please")
'If Not inc = "clinic" Then
'MsgBox "Wrong Password"
'Exit Sub
'End If
'End If

FrmAddDoctor.Show

End Sub

Private Sub exit_Click()
INS = MsgBox("Do you want to Quit?", vbQuestion + vbYesNo)
If INS = vbYes Then
mIDLE.Terminate
HideWithAnim
End
Else
Exit Sub
End If

End Sub

Private Sub Form_Load()
Set DB = OpenDatabase(MainPath + "\DATA\USROUTINE.mdb")
        Set Rs = DB.OpenRecordset("AROUTINE")
        If Rs.RecordCount = 0 Then
        USStyle1.Enabled = False
        'USStyle2.Enabled = False
        adentries.Enabled = False
        End If
        
On Error Resume Next
         
   
        
Me.Caption = "EYE Clinic Zone :- " & "(LOGED USER IS: " & UCase(USER_LOG) & ")"
    mIDLE.IDLE = 120 ' seconds
    mIDLE.Init Me.hWnd, 10
 CName = GetSetting("USEYE", "Settings", "Comp_Name")
 CAddress = GetSetting("USEYE", "Settings", "Comp_Add")
 Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
 
 Label1.Caption = CName
 Label2.Caption = CAddress
 
 If CName = "" Then
 CompName_Click
 Else
 Exit Sub
 End If
 
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Image1.Left = Me.Width - Image1.Width - 100
Image2.Left = 0
Image2.Top = 1220
Image2.Height = Me.Height - 2000
Image2.Width = Me.Width
Image4.Top = Image2.Top + 600
Image4.Left = Me.Width / 2 - Image4.Width / 3
Image4.Top = Me.Height / 2 - Image4.Height / 2
Label1.Width = Me.Width - 300
Label2.Width = Me.Width - 200
Image5.Width = Me.Width + 100
Picture1.Height = Me.Height - 1800
USStyle1.Height = Picture1.Height / 5.6
USStyle2.Height = USStyle1.Height
USStyle2.Top = USStyle1.Top + USStyle1.Height + 100
USStyle3.Height = USStyle1.Height
USStyle3.Top = USStyle2.Top + USStyle2.Height + 100
USStyle4.Height = USStyle1.Height
USStyle4.Top = USStyle3.Top + USStyle3.Height + 100
USStyle5.Height = USStyle1.Height
USStyle5.Top = USStyle4.Top + USStyle4.Height + 100
Shape1.Height = USStyle5.Top + USStyle5.Height + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)

exit_Click
Cancel = True

End Sub

Private Sub logout_Click()
FrmSplash.Show
Me.Hide

End Sub

Private Sub otentry_Click()
'inp = InputBox("Your ID Please")
'If Not inp = "abc" Then
'MsgBox "Wrong ID"
'Exit Sub
'Else
'inc = InputBox("Your Password Please")
'If Not inc = "123" Then
'MsgBox "Wrong Password"
'Exit Sub
'End If
'End If
FrmAddOT.Show
End Sub

Private Sub systemSett_Click()
FrmPer.Show

End Sub

Private Sub Timer1_Timer()
If Me.Width <= 15000 Then
MsgBox "Software Required resolution of Desktop 1024 X 768 for best view of main screen picture of software and others "
Timer1.Enabled = False
End If

End Sub

Private Sub todayads_Click()
FrmToDayAdd.Show
End Sub

Private Sub USER_Click()
FrmCreatUser.Show
End Sub

Private Sub USStyle1_Click()
FrmRegistration.Show
End Sub

Private Sub USStyle2_Click()
FrmTestVisit.Show
End Sub

Private Sub USStyle3_Click()
FrmAddDoctor.Show
End Sub

Private Sub USStyle4_Click()
FrmAddOT.Show
End Sub

Private Sub USStyle5_Click()
End

End Sub

Private Sub viewads_Click()
FrmViewAdd.Show
End Sub

Private Sub withoutrono_Click()
FrmDispRONI.Show
End Sub

Private Sub withoutronoall_Click()
FrmWithoutRO.Show
End Sub

Private Sub withrono_Click()
FrmDispRO.Show

End Sub

Private Sub withronoall_Click()
FrmWithRO.Show
End Sub
