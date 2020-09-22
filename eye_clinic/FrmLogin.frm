VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6960
      Top             =   3600
   End
   Begin Project1.USStyle USStyle3 
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle2 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Quit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Help ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "."
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   4200
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   120
      Picture         =   "FrmLogin.frx":0CCA
      Top             =   120
      Width           =   3450
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2535
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   8400
      TabIndex        =   12
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Note:- Dafault User Name= us and Password=us only for first time. when user created then it will't work."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "FrmLogin"
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


Dim MyDb As Database, MyRs As Recordset
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
Dim Exp
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

Private Sub Command1_Click()
On Error Resume Next
Dim USER, Pass

SQL = "SELECT * FROM USERDB WHERE UserName='" & Text1.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(MainPath & "\user\USER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
USER = MyRs!UserName
Pass = MyRs!Pass
USER_LOG = USER
If Text1.Text = "" Then
MsgBox "Please Enter the User Name."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please Enter the Password."
Text2.SetFocus
Exit Sub
End If

If Text1.Text = USER And Text2.Text = Pass Then
HideWithAnim
FrmMain.Show
Unload FrmSplash
Unload Me
Exit Sub
Else

MsgBox "WRONG USER NAME AND PASSWORD"

Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If


MyDb.Close

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next

If Text1.Text = "" Then
MsgBox "Please Enter the User Name."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please Enter the Password."
Text2.SetFocus
Exit Sub
End If

If Text1.Text = "us" And Text2.Text = "us" Then
USER_LOG = "us"
HideWithAnim
FrmMain.Show
Unload FrmSplash
Unload Me
Else
MsgBox "Wrong User Name or Password"

Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Exit Sub
End If

End Sub

Private Sub Command4_Click()
FrmHelp2.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Exp = GetSetting("10009", "Settings", "Expiry")
If Exp = "" Then
Call SaveSetting("10009", "Settings", "Expiry", "0")
End If

Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
Me.Top = FrmSplash.Height - Me.Height - 1000
Me.Left = FrmSplash.Left + 1600

Set MyDb = DBEngine.Workspaces(0).OpenDatabase(MainPath & "\user\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)

If MyRs.RecordCount <= 0 Then
Command3.Visible = True
USStyle3.Visible = True
logdb = "No"
Exit Sub
Else
Command3.Visible = False
USStyle3.Visible = False
logdb = "Yes"

End If
 MyRs.Close
 MyDb.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Text2.SetFocus
   End If
   
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If logdb = "Yes" Then
 Command1_Click
 Else
  Command3_Click
 End If
 End If
 
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
Dim d, d2
d2 = "060906"
If Exp >= d2 Then
MsgBox "Time of Trial is over"
End
End If
d = Format(Date, "ddmmyy")
If d >= d2 Then
Call SaveSetting("10009", "Settings", "Expiry", d)

MsgBox "Time of Trial is over"
End
End If

End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Unload Me
Command2_Click
End Sub

Private Sub USStyle3_Click()
Command3_Click
End Sub

