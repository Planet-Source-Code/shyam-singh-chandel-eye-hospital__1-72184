VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmCreatUser 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creat User"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "FrmCreatUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin Project1.USStyle USStyle2 
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
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
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Creat User"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Height          =   2370
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   1
      Top             =   2535
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   0
      Top             =   1725
      Width           =   2775
   End
   Begin Project1.USStyle USStyle3 
      Height          =   615
      Left            =   4440
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dalete User"
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
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Creat User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   330
      TabIndex        =   12
      Top             =   105
      Width           =   5655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Creat User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "FrmCreatUser"
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
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(MainPath & "\user\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Plaes enter the user name and password"
Exit Sub
End If

MyRs.AddNew
     MyRs!UserName = Text1.Text
     MyRs!Pass = Text2.Text
MyRs.Update
MsgBox "User has been created"

MyDb.Close
Text1 = ""
Text2 = ""
Form_Load
     
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
List1.BackColor = GetSetting("USEYE", "Settings", "List Color")
Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
List1.ForeColor = GetSetting("USEYE", "Settings", "Font Color")


List1.Clear
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(MainPath & "\user\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)
If MyRs.RecordCount = 0 Then
Exit Sub
Else
Do While Not MyRs.EOF
List1.AddItem MyRs!UserName
MyRs.MoveNext
Loop
MyDb.Close

End If
Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
HideWithAnim
End Sub

Private Sub List1_Click()
SQL = "select * from USERDB where USERNAME='" & List1.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(MainPath & "\user\USER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
     Text1.Text = MyRs!UserName
     Text2.Text = MyRs!Pass
End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Unload Me
End Sub

Private Sub USStyle3_Click()
ans = MsgBox("Are you sure to Delete User " & List1.Text & ".", vbQuestion + vbYesNo)
If ans = vbYes Then
MyRs.Delete
Text1 = ""
Text2 = ""
Else
Exit Sub
End If
Form_Load
End Sub

