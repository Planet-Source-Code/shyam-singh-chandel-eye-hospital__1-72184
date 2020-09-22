VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   7440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7560
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   360
      X2              =   7440
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   1010
      Width           =   11895
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   1560
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   7800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6360
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "shyamschandel@rediffmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2865
      TabIndex        =   5
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   330
      TabIndex        =   4
      Top             =   330
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contect us :-        "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1200
      Left            =   7980
      Picture         =   "FrmSplash.frx":0CCA
      Top             =   5160
      Width           =   5310
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "shyamschandel@rediffmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by :-        "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   345
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   1560
      Picture         =   "FrmSplash.frx":1598C
      Top             =   1815
      Width           =   4575
   End
End
Attribute VB_Name = "FrmSplash"
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
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim ColorBack, INS
Dim SPATH
Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
MainPath = GetSetting("INIWINAGE", "LOCATION", "APPNON")
If MainPath = "" Then
INS = App.Path & "\data"
Call SaveSetting("INIWINAGE", "LOCATION", "APPNON", INS)
MainPath = GetSetting("INIWINAGE", "LOCATION", "APPNON")
End If

CreateDirectory MainPath
CreateDirectory MainPath & "\data\"
'CreateDirectory MainPath & "\Schedule\"
CreateDirectory MainPath & "\user\"
Call USERDB(MainPath & "\user\user.mdb")
Call MAINDB(MainPath & "\DATA\USECZ.mdb")
Call ROUTINE(MainPath & "\DATA\USROUTINE.mdb")
End Sub

Private Sub Form_Unload(Cancel As Integer)
HideWithAnim
End Sub

Private Sub Timer1_Timer()
FrmLogin.Show
 'MAINDB (MainPath & "\DATA\USECZ.mdb")
'Unload Me
Timer1.Enabled = False
End Sub
