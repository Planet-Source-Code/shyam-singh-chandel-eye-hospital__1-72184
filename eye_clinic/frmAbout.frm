VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4830
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   9390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0CCA
   ScaleHeight     =   4830
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   2520
   End
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   65535
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16744576
      ColorButtonUp   =   16711680
      ColorButtonDown =   16761024
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image Image2 
      Height          =   1065
      Left            =   1320
      Picture         =   "frmAbout.frx":195D84
      Top             =   105
      Width           =   5880
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   9360
      X2              =   9360
      Y1              =   0
      Y2              =   5520
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5520
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   9360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   9360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Win98, Win2000, WinME, XP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   8
      Top             =   1440
      Width           =   3810
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   " Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email: - shyamschandel@rediffmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Platform: -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "400 MHz and above Processor,  64 MB RAM and above   MSAccess 2000 and above"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   4770
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Requirment: -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   2985
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contect us: -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
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

Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub



Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
'Me.Visible = False
End Sub

Private Sub USStyle1_Click()
Unload Me

End Sub
