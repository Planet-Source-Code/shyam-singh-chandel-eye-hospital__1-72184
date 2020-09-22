VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmColorSet 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solor Scheme"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "FrmColorSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Label Font Color"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Frame Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Default Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "List Font Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -120
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Report List Font Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5520
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Report List Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Background Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Schem"
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
      Height          =   930
      Left            =   420
      TabIndex        =   9
      Top             =   225
      Width           =   4560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Schem"
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
      Height          =   930
      Left            =   450
      TabIndex        =   8
      Top             =   240
      Width           =   6000
   End
End
Attribute VB_Name = "FrmColorSet"
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



Private Sub Command1_Click()
CD1.ShowColor
Command1.BackColor = CD1.color
Call SaveSetting("USEYE", "Settings", "Back Color", CD1.color)
Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")

End Sub

Private Sub Command2_Click()
CD1.ShowColor
Command2.BackColor = CD1.color
Call SaveSetting("USEYE", "Settings", "List Color", CD1.color)

End Sub

Private Sub Command3_Click()
CD1.ShowColor
Command3.BackColor = CD1.color
Label1.BackColor = CD1.color
Call SaveSetting("USEYE", "Settings", "Font Color", CD1.color)
If Label1.BackColor = vbWhite Then
Label1.ForeColor = vbBlack
Else
Label1.ForeColor = vbWhite
End If

End Sub

Private Sub Command4_Click()
Call SaveSetting("USEYE", "Settings", "List Color", vbWhite)
Call SaveSetting("USEYE", "Settings", "Back Color", &HC0E0FF)
Call SaveSetting("USEYE", "Settings", "Font Color", vbBlack)
Call SaveSetting("USEYE", "Settings", "Frame Color", &H80C0FF)
Form_Load

End Sub

Private Sub Command5_Click()
CD1.ShowColor
Command5.BackColor = CD1.color
Call SaveSetting("USEYE", "Settings", "Frame Color", CD1.color)

End Sub

Private Sub Form_Load()
On Error Resume Next
Command1.BackColor = GetSetting("USEYE", "Settings", "Back Color")
Command2.BackColor = GetSetting("USEYE", "Settings", "List Color")
Command3.BackColor = GetSetting("USEYE", "Settings", "Font Color")
Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
Command5.BackColor = GetSetting("USEYE", "Settings", "Frame Color")

Label1.BackColor = GetSetting("USEYE", "Settings", "Font Color")
If Label1.BackColor = vbWhite Then
Label1.ForeColor = vbBlack
Else
Label1.ForeColor = vbWhite
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'FrmMain.BackColor = GetSetting("USEYE", "Settings", "Back Color")
End Sub
