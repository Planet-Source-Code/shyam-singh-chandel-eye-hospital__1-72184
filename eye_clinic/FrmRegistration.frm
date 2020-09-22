VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRegistration 
   BackColor       =   &H0080FF80&
   Caption         =   "Patient Registration and Re Registration       "
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11925
   Icon            =   "FrmRegistration.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin Project1.USStyle USStyle12 
      Height          =   255
      Left            =   2520
      TabIndex        =   115
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Refresh"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.USStyle USStyle10 
      Height          =   255
      Left            =   360
      TabIndex        =   113
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Text45 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      MaxLength       =   200
      TabIndex        =   112
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text44 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      MaxLength       =   200
      TabIndex        =   31
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text43 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      MaxLength       =   200
      TabIndex        =   32
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text42 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      MaxLength       =   200
      TabIndex        =   108
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text41 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4130
      MaxLength       =   200
      TabIndex        =   106
      Top             =   7220
      Width           =   1140
   End
   Begin VB.TextBox Text40 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      MaxLength       =   200
      TabIndex        =   105
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text39 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   50
      TabIndex        =   104
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text38 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   103
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text37 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   102
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5760
      ScaleHeight     =   615
      ScaleWidth      =   4095
      TabIndex        =   99
      Top             =   7680
      Width           =   4095
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6120
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFF00&
      Height          =   1815
      Left            =   17760
      ScaleHeight     =   1755
      ScaleWidth      =   11115
      TabIndex        =   97
      Top             =   1800
      Visible         =   0   'False
      Width           =   11175
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while Finding Record. . . ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   11055
      End
   End
   Begin Project1.USStyle USStyle8 
      Height          =   375
      Left            =   13440
      TabIndex        =   95
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Previous"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.USStyle USStyle7 
      Height          =   375
      Left            =   15000
      TabIndex        =   94
      Top             =   6720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Next"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   360
      TabIndex        =   19
      Top             =   5160
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   92
      Top             =   7530
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   22826
            MinWidth        =   22826
            Text            =   "                                                 Total Similar Records:- "
            TextSave        =   "                                                 Total Similar Records:- "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      TabIndex        =   23
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text36 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   200
      TabIndex        =   8
      Top             =   4440
      Width           =   2415
   End
   Begin Project1.USStyle USStyle5 
      Height          =   375
      Left            =   360
      TabIndex        =   88
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Refresh"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   17760
      ScaleHeight     =   1755
      ScaleWidth      =   11115
      TabIndex        =   85
      Top             =   1920
      Visible         =   0   'False
      Width           =   11175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   240
         TabIndex        =   86
         Top             =   240
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483636
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label6 
         Caption         =   "   Reg.             Others                        Name                        Age   Sex    Diagnosis     LastVisit     NextVisit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Width           =   11055
      End
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8640
      TabIndex        =   38
      Top             =   6600
      Width           =   2895
   End
   Begin VB.TextBox Text35 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   17280
      MaxLength       =   20
      TabIndex        =   40
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11640
      Top             =   6840
   End
   Begin VB.TextBox Text34 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12000
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   9480
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox Text33 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   18120
      MaxLength       =   10
      TabIndex        =   39
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text32 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      MaxLength       =   20
      TabIndex        =   36
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text31 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      MaxLength       =   10
      TabIndex        =   35
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text30 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      MaxLength       =   20
      TabIndex        =   34
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text29 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   33
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   200
      TabIndex        =   30
      Top             =   5880
      Width           =   3015
   End
   Begin VB.TextBox Text27 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10200
      MaxLength       =   100
      TabIndex        =   29
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text26 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9360
      MaxLength       =   20
      TabIndex        =   28
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text25 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      MaxLength       =   20
      TabIndex        =   27
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   26
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      MaxLength       =   50
      TabIndex        =   25
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text22 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   22
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text21 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   15600
      TabIndex        =   41
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   20
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   24
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9960
      MaxLength       =   80
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      MaxLength       =   80
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      MaxLength       =   20
      TabIndex        =   16
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   15
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      MaxLength       =   20
      TabIndex        =   14
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   13
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   12
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   11
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      MaxLength       =   200
      TabIndex        =   7
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12600
      MaxLength       =   100
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin Project1.USStyle USStyle2 
      Height          =   375
      Left            =   10080
      TabIndex        =   44
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle3 
      Height          =   375
      Left            =   6000
      TabIndex        =   43
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Register"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      PictureAlignment=   2
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
      Height          =   375
      Left            =   12720
      TabIndex        =   46
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle4 
      Height          =   615
      Left            =   14760
      TabIndex        =   47
      Top             =   10320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     Delete"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmRegistration.frx":0CCA
      PictureAlignment=   2
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   8800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   17400
      TabIndex        =   81
      Top             =   3360
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3413
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   29
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reg. No."
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   3883
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Age"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sex"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Reg"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Con"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "A/R"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "TN"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fun"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "S/L"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "IDO"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Gn"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "AScan"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Pari.Type"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Pari.Amt."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "OT Type"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "OT Amt."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Shir."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "R.B.S"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "F.F.A"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "B.T/C.T"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Diagnosis"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Medicines"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Remarks"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Balance"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Count"
         Object.Width           =   1235
      EndProperty
   End
   Begin Project1.USStyle USStyle6 
      Height          =   375
      Left            =   7920
      TabIndex        =   89
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Re Register"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle9 
      Height          =   375
      Left            =   13560
      TabIndex        =   96
      Top             =   6120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "First"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   255
      Left            =   6360
      TabIndex        =   37
      Top             =   6600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "99\/99\/99"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ListView List 
      Height          =   1695
      Left            =   360
      TabIndex        =   101
      Top             =   1800
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2990
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   39
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Registration"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Other"
         Object.Width           =   3527
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   5118
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Age"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sex"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tmp. Address"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Per.Add"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Tel"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Reg"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Con"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "A/R"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "TN"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Fun"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "S/L"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Others"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "IDO"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Gon"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "A/K"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "P. Type"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "P Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Min OT"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Major OT"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Shir"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "R.B.A"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "F.F.A"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "BT/CT"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Diagnosis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Medicines"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Remark"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Total"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Paid"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Doctor"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Object.Width           =   88
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Object.Width           =   88
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Object.Width           =   88
      EndProperty
   End
   Begin MSComctlLib.ListView List2 
      Height          =   1095
      Left            =   360
      TabIndex        =   110
      Top             =   6120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1931
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Medicines"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2118
      EndProperty
   End
   Begin Project1.USStyle USStyle11 
      Height          =   255
      Left            =   1440
      TabIndex        =   114
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Shape Shape1 
      Height          =   330
      Left            =   4080
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Index           =   35
      Left            =   3360
      TabIndex        =   111
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amt"
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
      Index           =   34
      Left            =   5880
      TabIndex        =   109
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amounts"
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
      Index           =   33
      Left            =   4080
      TabIndex        =   107
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Registration and Re Registration       "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   360
      TabIndex        =   93
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "           Major OT                  Type              Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   4920
      TabIndex        =   91
      Top             =   4755
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Address"
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
      Index           =   31
      Left            =   360
      TabIndex        =   90
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   8640
      TabIndex        =   84
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "( dd/mm/yy )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   7200
      TabIndex        =   83
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Visit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   6360
      TabIndex        =   82
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL SIMILAR RECORD FOUND:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   14760
      TabIndex        =   80
      Top             =   9480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   11880
      TabIndex        =   79
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   18120
      TabIndex        =   78
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   10440
      TabIndex        =   77
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "B.T/C.T"
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
      Index           =   24
      Left            =   9360
      TabIndex        =   76
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F.F.A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   8640
      TabIndex        =   75
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "R.B.S"
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
      Index           =   22
      Left            =   7920
      TabIndex        =   74
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IDO"
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
      Index           =   21
      Left            =   8160
      TabIndex        =   73
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Others"
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
      Index           =   20
      Left            =   7200
      TabIndex        =   72
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S/L"
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
      Index           =   19
      Left            =   6480
      TabIndex        =   71
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Temp. Address"
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
      Index           =   18
      Left            =   7200
      TabIndex        =   70
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
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
      Index           =   17
      Left            =   12600
      TabIndex        =   69
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
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
      Index           =   0
      Left            =   5760
      TabIndex        =   68
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   9240
      TabIndex        =   67
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   8040
      TabIndex        =   66
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   6960
      TabIndex        =   65
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicines"
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
      Index           =   13
      Left            =   360
      TabIndex        =   64
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis"
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
      Index           =   3
      Left            =   10200
      TabIndex        =   63
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shirmir"
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
      Index           =   12
      Left            =   7080
      TabIndex        =   62
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "           Minor OT                   Type              Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2520
      TabIndex        =   61
      Top             =   4755
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "            Peri                       Type            Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   60
      Top             =   4755
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ascan/Kerometry"
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
      Index           =   9
      Left            =   9960
      TabIndex        =   59
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gonios"
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
      Index           =   2
      Left            =   8880
      TabIndex        =   58
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fun"
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
      Index           =   8
      Left            =   5760
      TabIndex        =   57
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tn"
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
      Index           =   7
      Left            =   5040
      TabIndex        =   56
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A/R"
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
      Index           =   6
      Left            =   4320
      TabIndex        =   55
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Con"
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
      Index           =   5
      Left            =   3600
      TabIndex        =   54
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   53
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   52
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   51
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   50
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   49
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label O 
      BackStyle       =   0  'Transparent
      Caption         =   "Others"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   48
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   45
      Top             =   1140
      Width           =   12015
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu MnuRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu mnuReRegister 
         Caption         =   "Re Register"
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
   End
End
Attribute VB_Name = "FrmRegistration"
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



Public DB As Database
Public Rs As Recordset
Public Rs2 As Recordset
Public MyDb As Database
Public MyRs As Recordset
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Dim M As sPrint
Dim blnAuto As Boolean
Dim sTemp As Recordset
Public MakeTot As Integer

Public logdb
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long

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
Private Sub combo1_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not blnAuto And Combo1.Text <> "" Then
        iStart = Combo1.SelStart
        strPart = Left$(Combo1.Text, iStart)
        For iLoop = 0 To Combo1.ListCount - 1
            strItem = UCase$(Combo1.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo1.Text) Then
                blnAuto = True
                Combo1.SelText = Mid$(Combo1.List(iLoop), iStart + 1) 'add on the new ending
                Combo1.SelStart = iStart   'reset the selection
                Combo1.SelLength = Len(Combo1.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo1.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        combo1_LostFocus
        Combo1.SelStart = Len(Combo1.Text)
        If KeyCode = 13 Then
       If USStyle3.Enabled = True Then
       USStyle3_Click
       End If
       If USStyle6.Enabled = True Then
       USStyle6_Click
       End If
       
        'Exit Sub
       End If
    End If
End Sub

Private Sub combo1_LostFocus()
Dim iLoop As Integer
    If Combo1.Text <> "" Then
        For iLoop = 0 To Combo1.ListCount - 1
            If UCase$(Combo1.List(iLoop)) = UCase$(Combo1.Text) Then
                blnAuto = True
                Combo1.Text = Combo1.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub combo2_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not blnAuto And Combo2.Text <> "" Then
        iStart = Combo2.SelStart
        strPart = Left$(Combo2.Text, iStart)
        For iLoop = 0 To Combo2.ListCount - 1
            strItem = UCase$(Combo2.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo2.Text) Then
                blnAuto = True
                Combo2.SelText = Mid$(Combo2.List(iLoop), iStart + 1) 'add on the new ending
                Combo2.SelStart = iStart   'reset the selection
                Combo2.SelLength = Len(Combo2.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next

        SQL = "select * from AROUTINE where PARITYPE='" & Combo2.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text20.Text = MyRs!PariAmount
                
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo2.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        combo2_LostFocus
        Combo2.SelStart = Len(Combo2.Text)
        If KeyCode = 13 Then
        SQL = "select * from AROUTINE where PARITYPE='" & Combo2.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text20.Text = MyRs!PariAmount
        'Text20.SetFocus
        Combo3.SetFocus
        Exit Sub
       End If
    End If
End Sub

Private Sub combo2_LostFocus()
Dim iLoop As Integer
    If Combo2.Text <> "" Then
        For iLoop = 0 To Combo2.ListCount - 1
            If UCase$(Combo2.List(iLoop)) = UCase$(Combo2.Text) Then
                blnAuto = True
                Combo2.Text = Combo2.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub combo3_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not blnAuto And Combo3.Text <> "" Then
        iStart = Combo3.SelStart
        strPart = Left$(Combo3.Text, iStart)
        For iLoop = 0 To Combo3.ListCount - 1
            strItem = UCase$(Combo3.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo3.Text) Then
                blnAuto = True
                Combo3.SelText = Mid$(Combo3.List(iLoop), iStart + 1) 'add on the new ending
                Combo3.SelStart = iStart   'reset the selection
                Combo3.SelLength = Len(Combo3.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Combo3_Click()
On Error Resume Next
         SQL = "select * from AROUTINE where MINOROTTYPE='" & Combo3.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text22.Text = MyRs!MINOROTRates
              
End Sub

Private Sub combo3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo3.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        combo3_LostFocus
        Combo3.SelStart = Len(Combo3.Text)
        If KeyCode = 13 Then
        SQL = "select * from AROUTINE where MINOROTTYPE='" & Combo3.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text22.Text = MyRs!MINOROTRates
        Combo4.SetFocus
        Exit Sub
       End If
    End If
End Sub

Private Sub combo3_LostFocus()
Dim iLoop As Integer
    If Combo3.Text <> "" Then
        For iLoop = 0 To Combo3.ListCount - 1
            If UCase$(Combo3.List(iLoop)) = UCase$(Combo3.Text) Then
                blnAuto = True
                Combo3.Text = Combo3.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub combo4_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not blnAuto And Combo4.Text <> "" Then
        iStart = Combo4.SelStart
        strPart = Left$(Combo4.Text, iStart)
        For iLoop = 0 To Combo4.ListCount - 1
            strItem = UCase$(Combo4.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo4.Text) Then
                blnAuto = True
                Combo4.SelText = Mid$(Combo4.List(iLoop), iStart + 1) 'add on the new ending
                Combo4.SelStart = iStart   'reset the selection
                Combo4.SelLength = Len(Combo4.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Combo4_Click()
On Error Resume Next

SQL = "select * from AROUTINE where MAJOROTTYPE='" & Combo4.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text19.Text = MyRs!MAJOROTRates
        
End Sub

Private Sub combo4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo4.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        combo4_LostFocus
        Combo4.SelStart = Len(Combo4.Text)
        If KeyCode = 13 Then
        SQL = "select * from AROUTINE where MAJOROTTYPE='" & Combo4.Text & "'"
        Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
        Text19.Text = MyRs!MAJOROTRates
        Text23.SetFocus
        Exit Sub
       End If
    End If
End Sub

Private Sub combo4_LostFocus()
Dim iLoop As Integer
    If Combo4.Text <> "" Then
        For iLoop = 0 To Combo4.ListCount - 1
            If UCase$(Combo4.List(iLoop)) = UCase$(Combo4.Text) Then
                blnAuto = True
                Combo4.Text = Combo4.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
  MakeTot = 0
 MAINDB (MainPath & "\DATA\USECZ.MDB")
 Set DB = OpenDatabase(MainPath + "\DATA\USECZ.mdb")
 
 With Adodc1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            MainPath & "\DATA\USECZ.mdb;Persist Security Info=False"
        .RecordSource = "select Oldid, Type, Name, Age, Sex, Diagnosis, LastVisit, NextVisit from USTestVisit" ' order by Name, LastVisit"
    End With
    Set MSHFlexGrid1.DataSource = Adodc1
    MSHFlexGrid1.FormatString = "Old Reg | Others |Name |Age | Sex |Diagnosis | LastVisit | NextVisit"
    MinHeight = Me.Height
    MinWidth = Me.Width
        If Right(MainPath, 1) = "\" Then
        Set DB = OpenDatabase(MainPath + "\DATA" & "USECZ.mdb")
    Else
        Set DB = OpenDatabase(MainPath + "\DATA\USECZ.mdb")
    End If
        Set Rs = DB.OpenRecordset("USTestVisit")
    ''''''''''''''''
    On Error Resume Next
    Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
          If Right(MainPath, 1) = "\" Then
        Set MyDb = OpenDatabase(MainPath + "\DATA" & "USROUTINE.mdb")
    Else
        Set MyDb = OpenDatabase(MainPath + "\DATA\USROUTINE.mdb")
    End If
        Set MyRs = MyDb.OpenRecordset("AROUTINE")
        Do While Not MyRs.EOF
            Combo1.AddItem MyRs!Name
            If Not MyRs!PariType = "....." Then
            Resume Next
            Combo2.AddItem MyRs!PariType
            End If
            If Not MyRs!MINOROTType = "....." Then
            Resume Next
            Combo3.AddItem MyRs!MINOROTType
            End If
            If Not MyRs!MAJOROTType = "....." Then
            Resume Next
            Combo4.AddItem MyRs!MAJOROTType
            End If
            'AutoCombo1.AddItem MyRs!paritype
        MyRs.MoveNext
        Loop
        Call Form_Resize
 MSHFlexGrid1.FormatString = "Old Reg | Others |Name |Age | Sex |Diagnosis | LastVisit | NextVisit"
    
 Set M = New sPrint
Set GPrint.ListViewName = List
GPrint.DrawHorizontalLines = True
GPrint.DrawVerticalLines = True
GPrint.DrawBorder = True
GPrint.BorderDistance = 2
GPrint.PosX = 300    'Value in Twips
GPrint.PosY = 1400  'Value in Twips
GPrint.HasPicture = True
' Combo1.ListIndex = 0
'  Combo2.ListIndex = 0
'   Combo3.ListIndex = 0
'    Combo4.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    If MinHeight > Me.Height Then
        Me.Height = MinHeight
        Exit Sub
    ElseIf MinWidth > Me.Width Then
        Me.Width = MinWidth
        Exit Sub
    End If
    MSHFlexGrid1.Width = Me.ScaleWidth - 650
    MSHFlexGrid1.ColWidth(0) = MSHFlexGrid1.Width / 9
    MSHFlexGrid1.ColWidth(1) = MSHFlexGrid1.ColWidth(0)
    MSHFlexGrid1.ColWidth(2) = (MSHFlexGrid1.ColWidth(0) * 2) - 60
    MSHFlexGrid1.ColWidth(3) = MSHFlexGrid1.ColWidth(0) / 2
    MSHFlexGrid1.ColWidth(4) = MSHFlexGrid1.ColWidth(3)
    MSHFlexGrid1.ColWidth(5) = MSHFlexGrid1.ColWidth(0)
    MSHFlexGrid1.ColWidth(6) = MSHFlexGrid1.ColWidth(0)
    MSHFlexGrid1.ColWidth(7) = MSHFlexGrid1.ColWidth(0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
HideWithAnim
End Sub

Private Sub List_Click()
On Error Resume Next

Dim vI As Integer
    If List.ListItems.Count = 0 Then Exit Sub
       
    vI = List.SelectedItem.Index
    vID = List.SelectedItem
    Text1 = List.ListItems(vI).ListSubItems(1)
    Text2 = List.ListItems(vI).ListSubItems(2)
    Text3 = List.ListItems(vI).ListSubItems(3)
    Text4 = List.ListItems(vI).ListSubItems(4)
    Text5 = List.ListItems(vI).ListSubItems(5)
    Text8 = List.ListItems(vI).ListSubItems(6)
    Text36 = List.ListItems(vI).ListSubItems(7)
    Text6 = List.ListItems(vI).ListSubItems(8)
    Text9 = List.ListItems(vI).ListSubItems(9)
    Text10 = List.ListItems(vI).ListSubItems(10)
    Text11 = List.ListItems(vI).ListSubItems(11)
    Text12 = List.ListItems(vI).ListSubItems(12)
    Text13 = List.ListItems(vI).ListSubItems(13)
    Text14 = List.ListItems(vI).ListSubItems(14)
    Text15 = List.ListItems(vI).ListSubItems(15)
    Text16 = List.ListItems(vI).ListSubItems(16)
    
    Text17 = List.ListItems(vI).ListSubItems(17)
    Text18 = List.ListItems(vI).ListSubItems(18)
    Text39 = List.ListItems(vI).ListSubItems(19)
    'Combo2.Text = ""
    'Combo2.Text = List.ListItems(vI).ListSubItems(19)
    Text20 = List.ListItems(vI).ListSubItems(20)
    Text38 = List.ListItems(vI).ListSubItems(21)
    'Combo3.Text = ""
    'Combo3.Text = List.ListItems(vI).ListSubItems(21)
    Text22 = List.ListItems(vI).ListSubItems(22)
    Text37 = List.ListItems(vI).ListSubItems(23)
    'Combo4.Text = ""
    'Combo4.Text = List.ListItems(vI).ListSubItems(23)
    Text19 = List.ListItems(vI).ListSubItems(24)
    Text23 = List.ListItems(vI).ListSubItems(25)
    Text24 = List.ListItems(vI).ListSubItems(26)
    Text25 = List.ListItems(vI).ListSubItems(27)
    Text26 = List.ListItems(vI).ListSubItems(28)
    Text27 = List.ListItems(vI).ListSubItems(29)
    Text28 = List.ListItems(vI).ListSubItems(30)
    Text29 = List.ListItems(vI).ListSubItems(31)
    'Text30 = List.ListItems(vI).ListSubItems(32)
    Text31 = List.ListItems(vI).ListSubItems(33)
    'Text32 = List.ListItems(vI).ListSubItems(34)
    Text40 = List.ListItems(vI).ListSubItems(35)
    USDate = List.ListItems(vI).ListSubItems(36)
    USMonth = List.ListItems(vI).ListSubItems(37)
    USYear = List.ListItems(vI).ListSubItems(38)
    Text37.Visible = True
    Text38.Visible = True
    Text39.Visible = True
    Text40.Visible = True
    findmedicines
End Sub

Private Sub List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Select Case ColumnHeader.Index
    Case 1: SotrListView List, ColumnHeader, "Old Reg."
    Case 2: SotrListView List, ColumnHeader, "New Reg."
    Case 3: SotrListView List, ColumnHeader, "Name"
    Case 4: SotrListView List, ColumnHeader, "Address"
    Case 5: SotrListView List, ColumnHeader, "Giagnosis"
    Case 6: SotrListView List, ColumnHeader, "Last Visit"
    Case 7: SotrListView List, ColumnHeader, "Next Visit"
    
  End Select
End Sub


Private Sub List2_Click()
On Error Resume Next

Dim vI As Integer
    If List2.ListItems.Count = 0 Then Exit Sub
       
    vI = List2.SelectedItem.Index
    vID = List2.SelectedItem
    Text28 = List2.ListItems(vI).ListSubItems(1)
    Text44 = List2.ListItems(vI).ListSubItems(2)
    Text43 = List2.ListItems(vI).ListSubItems(3)
    USMedicine = Text28.Text
End Sub

Private Sub MSHFlexGrid1_Click()
USStyle4.Enabled = True
USStyle3.Enabled = False
End Sub


Private Sub Text1_GotFocus()
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = vbBlack
Text2.ForeColor = vbWhite
End Sub


Private Sub Text29_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = vbBlack
Text3.ForeColor = vbWhite
End Sub
Private Sub Text43_GotFocus()
Text43.BackColor = vbBlack
Text43.ForeColor = vbWhite
End Sub
Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text36_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text39_Click()
Text39.Visible = False
End Sub
Private Sub Text38_Click()
Text38.Visible = False
End Sub
Private Sub Text37_Click()
Text37.Visible = False
End Sub
Private Sub Text4_GotFocus()
Text4.BackColor = vbBlack
Text4.ForeColor = vbWhite
End Sub

Private Sub Text40_Click()
Text40.Visible = False
End Sub

Private Sub Text43_LostFocus()
Text43.BackColor = vbWhite
Text43.ForeColor = vbBlack
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = vbBlack
Text5.ForeColor = vbWhite
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = vbBlack
Text6.ForeColor = vbWhite
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = vbBlack
Text7.ForeColor = vbWhite
End Sub

Private Sub Text8_GotFocus()
Text8.BackColor = vbBlack
Text8.ForeColor = vbWhite
End Sub


Private Sub Text9_GotFocus()
Text9.BackColor = vbBlack
Text9.ForeColor = vbWhite
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = vbBlack
Text10.ForeColor = vbWhite
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = vbBlack
Text11.ForeColor = vbWhite
End Sub
Private Sub Text12_GotFocus()
Text12.BackColor = vbBlack
Text12.ForeColor = vbWhite
End Sub

Private Sub Text13_GotFocus()
Text13.BackColor = vbBlack
Text13.ForeColor = vbWhite
End Sub

Private Sub Text14_GotFocus()
Text14.BackColor = vbBlack
Text14.ForeColor = vbWhite
End Sub

Private Sub Text15_GotFocus()
Text15.BackColor = vbBlack
Text15.ForeColor = vbWhite
End Sub

Private Sub Text16_GotFocus()
Text16.BackColor = vbBlack
Text16.ForeColor = vbWhite
End Sub

Private Sub Text17_GotFocus()
Text17.BackColor = vbBlack
Text17.ForeColor = vbWhite
End Sub

Private Sub Text18_GotFocus()
Text18.BackColor = vbBlack
Text18.ForeColor = vbWhite
End Sub

Private Sub Text19_GotFocus()
Text19.BackColor = vbBlack
Text19.ForeColor = vbWhite
End Sub

Private Sub Text20_GotFocus()
Text20.BackColor = vbBlack
Text20.ForeColor = vbWhite
End Sub
Private Sub Text21_GotFocus()
Text21.BackColor = vbBlack
Text21.ForeColor = vbWhite
End Sub
Private Sub Text22_GotFocus()
Text22.BackColor = vbBlack
Text22.ForeColor = vbWhite
End Sub

Private Sub Text23_GotFocus()
Text23.BackColor = vbBlack
Text23.ForeColor = vbWhite
End Sub

Private Sub Text24_GotFocus()
Text24.BackColor = vbBlack
Text24.ForeColor = vbWhite
End Sub

Private Sub Text25_GotFocus()
Text25.BackColor = vbBlack
Text25.ForeColor = vbWhite
End Sub

Private Sub Text26_GotFocus()
Text26.BackColor = vbBlack
Text26.ForeColor = vbWhite
End Sub

Private Sub Text27_GotFocus()
Text27.BackColor = vbBlack
Text27.ForeColor = vbWhite
End Sub

Private Sub Text28_GotFocus()
Text28.BackColor = vbBlack
Text28.ForeColor = vbWhite
End Sub

Private Sub Text29_GotFocus()
Text29.BackColor = vbBlack
Text29.ForeColor = vbWhite
End Sub

Private Sub Text30_GotFocus()
Text30.BackColor = vbBlack
Text30.ForeColor = vbWhite
End Sub

Private Sub Text31_GotFocus()
Text31.BackColor = vbBlack
Text31.ForeColor = vbWhite
End Sub

Private Sub Text32_GotFocus()
Text32.BackColor = vbBlack
Text32.ForeColor = vbWhite
End Sub

Private Sub Text33_GotFocus()
Text33.BackColor = vbBlack
Text33.ForeColor = vbWhite
End Sub

'Private Sub Text34_GotFocus()
'Text34.BackColor = vbBlack
'Text34.ForeColor = vbWhite
'End Sub

Private Sub MaskEdBox1_GotFocus()
MaskEdBox1.BackColor = vbBlack
MaskEdBox1.ForeColor = vbWhite
End Sub
Private Sub Text36_GotFocus()
Text36.BackColor = vbBlack
Text36.ForeColor = vbWhite
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
Text1.ForeColor = vbBlack
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = vbWhite
Text2.ForeColor = vbBlack
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = vbWhite
Text3.ForeColor = vbBlack
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = vbWhite
Text4.ForeColor = vbBlack
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = vbWhite
Text5.ForeColor = vbBlack
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = vbWhite
Text6.ForeColor = vbBlack
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = vbWhite
Text7.ForeColor = vbBlack
End Sub

Private Sub Text8_LostFocus()
Text8.BackColor = vbWhite
Text8.ForeColor = vbBlack
End Sub

Private Sub Text9_LostFocus()
Text9.BackColor = vbWhite
Text9.ForeColor = vbBlack
End Sub

Private Sub Text10_LostFocus()
Text10.BackColor = vbWhite
Text10.ForeColor = vbBlack
End Sub

Private Sub Text11_LostFocus()
Text11.BackColor = vbWhite
Text11.ForeColor = vbBlack
End Sub
Private Sub Text12_LostFocus()
Text12.BackColor = vbWhite
Text12.ForeColor = vbBlack
End Sub

Private Sub Text13_LostFocus()
Text13.BackColor = vbWhite
Text13.ForeColor = vbBlack
End Sub

Private Sub Text14_LostFocus()
Text14.BackColor = vbWhite
Text14.ForeColor = vbBlack
End Sub

Private Sub Text15_LostFocus()
Text15.BackColor = vbWhite
Text15.ForeColor = vbBlack
End Sub

Private Sub Text16_LostFocus()
Text16.BackColor = vbWhite
Text16.ForeColor = vbBlack
End Sub

Private Sub Text17_LostFocus()
Text17.BackColor = vbWhite
Text17.ForeColor = vbBlack
End Sub

Private Sub Text18_LostFocus()
Text18.BackColor = vbWhite
Text18.ForeColor = vbBlack
End Sub

Private Sub Text19_LostFocus()
Text19.BackColor = vbWhite
Text19.ForeColor = vbBlack
End Sub

Private Sub Text20_LostFocus()
Text20.BackColor = vbWhite
Text20.ForeColor = vbBlack
End Sub
Private Sub Text21_LostFocus()
Text21.BackColor = vbWhite
Text21.ForeColor = vbBlack
End Sub
Private Sub Text22_LostFocus()
Text22.BackColor = vbWhite
Text22.ForeColor = vbBlack
End Sub

Private Sub Text23_LostFocus()
Text23.BackColor = vbWhite
Text23.ForeColor = vbBlack
End Sub

Private Sub Text24_LostFocus()
Text24.BackColor = vbWhite
Text24.ForeColor = vbBlack
End Sub

Private Sub Text25_LostFocus()
Text25.BackColor = vbWhite
Text25.ForeColor = vbBlack
End Sub

Private Sub Text26_LostFocus()
Text26.BackColor = vbWhite
Text26.ForeColor = vbBlack
End Sub

Private Sub Text27_LostFocus()
Text27.BackColor = vbWhite
Text27.ForeColor = vbBlack
End Sub

Private Sub Text28_LostFocus()
Text28.BackColor = vbWhite
Text28.ForeColor = vbBlack
End Sub

Private Sub Text29_LostFocus()
Text29.BackColor = vbWhite
Text29.ForeColor = vbBlack
End Sub

Private Sub Text30_LostFocus()
Text30.BackColor = vbWhite
Text30.ForeColor = vbBlack
End Sub

Private Sub Text31_LostFocus()
Text31.BackColor = vbWhite
Text31.ForeColor = vbBlack
End Sub

Private Sub Text32_LostFocus()
Text32.BackColor = vbWhite
Text32.ForeColor = vbBlack
End Sub

Private Sub Text33_LostFocus()
Text33.BackColor = vbWhite
Text33.ForeColor = vbBlack
End Sub

'Private Sub Text34_LostFocus()
'Text34.BackColor = vbWhite
'Text34.ForeColor = vbBlack
'End Sub

Private Sub MaskEdBox1_LostFocus()
MaskEdBox1.BackColor = vbWhite
MaskEdBox1.ForeColor = vbBlack
End Sub
Private Sub Text36_LostFocus()
Text36.BackColor = vbWhite
Text36.ForeColor = vbBlack
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

'On Error Resume Next
        FindsOldID Text1, List, DB, "USTestVisit", "OldID"
        
If Text3.Text = "" Then
Text2.SetFocus
USStyle1.Enabled = True
Else
USStyle1.Enabled = False
End If
If List.ListItems.Count <= 0 Then
USStyle3.Enabled = True
USStyle1.Enabled = False
USStyle6.Enabled = False
Else
USStyle3.Enabled = False
USStyle1.Enabled = True
USStyle6.Enabled = True
'Exit Sub
End If
End If
End Sub
Sub findmedicines()
MakeTot = 0
List2.ListItems.Clear
Set Rs2 = DB.OpenRecordset("select * from MEDICINE where REGNO='" & Text1.Text & "' and Date='" & USDate & "' and month='" & USMonth & "' and year='" & USYear & "'")
      Rs2.MoveFirst
      Do While Rs2.EOF = False
         Set Li = List2.ListItems.Add(, , Rs2!regno)
            Li.SubItems(1) = Rs2!MEDICINE
            Li.SubItems(2) = Rs2!QTY
            Li.SubItems(3) = Rs2!AMOUNT
            MakeTot = MakeTot + Val(Rs2!AMOUNT)
            Rs2.MoveNext
            Loop
            Text41.Text = MakeTot
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub




Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = 13 Then
 Picture2.Visible = True
 Wait (5)
 If Not Text1.Text = "" Then
Text3.SetFocus
Exit Sub
End If
 On Error Resume Next
        FindsNewID Text2, List, DB, "USTestVisit", "Type"
End If
If List.ListItems.Count <= 0 Then
USStyle3.Enabled = True
USStyle1.Enabled = False
USStyle6.Enabled = False
Else
USStyle3.Enabled = False
USStyle1.Enabled = True
USStyle6.Enabled = True
End If
Picture2.Visible = False
End Sub

Private Sub Text20_Change()
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text22_Change()
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
Picture2.Visible = True
Wait (5)
If Not Text1.Text = "" Then
Text4.SetFocus
Exit Sub

End If
 On Error Resume Next
        FindsName Text3, List, DB, "USTestVisit", "Name"
End If
If List.ListItems.Count <= 0 Then
USStyle3.Enabled = True
USStyle1.Enabled = False
USStyle6.Enabled = False
Else
USStyle3.Enabled = False
USStyle1.Enabled = True
USStyle6.Enabled = True
End If
Picture2.Visible = False
End Sub

Private Sub Text30_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text35.Visible = False
Text5.SetFocus
End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text8.SetFocus
End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text8.SetFocus
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text36.SetFocus
End If
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text10.SetFocus
End If
End Sub
Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text11.SetFocus
End If
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text12.SetFocus
End If
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text13.SetFocus
End If
End Sub
Private Sub Text13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text14.SetFocus
End If
End Sub
Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text15.SetFocus
End If
End Sub
Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text16.SetFocus
End If
End Sub
Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text17.SetFocus
End If
End Sub
Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text37.Visible = False
    Text38.Visible = False
    Text39.Visible = False
    Text40.Visible = False
Text18.SetFocus
End If
End Sub
Private Sub Text18_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text37.Visible = False
    Text38.Visible = False
    Text39.Visible = False
    Text40.Visible = False
Combo2.SetFocus
End If
End Sub
Private Sub Text19_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text23.SetFocus
End If
End Sub

Private Sub Text20_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text21.SetFocus
End If
End Sub

Private Sub Text21_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text22.SetFocus
End If
End Sub
Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text23.SetFocus
End If
End Sub
Private Sub Text23_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text24.SetFocus
End If
End Sub
Private Sub Text24_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text25.SetFocus
End If
End Sub

Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text26.SetFocus
End If
End Sub
Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text27.SetFocus
End If
End Sub
Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text28.SetFocus
End If
End Sub
Private Sub Text28_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text44.SetFocus
End If
End Sub
Private Sub Text44_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text43.SetFocus
End If
End Sub
Private Sub Text43_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Text1.Text = "" Then
Text43.Text = Format(Text43.Text, "0.00")
MsgBox "Please fillup the information of patient first then start entering Medicines record." & Chr(13) & "Thanks"
  Text28 = ""
  Text43 = ""
  Text44 = ""
Text1.SetFocus
Exit Sub
Else
Set Rs2 = DB.OpenRecordset("select * from MEDICINE")  'where REGNO='" & Text1.Text & "'")

       Rs2.AddNew
            Rs2!regno = Text1.Text
            Rs2!MEDICINE = Text28.Text
            Rs2!AMOUNT = Text43.Text
            Rs2!QTY = Text44.Text
            Rs2!Date = Format(Date, "DD")
            Rs2!Month = Format(Date, "MM")
            Rs2!Year = Format(Date, "YY")
       Rs2.Update
List2.ListItems.Clear
Set Rs2 = DB.OpenRecordset("select * from MEDICINE where REGNO='" & Text1.Text & "' and Date='" & Format(Date, "dd") & "' and month='" & Format(Date, "mm") & "' and year='" & Format(Date, "yy") & "'")
      Rs2.MoveFirst
      Do While Rs2.EOF = False
         Set Li = List2.ListItems.Add(, , Rs2!regno)
            Li.SubItems(1) = Rs2!MEDICINE
            Li.SubItems(2) = Rs2!QTY
            Li.SubItems(3) = Rs2!AMOUNT
            MakeTot = MakeTot + Val(Rs2!AMOUNT)
            Rs2.MoveNext
            Loop
            Text41.Text = MakeTot
  Text28 = ""
  Text43 = ""
  Text44 = ""
  
Text28.SetFocus
End If
End If

End Sub
Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text30.SetFocus
End If
End Sub
Private Sub Text30_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text31.SetFocus
End If
End Sub
Private Sub Text31_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text32.SetFocus
End If
End Sub
Private Sub Text32_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
MaskEdBox1.SetFocus
End If
End Sub
Private Sub Text33_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
MaskEdBox1.SetFocus
End If
End Sub

'Private Sub Text34_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'MaskEdBox1.SetFocus
'End If
'End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Picture2.Visible = True
Wait (5)
If Not Text1.Text = "" Then
Text5.Text = UCase(Text5.Text)
Text6.SetFocus
Exit Sub
End If
 On Error Resume Next
        FindsSex Text5, List, DB, "USTestVisit", "sex"
End If
If List.ListItems.Count <= 0 Then
USStyle3.Enabled = True
USStyle1.Enabled = False
USStyle6.Enabled = False
Else
USStyle3.Enabled = False
USStyle1.Enabled = True
USStyle6.Enabled = True
'Exit Sub
End If
Picture2.Visible = False
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
Dim TrackKey As String
    TrackKey = Chr(KeyAscii)
    If (Not IsNumeric(TrackKey) And Not (KeyAscii = vbKeyBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Text30.Text = Val(Text9) + Val(Text10) + Val(Text11) + Val(Text12) + Val(Text13) + Val(Text14) + Val(Text15) + Val(Text16) + Val(Text17) + Val(Text18) + Val(Text19) + Val(Text20) + Val(Text21) + Val(Text22) + Val(Text23) + Val(Text24) + Val(Text25) + Val(Text26) + Val(Text41)
Text32.Text = Val(Text30) - Val(Text31)

End Sub

Private Sub Timer2_Timer()
On Error Resume Next

Me.Width = 12015
Me.Height = 8835
End Sub

Private Sub USStyle1_Click()
 EditRecord
End Sub

Private Sub USStyle10_Click()
Set Rs2 = DB.OpenRecordset("select * from MEDICINE where REGNO='" & Text1.Text & "' and Date='" & USDate & "' and month='" & USMonth & "' and year='" & USYear & "' and MEDICINE='" & USMedicine & "'")
      If Rs2.EOF = False Then
       Rs2.Edit
         Rs2!MEDICINE = Text28.Text
         Rs2!QTY = Text44.Text
         Rs2!AMOUNT = Text43.Text
       Rs2.Update
         MsgBox "Record Edited"
      Text28.Text = ""
      Text43.Text = ""
      Text44.Text = ""
      findmedicines
      Else
      MsgBox "Record not found"
      End If
      
End Sub

Private Sub USStyle11_Click()
Set Rs2 = DB.OpenRecordset("select * from MEDICINE where REGNO='" & Text1.Text & "' and Date='" & USDate & "' and month='" & USMonth & "' and year='" & USYear & "' and MEDICINE='" & USMedicine & "'")
      If Rs2.EOF = False Then
       Rs2.Delete
        MsgBox "Record Deleted"
      Text28.Text = ""
      Text43.Text = ""
      Text44.Text = ""
      findmedicines
      Else
      MsgBox "Record not found"
      End If
End Sub

Private Sub USStyle12_Click()
      Text28.Text = ""
      Text43.Text = ""
      Text44.Text = ""
      Text28.SetFocus
End Sub

Private Sub USStyle2_Click()
Unload Me
End Sub

Private Sub USStyle3_Click()
SaveRecord
End Sub

Private Sub USStyle4_Click()
    With Adodc1.Recordset
        .Move (MSHFlexGrid1.row - 1) ' we minus one because row zero is the header row
        .Delete
        .Requery
    End With
    Adodc1.Refresh
    Set MSHFlexGrid1.DataSource = Adodc1
    MSHFlexGrid1.FormatString = "Old Reg | New Reg |Name |Age | Sex |Diagnosis | LastVisit | NextVisit"
    Call Form_Resize
End Sub

Private Sub EditRecord()
If Text1.Text = "" Then
MsgBox "Please Enter Registration No."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "please enter the Others"
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "Please Enter the Name"
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "please enter the Age"
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "please enter the Sex"
Text5.SetFocus
Exit Sub
End If
If Text6.Text = "" Then
Text6.Text = "..."
End If
If Text7.Text = "" Then
Text7.Text = "..."
End If
If Text8.Text = "" Then
MsgBox "Please Enter Address"
Text8.SetFocus
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "Please Enter Reg. Amount"
Text9.SetFocus
Exit Sub
End If
If Text10.Text = "" Then
Text10.Text = "..."
End If
If Text11.Text = "" Then
Text11.Text = "..."
End If
If Text12.Text = "" Then
Text12.Text = "..."
End If
If Text13.Text = "" Then
Text13.Text = "..."
End If
If Text14.Text = "" Then
Text14.Text = "..."
End If
If Text15.Text = "" Then
Text15.Text = "..."
End If
If Text16.Text = "" Then
Text16.Text = "..."
End If
If Text17.Text = "" Then
Text17.Text = "..."
End If
If Text18.Text = "" Then
Text18.Text = "..."
End If
If Text19.Text = "" Then
Text19.Text = "..."
End If
If Text20.Text = "" Then
Text20.Text = "..."
End If
If Text21.Text = "" Then
Text21.Text = "..."
End If
If Combo2.Text = "" Then
Combo2.Text = "..."
Text20.Text = "..."
End If
If Combo3.Text = "" Then
Combo3.Text = "..."
Text22.Text = "..."
End If
If Text23.Text = "" Then
Text23.Text = "..."
End If
If Text24.Text = "" Then
Text24.Text = "..."
End If
If Text25.Text = "" Then
Text25.Text = "0"
End If
If Text26.Text = "" Then
Text26.Text = "0"
End If
If Text27.Text = "" Then
Text27.Text = "..."
End If
If Text28.Text = "" Then
Text28.Text = "..."
End If
If Text29.Text = "" Then
Text29.Text = "..."
End If
If Text30.Text = "" Then
Text30.Text = "..."
End If
If Text31.Text = "" Then
Text31.Text = "0"
End If
If Text32.Text = "" Then
Text32.Text = "0"
End If
If Text33.Text = "" Then
Text33.Text = "1"
End If
If Text36.Text = "" Then
Text36.Text = "0"
End If
'Set Rs = DB.OpenRecordset("select * from USTestVisit where OldID='"

Rs.Edit
   Rs!OldID = Text1.Text
   'Rs!NewID = Text2.Text
   Rs!Name = Text3.Text
   Rs!TmpAddress = Text8.Text
   Rs!perAddress = Text36.Text
   Rs!Age = Text4.Text
   Rs!Sex = Text5.Text
   Rs!Tel = Text6.Text
   Rs!Email = Text7.Text
   Rs!Treatment = Text27.Text
   Rs!LastVisit = Format(Date, "d/mmm/yy")
   Rs!NextVisit = MaskEdBox1.Text
   
   Rs!Type = "Rew"
   If Combo1.Text = "" Then
   Combo1.Clear
   Combo1.Text = Text40.Text
   End If
   Rs!DOCTOR = Combo1.Text
   Rs!Reg = Text9.Text
   Rs!Con = Text10.Text
   Rs!AR = Text11.Text
   Rs!TN = Text12.Text
   Rs!fun = Text13.Text
   Rs!sl = Text14.Text
   Rs!OTHERS = Text15.Text
   Rs!IDO = Text16.Text
   Rs!GONIOS = Text17.Text
   Rs!AscanKerometry = Text18.Text
   If Combo2.Text = "" Then
   Combo2.Clear
   Combo2.Text = Text39.Text
   End If
   Rs!PeRITYPE = Combo2.Text
   Rs!PeRIAMOUNT = Text20.Text
   If Combo3.Text = "" Then
   Combo3.Clear
   Combo3.Text = Text38.Text
   End If
   Rs!MINOROTType = Combo3.Text
   Rs!minorOTAMOUNT = Text22.Text
   If Combo4.Text = "" Then
   Combo4.Clear
   Combo4.Text = Text37.Text
   Rs!MAJOROTType = Combo4.Text
   Rs!MAJOROTAMOUNT = Text19.Text
   Rs!SHIRMER = Text23.Text
   Rs!RBS = Text24.Text
   Rs!FFA = Text25.Text
   Rs!BTCT = Text26.Text
   Rs!DIAGNOSIS = Text27.Text
   Rs!MEDICINES = Text28.Text
   Rs!REMARK = Text29.Text
   Rs!PAID = Text31.Text
   Rs!BALANCE = Text32.Text
   Rs!TOTAL = Text30.Text
   Rs!YearF = Format(Date, "yy")
   Rs!DateF = Format(Date, "dd")
   Rs!MonthF = Format(Date, "mm")
   Rs.Update
   'Form_Load
   Text1.SetFocus
   MsgBox "Record has been edited"
   'BLANK
   'List.Visible = False
   FindsOldID Text1, List, DB, "USTestVisit", "OldID"
   End If
   BLANK
End Sub

Private Sub SaveRecord()
If Text1.Text = "" Then
MsgBox "Please Enter Registration No."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "please enter the Others"
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "Please Enter the Name"
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "please enter the Age"
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "please enter the Sex"
Text5.SetFocus
Exit Sub
End If
If Text6.Text = "" Then
Text6.Text = "..."
End If
If Text7.Text = "" Then
Text7.Text = "..."
End If
If Text8.Text = "" Then
MsgBox "Please Enter Address"
Text8.SetFocus
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "Please Enter Reg. Amount"
Text9.SetFocus
Exit Sub
End If
If Text10.Text = "" Then
Text10.Text = "..."
End If
If Text11.Text = "" Then
Text11.Text = "..."
End If
If Text12.Text = "" Then
Text12.Text = "..."
End If
If Text13.Text = "" Then
Text13.Text = "..."
End If
If Text14.Text = "" Then
Text14.Text = "..."
End If
If Text15.Text = "" Then
Text15.Text = "..."
End If
If Text16.Text = "" Then
Text16.Text = "..."
End If
If Text17.Text = "" Then
Text17.Text = "..."
End If
If Text18.Text = "" Then
Text18.Text = "..."
End If
If Text19.Text = "" Then
Text19.Text = "..."
End If
If Text20.Text = "" Then
Text20.Text = "..."
End If
If Text21.Text = "" Then
Text21.Text = "..."
End If
If Combo2.Text = "" Then
Combo2.Text = "..."
Text20.Text = "..."
End If
If Combo3.Text = "" Then
Combo3.Text = "..."
Text22.Text = "..."
End If
If Text23.Text = "" Then
Text23.Text = "..."
End If
If Text24.Text = "" Then
Text24.Text = "..."
End If
If Text25.Text = "" Then
Text25.Text = "0"
End If
If Text26.Text = "" Then
Text26.Text = "0"
End If
If Text27.Text = "" Then
Text27.Text = "..."
End If
If Text28.Text = "" Then
Text28.Text = "..."
End If
If Text29.Text = "" Then
Text29.Text = "..."
End If
If Text30.Text = "" Then
Text30.Text = "..."
End If
If Text31.Text = "" Then
Text31.Text = "0"
End If
If Text32.Text = "" Then
Text32.Text = "0"
End If
If Text33.Text = "" Then
Text33.Text = "1"
End If
If Text36.Text = "" Then
Text36.Text = "0"
End If
If Combo1.Text = "" Then
MsgBox "Please Select the Doctor Name"
Combo1.SetFocus
End If
If Combo2.Text = "" Then
Combo2.Text = "NIL"
Combo2_Click
End If
If Combo3.Text = "" Then
Combo3.Text = "NIL"
Combo3_Click
End If
If Combo4.Text = "" Then
Combo4.Text = "NIL"
Combo4_Click
End If

Rs.AddNew
   Rs!OldID = Text1.Text
   'Rs!NewID = Text2.Text
   Rs!Name = Text3.Text
   Rs!TmpAddress = Text8.Text
   Rs!perAddress = Text36.Text
   Rs!Age = Text4.Text
   Rs!Sex = Text5.Text
   Rs!Tel = Text6.Text
   'Rs!Email = Text7.Text
   Rs!Treatment = Text27.Text
   Rs!LastVisit = Format(Date, "d/mmm/yy")
   Rs!NextVisit = MaskEdBox1.Text
   'Rs!Narration = Text34.Text
   Rs!Type = "New"
   Rs!DOCTOR = Combo1.Text
   Rs!Reg = Text9.Text
   Rs!Con = Text10.Text
   Rs!AR = Text11.Text
   Rs!TN = Text12.Text
   Rs!fun = Text13.Text
   Rs!sl = Text14.Text
   Rs!OTHERS = Text15.Text
   Rs!IDO = Text16.Text
   Rs!GONIOS = Text17.Text
   Rs!AscanKerometry = Text18.Text
   Rs!PeRITYPE = Combo2.Text
   Rs!PeRIAMOUNT = Text20.Text
   Rs!MINOROTType = Combo3.Text
   Rs!minorOTAMOUNT = Text22.Text
   Rs!MAJOROTType = Combo4.Text
   Rs!MAJOROTAMOUNT = Text19.Text
   Rs!SHIRMER = Text23.Text
   Rs!RBS = Text24.Text
   Rs!FFA = Text25.Text
   Rs!BTCT = Text26.Text
   Rs!DIAGNOSIS = Text27.Text
   Rs!MEDICINES = Text28.Text
   Rs!REMARK = Text29.Text
   Rs!PAID = Text31.Text
   Rs!BALANCE = Text32.Text
   Rs!TOTAL = Text30.Text
   Rs!YearF = Format(Date, "yy")
   Rs!DateF = Format(Date, "dd")
   Rs!MonthF = Format(Date, "mm")
   Rs.Update
   MsgBox "Record has been saved"
   
   'List.Visible = False
   'Form_Load
   FindsOldID Text1, List, DB, "USTestVisit", "OldID"
   Text1.SetFocus
   BLANK
End Sub

Private Sub ReSaveRecord()

If Text1.Text = "" Then
MsgBox "Please Enter Registration No."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "please enter the Others"
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
MsgBox "Please Enter the Name"
Text3.SetFocus
Exit Sub
End If
If Text4.Text = "" Then
MsgBox "please enter the Age"
Text4.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
MsgBox "please enter the Sex"
Text5.SetFocus
Exit Sub
End If
If Text6.Text = "" Then
Text6.Text = "..."
End If
If Text7.Text = "" Then
Text7.Text = "..."
End If
If Text8.Text = "" Then
MsgBox "Please Enter Address"
Text8.SetFocus
Exit Sub
End If
If Text9.Text = "" Then
MsgBox "Please Enter Reg. Amount"
Text9.SetFocus
Exit Sub
End If
If Text10.Text = "" Then
Text10.Text = "..."
End If
If Text11.Text = "" Then
Text11.Text = "..."
End If
If Text12.Text = "" Then
Text12.Text = "..."
End If
If Text13.Text = "" Then
Text13.Text = "..."
End If
If Text14.Text = "" Then
Text14.Text = "..."
End If
If Text15.Text = "" Then
Text15.Text = "..."
End If
If Text16.Text = "" Then
Text16.Text = "..."
End If
If Text17.Text = "" Then
Text17.Text = "..."
End If
If Text18.Text = "" Then
Text18.Text = "..."
End If
If Text19.Text = "" Then
Text19.Text = "..."
End If
If Text20.Text = "" Then
Text20.Text = "..."
End If
If Text21.Text = "" Then
Text21.Text = "..."
End If
If Combo2.Text = "" Then
Combo2.Text = "..."
Text20.Text = "..."
End If
If Combo3.Text = "" Then
Combo3.Text = "..."
Text22.Text = "..."
End If
If Text23.Text = "" Then
Text23.Text = "..."
End If
If Text24.Text = "" Then
Text24.Text = "..."
End If
If Text25.Text = "" Then
Text25.Text = "0"
End If
If Text26.Text = "" Then
Text26.Text = "0"
End If
If Text27.Text = "" Then
Text27.Text = "..."
End If
If Text28.Text = "" Then
Text28.Text = "..."
End If
If Text29.Text = "" Then
Text29.Text = "..."
End If
If Text30.Text = "" Then
Text30.Text = "..."
End If
If Text31.Text = "" Then
Text31.Text = "0"
End If
If Text32.Text = "" Then
Text32.Text = "0"
End If
If Text33.Text = "" Then
Text33.Text = "1"
End If
If Text36.Text = "" Then
Text36.Text = "0"
End If


Rs.AddNew
   Rs!OldID = Text1.Text
   'Rs!NewID = Text2.Text
   Rs!Name = Text3.Text
   Rs!TmpAddress = Text8.Text
   Rs!perAddress = Text36.Text
   Rs!Age = Text4.Text
   Rs!Sex = Text5.Text
   Rs!Tel = Text6.Text
   Rs!Email = Text7.Text
   Rs!Treatment = Text27.Text
   Rs!LastVisit = Format(Date, "d/mmm/yy")
   Rs!NextVisit = MaskEdBox1.Text
   
   Rs!Type = "Rew"
   If Combo1.Text = "" Then
   Combo1.Clear
   Combo1.Text = Text40.Text
   End If
   Rs!DOCTOR = Combo1.Text
   Rs!Reg = Text9.Text
   Rs!Con = Text10.Text
   Rs!AR = Text11.Text
   Rs!TN = Text12.Text
   Rs!fun = Text13.Text
   Rs!sl = Text14.Text
   Rs!OTHERS = Text15.Text
   Rs!IDO = Text16.Text
   Rs!GONIOS = Text17.Text
   Rs!AscanKerometry = Text18.Text
   If Combo2.Text = "" Then
   Combo2.Clear
   Combo2.Text = Text39.Text
   End If
   Rs!PeRITYPE = Combo2.Text
   Rs!PeRIAMOUNT = Text20.Text
   If Combo3.Text = "" Then
   Combo3.Clear
   Combo3.Text = Text38.Text
   End If
   Rs!MINOROTType = Combo3.Text
   Rs!minorOTAMOUNT = Text22.Text
   If Combo4.Text = "" Then
   Combo4.Clear
   Combo4.Text = Text37.Text
   Rs!MAJOROTType = Combo4.Text
   Rs!MAJOROTAMOUNT = Text19.Text
   Rs!SHIRMER = Text23.Text
   Rs!RBS = Text24.Text
   Rs!FFA = Text25.Text
   Rs!BTCT = Text26.Text
   Rs!DIAGNOSIS = Text27.Text
   Rs!MEDICINES = Text28.Text
   Rs!REMARK = Text29.Text
   Rs!PAID = Text31.Text
   Rs!BALANCE = Text32.Text
   Rs!TOTAL = Text30.Text
   Rs!YearF = Format(Date, "yy")
   Rs!DateF = Format(Date, "dd")
   Rs!MonthF = Format(Date, "mm")
   Rs.Update
   'Form_Load
   Text1.SetFocus
   MsgBox "ReRegister Record has been saved"
  
   'List.Visible = False
   FindsOldID Text1, List, DB, "USTestVisit", "OldID"
   Text1.SetFocus
   End If
    BLANK
End Sub

Public Function FindsOldID(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
'Dim sTemp As Recordset
Label5.Caption = "0"
            'Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo3.Text = ""
            Combo4.Text = ""
            
 FindPat = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
            sTemp.MoveFirst
            sList.ListItems.Clear
            Do While Not sTemp.EOF
            Set Li = List.ListItems.Add(, , sTemp!OldID)
            Li.SubItems(1) = sTemp!OldID
            Li.SubItems(2) = sTemp!Type
            Li.SubItems(3) = sTemp!Name
            Li.SubItems(4) = sTemp!Age
            Li.SubItems(5) = sTemp!Sex
            Li.SubItems(6) = sTemp!TmpAddress
            Li.SubItems(7) = sTemp!perAddress
            Li.SubItems(8) = sTemp!Tel
            Li.SubItems(9) = sTemp!Reg
            Li.SubItems(10) = sTemp!Con
            Li.SubItems(11) = sTemp!AR
            Li.SubItems(12) = sTemp!TN
            Li.SubItems(13) = sTemp!fun
            Li.SubItems(14) = sTemp!sl
            Li.SubItems(15) = sTemp!OTHERS
            Li.SubItems(16) = sTemp!IDO
            Li.SubItems(17) = sTemp!GONIOS
            Li.SubItems(18) = sTemp!AscanKerometry
            Li.SubItems(19) = sTemp!PeRITYPE
            Li.SubItems(20) = sTemp!PeRIAMOUNT
            Li.SubItems(21) = sTemp!MINOROTType
            Li.SubItems(22) = sTemp!minorOTAMOUNT
            Li.SubItems(23) = sTemp!MAJOROTType
            Li.SubItems(24) = sTemp!MAJOROTAMOUNT
            Li.SubItems(25) = sTemp!SHIRMER
            Li.SubItems(26) = sTemp!RBS
            Li.SubItems(27) = sTemp!FFA
            Li.SubItems(28) = sTemp!BTCT
            Li.SubItems(29) = sTemp!DIAGNOSIS
            Li.SubItems(30) = sTemp!MEDICINES
            Li.SubItems(31) = sTemp!REMARK
            Li.SubItems(32) = sTemp!TOTAL
            Li.SubItems(33) = sTemp!PAID
            Li.SubItems(34) = sTemp!BALANCE
            Li.SubItems(35) = sTemp!DOCTOR
            Li.SubItems(36) = sTemp!DateF
            Li.SubItems(37) = sTemp!MonthF
            Li.SubItems(38) = sTemp!YearF
        
            
            sTemp.MoveNext
            Loop
            
        End If
        Label5.Caption = List.ListItems.Count
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                 FindPat = True
    Else
        sList.ListItems.Clear
    End If
End If
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Picture1.Visible = False
End Function


Public Function FindsNewID(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
'Dim sTemp As Recordset
Label5.Caption = "0"
            Text1.Text = ""
           ' Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            'Text33.Text = ""
            'Text34.Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo3.Text = ""
            Combo4.Text = ""
 FindPat = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
             sTemp.MoveFirst
            sList.ListItems.Clear
            Do While Not sTemp.EOF
            Set Li = List.ListItems.Add(, , sTemp!OldID)
            Li.SubItems(1) = sTemp!OldID
            Li.SubItems(2) = sTemp!Type
            Li.SubItems(3) = sTemp!Name
            Li.SubItems(4) = sTemp!Age
            Li.SubItems(5) = sTemp!Sex
            Li.SubItems(6) = sTemp!TmpAddress
            Li.SubItems(7) = sTemp!perAddress
            Li.SubItems(8) = sTemp!Tel
            Li.SubItems(9) = sTemp!Reg
            Li.SubItems(10) = sTemp!Con
            Li.SubItems(11) = sTemp!AR
            Li.SubItems(12) = sTemp!TN
            Li.SubItems(13) = sTemp!fun
            Li.SubItems(14) = sTemp!sl
            Li.SubItems(15) = sTemp!OTHERS
            Li.SubItems(16) = sTemp!IDO
            Li.SubItems(17) = sTemp!GONIOS
            Li.SubItems(18) = sTemp!AscanKerometry
            Li.SubItems(19) = sTemp!PeRITYPE
            Li.SubItems(20) = sTemp!PeRIAMOUNT
            Li.SubItems(21) = sTemp!MINOROTType
            Li.SubItems(22) = sTemp!minorOTAMOUNT
            Li.SubItems(23) = sTemp!MAJOROTType
            Li.SubItems(24) = sTemp!MAJOROTAMOUNT
            Li.SubItems(25) = sTemp!SHIRMER
            Li.SubItems(26) = sTemp!RBS
            Li.SubItems(27) = sTemp!FFA
            Li.SubItems(28) = sTemp!BTCT
            Li.SubItems(29) = sTemp!DIAGNOSIS
            Li.SubItems(30) = sTemp!MEDICINES
            Li.SubItems(31) = sTemp!REMARK
            Li.SubItems(32) = sTemp!TOTAL
            Li.SubItems(33) = sTemp!PAID
            Li.SubItems(34) = sTemp!BALANCE
            Li.SubItems(35) = sTemp!DOCTOR
            Li.SubItems(36) = sTemp!DateF
            Li.SubItems(37) = sTemp!MonthF
            Li.SubItems(38) = sTemp!YearF
            sTemp.MoveNext
            Loop
           
        End If
        Label5.Caption = List.ListItems.Count
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                 FindPat = True
    Else
        sList.ListItems.Clear
    End If
End If
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Picture1.Visible = False
End Function

Public Function FindsName(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
'Dim sTemp As Recordset
Label5.Caption = "0"
            Text1.Text = ""
            Text2.Text = ""
            Text36.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            'Text33.Text = ""
            'Text34.Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo3.Text = ""
            Combo4.Text = ""
 FindPat = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
            sTemp.MoveFirst
            sList.ListItems.Clear
            Do While Not sTemp.EOF
            Set Li = List.ListItems.Add(, , sTemp!OldID)
            Li.SubItems(1) = sTemp!OldID
            Li.SubItems(2) = sTemp!Type
            Li.SubItems(3) = sTemp!Name
            Li.SubItems(4) = sTemp!Age
            Li.SubItems(5) = sTemp!Sex
            Li.SubItems(6) = sTemp!TmpAddress
            Li.SubItems(7) = sTemp!perAddress
            Li.SubItems(8) = sTemp!Tel
            Li.SubItems(9) = sTemp!Reg
            Li.SubItems(10) = sTemp!Con
            Li.SubItems(11) = sTemp!AR
            Li.SubItems(12) = sTemp!TN
            Li.SubItems(13) = sTemp!fun
            Li.SubItems(14) = sTemp!sl
            Li.SubItems(15) = sTemp!OTHERS
            Li.SubItems(16) = sTemp!IDO
            Li.SubItems(17) = sTemp!GONIOS
            Li.SubItems(18) = sTemp!AscanKerometry
            Li.SubItems(19) = sTemp!PeRITYPE
            Li.SubItems(20) = sTemp!PeRIAMOUNT
            Li.SubItems(21) = sTemp!MINOROTType
            Li.SubItems(22) = sTemp!minorOTAMOUNT
            Li.SubItems(23) = sTemp!MAJOROTType
            Li.SubItems(24) = sTemp!MAJOROTAMOUNT
            Li.SubItems(25) = sTemp!SHIRMER
            Li.SubItems(26) = sTemp!RBS
            Li.SubItems(27) = sTemp!FFA
            Li.SubItems(28) = sTemp!BTCT
            Li.SubItems(29) = sTemp!DIAGNOSIS
            Li.SubItems(30) = sTemp!MEDICINES
            Li.SubItems(31) = sTemp!REMARK
            Li.SubItems(32) = sTemp!TOTAL
            Li.SubItems(33) = sTemp!PAID
            Li.SubItems(34) = sTemp!BALANCE
            Li.SubItems(35) = sTemp!DOCTOR
            Li.SubItems(36) = sTemp!DateF
            Li.SubItems(37) = sTemp!MonthF
            Li.SubItems(38) = sTemp!YearF
            sTemp.MoveNext
            Loop
            
        End If
        Label5.Caption = List.ListItems.Count
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                 FindPat = True
    Else
        sList.ListItems.Clear
    End If
End If
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
Picture1.Visible = False
End Function

Public Function FindsSex(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
'Dim sTemp As Recordset
Label5.Caption = "0"
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            'Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            'Text33.Text = ""
            'Text34.Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo3.Text = ""
            Combo4.Text = ""
 FindPat = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
             sTemp.MoveFirst
            sList.ListItems.Clear
            Do While Not sTemp.EOF
            Set Li = List.ListItems.Add(, , sTemp!OldID)
            Li.SubItems(1) = sTemp!OldID
            Li.SubItems(2) = sTemp!Type
            Li.SubItems(3) = sTemp!Name
            Li.SubItems(4) = sTemp!Age
            Li.SubItems(5) = sTemp!Sex
            Li.SubItems(6) = sTemp!TmpAddress
            Li.SubItems(7) = sTemp!perAddress
            Li.SubItems(8) = sTemp!Tel
            Li.SubItems(9) = sTemp!Reg
            Li.SubItems(10) = sTemp!Con
            Li.SubItems(11) = sTemp!AR
            Li.SubItems(12) = sTemp!TN
            Li.SubItems(13) = sTemp!fun
            Li.SubItems(14) = sTemp!sl
            Li.SubItems(15) = sTemp!OTHERS
            Li.SubItems(16) = sTemp!IDO
            Li.SubItems(17) = sTemp!GONIOS
            Li.SubItems(18) = sTemp!AscanKerometry
            Li.SubItems(19) = sTemp!PeRITYPE
            Li.SubItems(20) = sTemp!PeRIAMOUNT
            Li.SubItems(21) = sTemp!MINOROTType
            Li.SubItems(22) = sTemp!minorOTAMOUNT
            Li.SubItems(23) = sTemp!MAJOROTType
            Li.SubItems(24) = sTemp!MAJOROTAMOUNT
            Li.SubItems(25) = sTemp!SHIRMER
            Li.SubItems(26) = sTemp!RBS
            Li.SubItems(27) = sTemp!FFA
            Li.SubItems(28) = sTemp!BTCT
            Li.SubItems(29) = sTemp!DIAGNOSIS
            Li.SubItems(30) = sTemp!MEDICINES
            Li.SubItems(31) = sTemp!REMARK
            Li.SubItems(32) = sTemp!TOTAL
            Li.SubItems(33) = sTemp!PAID
            Li.SubItems(34) = sTemp!BALANCE
            Li.SubItems(35) = sTemp!DOCTOR
            Li.SubItems(36) = sTemp!DateF
            Li.SubItems(37) = sTemp!MonthF
            Li.SubItems(38) = sTemp!YearF
            
            sTemp.MoveNext
            Loop
            
            Text35.Visible = True
        End If
        Label5.Caption = List.ListItems.Count
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                 FindPat = True
    Else
        sList.ListItems.Clear
    End If
End If
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
Picture1.Visible = False
End Function

Private Sub BLANK()
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            Combo1.Text = ""
            Combo2.Text = ""
            Combo3.Text = ""
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Text36.Text = ""
            Text37.Text = ""
            Text38.Text = ""
            Text39.Text = ""
            Text40.Text = ""
End Sub

Private Sub BLANK2()
            Text1.Text = ""
            'Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Text36.Text = ""
            Text37.Text = ""
            Text38.Text = ""
            Text39.Text = ""
            Text40.Text = ""
End Sub
Private Sub BLANK3()
            Text1.Text = ""
            Text2.Text = ""
            Text36.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Text37.Text = ""
            Text38.Text = ""
            Text39.Text = ""
            Text40.Text = ""
End Sub
Private Sub BLANK4()
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            'Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Text37.Text = ""
            Text38.Text = ""
            Text39.Text = ""
            Text40.Text = ""
End Sub
Public Sub SotrListView(ByVal lstView As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional ByVal TypeSort As String = "NUMBER")

    On Error Resume Next
  
    With lstView
    
        ' Display the hourglass cursor whilst sorting
        
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        LockWindowUpdate .hWnd
        
       
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.Index - 1
    
        Select Case UCase$(TypeSort)
        Case "DATE"
        
            ' Sort by date.
            
            strFormat = "YYYYMMDDHhNnSs"
        
            ' Loop through the values in this column. Re-format
            ' the dates so as they can be sorted alphabetically,
            ' having already stored their visible values in the
            ' tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
        Case "NUMBER"
        
            ' Sort Numerically
        
            strFormat = String(30, "0") & "." & String(30, "0")
        
            ' Loop through the values in this column. Re-format the values so as they
            ' can be sorted alphabetically, having already stored their visible
            ' values in the tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(Val(.Text)) Then
                                If CDbl(Val(.Text)) >= 0 Then
                                    .Text = Format(CDbl(Val(.Text)), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(Val(.Text)), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else   ' Assume sort by string
            
            ' Sort alphabetically. This is the only sort provided
            ' by the MS ListView control (at this time), and as
            ' such we don't really need to do much here
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
    
        ' Unlock the list window so that the OCX can update it
        
        LockWindowUpdate 0&
        
        ' Restore the previous cursor
        
        .MousePointer = lngCursor
    
    End With
    
    Set lstView = Nothing
    Set ColumnHeader = Nothing
End Sub


'****************************************************************
' InvNumber
' Function used to enable negative numbers to be sorted
' alphabetically by switching the characters
'----------------------------------------------------------------

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

Private Sub USStyle5_Click()
Picture1.Visible = True
List.ListItems.Clear
BLANK
USStyle3.Enabled = True
Text1.SetFocus
Picture2.Visible = False
End Sub

Private Sub USStyle6_Click()
ReSaveRecord
End Sub

Private Sub USStyle7_Click()
On Error Resume Next
 SQL = "select * from USTestVisit where name='" & Text3.Text & "' AND OLDID='" & Text1.Text & "'"
Set MyRs = MyDb.OpenRecordset(SQL)
          
          MyRs.MoveNext
    

            Text1.Text = MyRs!OldID
            Text2.Text = MyRs!Type
            Text3.Text = MyRs!Name
            Text4.Text = MyRs!Age
            Text5.Text = MyRs!Sex
            Text6.Text = MyRs!Tel
            Text7.Text = MyRs!Email
            Text8.Text = MyRs!TmpAddress
            Text36.Text = MyRs!perAddress
            Text9.Text = MyRs!Reg
            Text10.Text = MyRs!Con
            Text11.Text = MyRs!AR
            Text12.Text = MyRs!TN
            Text13.Text = MyRs!fun
            Text14.Text = MyRs!sl
            Text15.Text = MyRs!OTHERS
            Text16.Text = MyRs!IDO
            
            Text17.Text = MyRs!GONIOS
            Text18.Text = MyRs!AscanKerometry
            Text19.Text = MyRs!MAJOROTAMOUNT
            Combo2.Text = MyRs!PeRITYPE
            Text20.Text = MyRs!PeRIAMOUNT
            Combo3.Text = MyRs!MINOROTType
            Text22.Text = MyRs!minorOTAMOUNT
            Combo4.Text = MyRs!MAJOROTType
            Text19.Text = MyRs!MAJOROTAMOUNT
            
            Text23.Text = MyRs!Email
            Text24.Text = MyRs!Address
            Text25.Text = MyRs!Reg
            Text26.Text = MyRs!Con
            Text27.Text = MyRs!AR
            Text28.Text = MyRs!TN
            Text29.Text = MyRs!REMARK
            Text30.Text = MyRs!sl
            Text31.Text = MyRs!OTHERS
            Text32.Text = MyRs!BALANCE
            Text35.Text = MyRs!NextVisit
            'Text35.Text = MyRs!nextvisit
            Combo1.Text = MyRs!DOCTOR
End Sub

Private Sub USStyle8_Click()
On Error Resume Next
SQL = "select * from USTestVisit where name='" & Text3.Text & "' AND OLDID='" & Text1.Text & "'"
Set MyRs = MyDb.OpenRecordset(SQL)
          
          MyRs.MovePrevious
             Text1.Text = MyRs!OldID
            Text2.Text = MyRs!Type
            Text3.Text = MyRs!Name
            Text4.Text = MyRs!Age
            Text5.Text = MyRs!Sex
            Text6.Text = MyRs!Tel
            Text7.Text = MyRs!Email
            Text8.Text = MyRs!TmpAddress
            Text36.Text = MyRs!perAddress
            Text9.Text = MyRs!Reg
            Text10.Text = MyRs!Con
            Text11.Text = MyRs!AR
            Text12.Text = MyRs!TN
            Text13.Text = MyRs!fun
            Text14.Text = MyRs!sl
            Text15.Text = MyRs!OTHERS
            Text16.Text = MyRs!IDO
            
            Text17.Text = MyRs!GONIOS
            Text18.Text = MyRs!AscanKerometry
            Text19.Text = MyRs!MAJOROTAMOUNT
            Combo2.Text = MyRs!PeRITYPE
            Text20.Text = MyRs!PeRIAMOUNT
            Combo3.Text = MyRs!MINOROTType
            Text22.Text = MyRs!minorOTAMOUNT
            Combo4.Text = MyRs!MAJOROTType
            Text19.Text = MyRs!MAJOROTAMOUNT
            
            Text23.Text = MyRs!Email
            Text24.Text = MyRs!Address
            Text25.Text = MyRs!Reg
            Text26.Text = MyRs!Con
            Text27.Text = MyRs!AR
            Text28.Text = MyRs!TN
            Text29.Text = MyRs!REMARK
            Text30.Text = MyRs!sl
            Text31.Text = MyRs!OTHERS
            Text32.Text = MyRs!BALANCE
            Text35.Text = MyRs!NextVisit
            'Text35.Text = MyRs!nextvisit
            Combo1.Text = MyRs!DOCTOR
End Sub

Private Sub USStyle9_Click()
On Error Resume Next
SQL = "select * from USTestVisit where name='" & Text3.Text & "' AND OLDID='" & Text1.Text & "'"
Set MyRs = MyDb.OpenRecordset(SQL)
          
          MyRs.MoveFirst
            Text1.Text = MyRs!OldID
            Text2.Text = MyRs!Type
            Text3.Text = MyRs!Name
            Text4.Text = MyRs!Age
            Text5.Text = MyRs!Sex
            Text6.Text = MyRs!Tel
            Text7.Text = MyRs!Email
            Text8.Text = MyRs!TmpAddress
            Text36.Text = MyRs!perAddress
            Text9.Text = MyRs!Reg
            Text10.Text = MyRs!Con
            Text11.Text = MyRs!AR
            Text12.Text = MyRs!TN
            Text13.Text = MyRs!fun
            Text14.Text = MyRs!sl
            Text15.Text = MyRs!OTHERS
            Text16.Text = MyRs!IDO
            
            Text17.Text = MyRs!GONIOS
            Text18.Text = MyRs!AscanKerometry
            Text19.Text = MyRs!MAJOROTAMOUNT
            Combo2.Text = MyRs!PeRITYPE
            Text20.Text = MyRs!PeRIAMOUNT
            Combo3.Text = MyRs!MINOROTType
            Text22.Text = MyRs!minorOTAMOUNT
            Combo4.Text = MyRs!MAJOROTType
            Text19.Text = MyRs!MAJOROTAMOUNT
            
            Text23.Text = MyRs!Email
            Text24.Text = MyRs!Address
            Text25.Text = MyRs!Reg
            Text26.Text = MyRs!Con
            Text27.Text = MyRs!AR
            Text28.Text = MyRs!TN
            Text29.Text = MyRs!REMARK
            Text30.Text = MyRs!sl
            Text31.Text = MyRs!OTHERS
            Text32.Text = MyRs!BALANCE
            Text35.Text = MyRs!NextVisit
            'Text35.Text = MyRs!nextvisit
            Combo1.Text = MyRs!DOCTOR
End Sub
