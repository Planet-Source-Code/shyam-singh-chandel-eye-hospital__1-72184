VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Object = "{070DF375-7EED-4C5F-BF9C-48A3AA97B6BD}#1.0#0"; "USFormControl.ocx"
Begin VB.Form FrmMain2 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "EYE Clinic Zone - US Softwares"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   Icon            =   "FrmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin USForm.UserControl1 UserControl11 
      Align           =   1  'Align Top
      Height          =   8490
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   14975
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   2760
         ScaleHeight     =   3225
         ScaleWidth      =   2865
         TabIndex        =   25
         Top             =   890
         Visible         =   0   'False
         Width           =   2895
         Begin Project1.USStyle USStyle22 
            Height          =   615
            Left            =   -120
            TabIndex        =   26
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "New Registration"
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
         Begin Project1.USStyle USStyle23 
            Height          =   615
            Left            =   -360
            TabIndex        =   27
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Test and Visits"
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
         Begin Project1.USStyle USStyle24 
            Height          =   615
            Left            =   -360
            TabIndex        =   28
            Top             =   1200
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Patient Pres."
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
         Begin Project1.USStyle USStyle25 
            Height          =   615
            Left            =   -360
            TabIndex        =   29
            Top             =   1800
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Reports"
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
         Begin Project1.USStyle USStyle26 
            Height          =   135
            Left            =   -360
            TabIndex        =   30
            Top             =   2400
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   238
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16711680
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
         Begin Project1.USStyle USStyle27 
            Height          =   735
            Left            =   -360
            TabIndex        =   31
            Top             =   2520
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   1440
         ScaleHeight     =   3225
         ScaleWidth      =   2865
         TabIndex        =   18
         Top             =   890
         Visible         =   0   'False
         Width           =   2895
         Begin Project1.USStyle USStyle16 
            Height          =   615
            Left            =   -120
            TabIndex        =   19
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "New Registration"
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
         Begin Project1.USStyle USStyle17 
            Height          =   615
            Left            =   -360
            TabIndex        =   20
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Test and Visits"
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
         Begin Project1.USStyle USStyle18 
            Height          =   615
            Left            =   -360
            TabIndex        =   21
            Top             =   1200
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Patient Pres."
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
         Begin Project1.USStyle USStyle19 
            Height          =   615
            Left            =   -360
            TabIndex        =   22
            Top             =   1800
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Reports"
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
         Begin Project1.USStyle USStyle20 
            Height          =   135
            Left            =   -360
            TabIndex        =   23
            Top             =   2400
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   238
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16711680
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
         Begin Project1.USStyle USStyle21 
            Height          =   735
            Left            =   -360
            TabIndex        =   24
            Top             =   2520
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3225
         ScaleWidth      =   2865
         TabIndex        =   11
         Top             =   890
         Visible         =   0   'False
         Width           =   2895
         Begin Project1.USStyle USStyle10 
            Height          =   615
            Left            =   -120
            TabIndex        =   12
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "New Registration"
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
         Begin Project1.USStyle USStyle11 
            Height          =   615
            Left            =   -360
            TabIndex        =   13
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Test and Visits"
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
         Begin Project1.USStyle USStyle12 
            Height          =   615
            Left            =   -360
            TabIndex        =   14
            Top             =   1200
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Patient Pres."
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
         Begin Project1.USStyle USStyle13 
            Height          =   615
            Left            =   -360
            TabIndex        =   15
            Top             =   1800
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Print Reports"
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
         Begin Project1.USStyle USStyle14 
            Height          =   135
            Left            =   -360
            TabIndex        =   16
            Top             =   2400
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   238
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16711680
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
         Begin Project1.USStyle USStyle15 
            Height          =   735
            Left            =   -360
            TabIndex        =   17
            Top             =   2520
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   7905
         TabIndex        =   1
         Top             =   480
         Width           =   7935
         Begin Project1.USStyle USStyle9 
            Height          =   375
            Left            =   10560
            TabIndex        =   2
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16777088
            Checked         =   0   'False
            ColorButtonHover=   160
            ColorButtonUp   =   128
            ColorButtonDown =   240
            BorderBrightness=   0
            ColorBright     =   255
            DisplayHand     =   0   'False
            ColorScheme     =   4
         End
         Begin Project1.USStyle USStyle8 
            Height          =   375
            Left            =   5160
            TabIndex        =   3
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "About"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
         Begin Project1.USStyle USStyle7 
            Height          =   375
            Left            =   6480
            TabIndex        =   4
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
         Begin Project1.USStyle USStyle6 
            Height          =   375
            Left            =   7920
            TabIndex        =   5
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16777088
            Checked         =   0   'False
            ColorButtonHover=   160
            ColorButtonUp   =   128
            ColorButtonDown =   240
            BorderBrightness=   0
            ColorBright     =   255
            DisplayHand     =   0   'False
            ColorScheme     =   4
         End
         Begin Project1.USStyle USStyle5 
            Height          =   375
            Left            =   9240
            TabIndex        =   6
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            ForeColor       =   16777088
            Checked         =   0   'False
            ColorButtonHover=   160
            ColorButtonUp   =   128
            ColorButtonDown =   240
            BorderBrightness=   0
            ColorBright     =   255
            DisplayHand     =   0   'False
            ColorScheme     =   4
         End
         Begin Project1.USStyle USStyle4 
            Height          =   375
            Left            =   3840
            TabIndex        =   7
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Help?"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
         Begin Project1.USStyle USStyle3 
            Height          =   375
            Left            =   2520
            TabIndex        =   8
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "View"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
         Begin Project1.USStyle USStyle2 
            Height          =   375
            Left            =   1200
            TabIndex        =   9
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Edit"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
         Begin Project1.USStyle USStyle1 
            Height          =   375
            Left            =   -120
            TabIndex        =   10
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "File"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Style           =   6
            Checked         =   0   'False
            ColorButtonHover=   40960
            ColorButtonUp   =   32768
            ColorButtonDown =   49152
            BorderBrightness=   0
            ColorBright     =   65280
            DisplayHand     =   0   'False
            ColorScheme     =   5
         End
      End
   End
End
Attribute VB_Name = "FrmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub USStyle1_Click()
Picture2.Visible = False
If Picture2.Visible = True Then
Picture2.Visible = False
End If

End Sub

Private Sub USStyle1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
End Sub

Private Sub USStyle10_Click()
Picture2.Visible = False
End Sub

Private Sub USStyle11_Click()
Picture2.Visible = False
End Sub

Private Sub USStyle12_Click()
Picture2.Visible = False
End Sub

Private Sub USStyle13_Click()
Picture2.Visible = False
End Sub

Private Sub USStyle15_Click()
Picture2.Visible = False

End Sub

Private Sub USStyle16_Click()
Picture3.Visible = False
End Sub

Private Sub USStyle17_Click()
Picture3.Visible = False
End Sub

Private Sub USStyle18_Click()
Picture3.Visible = False
End Sub

Private Sub USStyle19_Click()
Picture3.Visible = False
End Sub

Private Sub USStyle2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
Picture2.Visible = False
Picture4.Visible = False
End Sub

Private Sub USStyle3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture4.Visible = True
Picture3.Visible = False
Picture2.Visible = False
End Sub

Private Sub USStyle4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture4.Visible = False
Picture3.Visible = False
Picture2.Visible = False
End Sub

Private Sub USStyle7_Click()
INS = MsgBox("Are you sure to Quit EXE Clinic Zone?", vbQuestion + vbYesNo)
If isn = vbYes Then
Unload Me
Else
Exit Sub
End If
End Sub
