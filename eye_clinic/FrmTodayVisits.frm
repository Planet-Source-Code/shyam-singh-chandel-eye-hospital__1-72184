VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmTodayVisits 
   BackColor       =   &H0080FF80&
   Caption         =   "Today's Visits"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   Icon            =   "FrmTodayVisits.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ListView List 
      Height          =   5055
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Old Registration"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "New Registration"
         Object.Width           =   3527
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   5118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Diagnosis"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Last Visit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Next Visit"
         Object.Width           =   2205
      EndProperty
   End
   Begin Project1.USStyle USStyle1 
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "   Print"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmTodayVisits.frx":0CCA
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
   Begin Project1.USStyle USStyle2 
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmTodayVisits.frx":15DC
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
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TODAY'S TOTAL VISITS:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Today's Visits  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "FrmTodayVisits"
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


Private Sub USStyle1_Click()
Text1.Text = "1"
Dim d
Dim Page As Integer
Dim sngTotalPage As Single
GPrint.NumOfRowsPerPage = 20
GPrint.RowHeight = 14 * 30

sngTotalPage = List.ListItems.Count / GPrint.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1
Me.ScaleMode = vbPixels 'this must be done, the container [LEDGER in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA4    ' vbPRPSA5
'Printer.PrintQuality = vbPRPQHigh
Printer.Orientation = vbPRORLandscape         '         vbPRORPortrait
Printer.Font = List.Font.Name
Printer.FontSize = List.Font.Size
While Not GPrint.LastRowPrinted
        Page = Page + 1
        GPrint.SetRows
        Printer.CurrentX = 3700
        Printer.CurrentY = 60: Printer.FontSize = 19: Printer.FontName = "Times New Roman"
        Printer.Print CName
        Printer.FontSize = 8: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 4000
        Printer.CurrentY = 430: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print CAddress
        Printer.FontSize = 8: Printer.FontName = "Courier New"
        Printer.CurrentX = 400
        Printer.CurrentY = 900: Printer.FontSize = 14: Printer.FontName = "Times New Roman"
        Printer.Print "Display Patient Report " & CboSerBy.Text & " wise."
        Printer.FontSize = 10: Printer.FontName = "Arial"
                                                                          
        GPrint.PrintHead Printer
        GPrint.PrintBody Printer
        Printer.FontSize = 14: Printer.FontName = "Arial"
        Printer.CurrentY = 900
        Printer.CurrentX = 400
        Printer.CurrentX = 12500
        Printer.Print "Page = " & Text1.Text  '& '" Page."         & Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
        
        Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 500
        Printer.CurrentY = 12300: Printer.FontBold = True
        Printer.CurrentY = 450
        Printer.CurrentX = 400
        Printer.CurrentX = 12300
        Printer.Print "Date:- " & Format(Date, "dd/mm/yy")

        Printer.NewPage
        Text1.Text = Val(Text1.Text) + 1
Wend
        Printer.EndDoc
        GPrint.LastRowPrinted = False
        Me.ScaleMode = vbTwips
End Sub

Private Sub USStyle2_Click()
Unload Me

End Sub
