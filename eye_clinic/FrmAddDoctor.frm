VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAddDoctor 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Doctors"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   Icon            =   "FrmAddDoctor.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6720
      TabIndex        =   14
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      Picture         =   "FrmAddDoctor.frx":0CCA
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
      Left            =   8160
      Top             =   5400
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1695
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   21
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin MSComctlLib.ListView List 
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sl"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Doctor Name"
         Object.Width           =   5291
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   3530
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tel"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Email"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Section"
         Object.Width           =   1764
      EndProperty
   End
   Begin Project1.USStyle USStyle2 
      Height          =   615
      Left            =   4800
      TabIndex        =   15
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "   Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmAddDoctor.frx":15DC
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
   Begin Project1.USStyle USStyle3 
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "      Save"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmAddDoctor.frx":22B6
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
   Begin Project1.USStyle USStyle4 
      Height          =   615
      Left            =   2520
      TabIndex        =   17
      Top             =   5880
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
      Caption         =   "       Delete"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmAddDoctor.frx":25D0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Doctor's Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label O 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Name"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
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
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Index           =   3
      Left            =   3840
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
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
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11895
   End
End
Attribute VB_Name = "FrmAddDoctor"
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
Dim M As sPrint
Private Sub Form_Load()
On Error Resume Next

Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
' set up the database connectivity for the ADO data control
    With Adodc1

        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            MainPath & "\DATA\USROUTINE.mdb;Persist Security Info=False"

        .RecordSource = "select * from AROUTINE" ' order by Name , Address , Tel ,Email, Section"
    End With

    ' set the Flex Grid data source to be the ADO data control.
    Set MSHFlexGrid1.DataSource = Adodc1

    ' set up the format string for the flex grid.
    MSHFlexGrid1.FormatString = "Doctor Name | Address | Tel      |Email                   |Section                    "

    ' position all the controls happily and store the form minimum size
    MinHeight = Me.Height
    MinWidth = Me.Width
    'Call Form_Resize
          If Right(MainPath, 1) = "\" Then
        Set DB = OpenDatabase(MainPath + "\DATA" & "USROUTINE.mdb")
    Else
        Set DB = OpenDatabase(MainPath + "\DATA\USROUTINE.mdb")
    End If
        Set Rs = DB.OpenRecordset("AROUTINE")
        
    Set M = New sPrint
Set GPrint.ListViewName = List
GPrint.DrawHorizontalLines = True
GPrint.DrawVerticalLines = True
GPrint.DrawBorder = True
GPrint.BorderDistance = 2
GPrint.PosX = 300    'Value in Twips
GPrint.PosY = 1400  'Value in Twips
GPrint.HasPicture = True
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = vbBlack
Text2.ForeColor = vbWhite
End Sub


Private Sub Text3_GotFocus()
Text3.BackColor = vbBlack
Text3.ForeColor = vbWhite
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = vbBlack
Text4.ForeColor = vbWhite
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = vbBlack
Text5.ForeColor = vbWhite
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'Text2.SetFocus
 On Error Resume Next
        FindsOldID Text1, List, DB, "AROUTINE", "Name"
        MSHFlexGrid1.Visible = False
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub USStyle1_Click()
Text6.Text = "1"
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
Printer.Orientation = vbPRORPortrait  'vbPRORLandscape
Printer.Font = List.Font.Name
Printer.FontSize = List.Font.Size
While Not GPrint.LastRowPrinted
        Page = Page + 1
        GPrint.SetRows
        Printer.CurrentX = 2000
        Printer.CurrentY = 60: Printer.FontSize = 19: Printer.FontName = "Times New Roman"
        Printer.Print CName
        Printer.FontSize = 8: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 2300
        Printer.CurrentY = 430: Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.Print CAddress
        Printer.FontSize = 8: Printer.FontName = "Courier New"
        Printer.CurrentX = 400
        Printer.CurrentY = 900: Printer.FontSize = 14: Printer.FontName = "Times New Roman"
        Printer.Print "Display Doctor's List."
        Printer.FontSize = 10: Printer.FontName = "Arial"
                                                                          
        GPrint.PrintHead Printer
        GPrint.PrintBody Printer
        Printer.FontSize = 14: Printer.FontName = "Arial"
        Printer.CurrentY = 900
        Printer.CurrentX = 400
        Printer.CurrentX = 12500
        Printer.Print "Page = " & Text6.Text  '& '" Page."         & Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
        
        Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 500
        Printer.CurrentY = 12300: Printer.FontBold = True
        Printer.CurrentY = 450
        Printer.CurrentX = 400
        Printer.CurrentX = 12300
        Printer.Print "Date:- " & Format(Date, "dd/mm/yy")

        Printer.NewPage
        Text6.Text = Val(Text6.Text) + 1
Wend
        Printer.EndDoc
        GPrint.LastRowPrinted = False
        Me.ScaleMode = vbTwips
End Sub

Private Sub USStyle2_Click()
Unload Me
End Sub

Private Sub USStyle3_Click()
' add a new entry to our table.
    With Adodc1.Recordset
        .AddNew
        !Name = Text1.Text
        !Address = Text2.Text
        !Tel = Text3.Text
        !Email = Text4.Text
        !Section = Text5.Text
        .Update
        .Requery
    End With
        
    ' refresh the data source and rebind it to the flexgrid (annoying!!)
    Adodc1.Refresh
    Set MSHFlexGrid1.DataSource = Adodc1
    MSHFlexGrid1.FormatString = "Doctor Name | Address | Tel      |Email                   |Section                    "
    
    ' clear the text fields once the new record is added
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    'Text6 = ""
    ' set the focus back to the  artist name textbox
    Text1.SetFocus
    MSHFlexGrid1.Visible = True
    MsgBox "Doctor Record is Saved"
    
End Sub

Public Function FindsOldID(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
Label5.Caption = "0"
sl = 1
            'Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            Text6.Text = ""
           
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
            Set Li = List.ListItems.Add(, , Format(sl, "00"))
            Li.SubItems(1) = sTemp!Name
            Li.SubItems(2) = sTemp!Address
            Li.SubItems(3) = sTemp!Tel
            Li.SubItems(4) = sTemp!Email
            Li.SubItems(5) = sTemp!Section
           
            Text1.Text = sTemp!Name
            Text2.Text = sTemp!Address
            Text3.Text = sTemp!Tel
            Text4.Text = sTemp!Email
            Text5.Text = sTemp!Section
            sl = sl + 1
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

Private Sub USStyle4_Click()
With Adodc1.Recordset
        .Move (MSHFlexGrid1.row - 1) ' we minus one because row zero is the header row
        .Delete
        .Requery
    End With
    
    ' refresh the data source and rebind it to the flexgrid (annoying!!)
    Adodc1.Refresh
    Set MSHFlexGrid1.DataSource = Adodc1
    MSHFlexGrid1.FormatString = "Doctor Name | Address | Tel      |Email                   |Section                    "
        
    'set the focus back to the first add field
    Text1.SetFocus
End Sub
