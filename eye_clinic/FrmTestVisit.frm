VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmTestVisit 
   BackColor       =   &H0080FF80&
   Caption         =   "Form4"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "FrmTestVisit.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      ScaleHeight     =   375
      ScaleWidth      =   1695
      TabIndex        =   15
      Top             =   8040
      Width           =   1695
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
         Height          =   615
         Left            =   0
         TabIndex        =   16
         Top             =   100
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFF00&
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   11355
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while Finding Record. . . . . . . ."
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
         Left            =   0
         TabIndex        =   13
         Top             =   2640
         Width           =   11055
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11760
      TabIndex        =   8
      Text            =   "1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
   End
   Begin Project1.USStyle USStyle1 
      Height          =   615
      Left            =   7800
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "FrmTestVisit.frx":0CCA
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox CboSerBy 
      Height          =   315
      ItemData        =   "FrmTestVisit.frx":15DC
      Left            =   240
      List            =   "FrmTestVisit.frx":15DE
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid FlexResult 
      Height          =   5535
      Left            =   12880
      TabIndex        =   4
      Top             =   2880
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   30
      Cols            =   7
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   2
      Appearance      =   0
      FormatString    =   $"FrmTestVisit.frx":15E0
   End
   Begin MSComctlLib.ListView List 
      Height          =   5535
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9763
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
      NumItems        =   34
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Registration"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Other"
         Object.Width           =   3527
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   5118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Age"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sex"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tmp. Address"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Per.Add"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tel"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Reg"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Con"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "A/R"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "TN"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Fun"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "S/L"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Others"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "IDO"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Gon"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "A/K"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "P. Type"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "P Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Min OT"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Major OT"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Amt"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Shir"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "R.B.A"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "F.F.A"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "BT/CT"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Diagnosis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Medicines"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Remark"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Total"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Paid"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.USStyle USStyle2 
      Height          =   615
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
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
      Picture         =   "FrmTestVisit.frx":1718
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
      Left            =   5160
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "     Register"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FrmTestVisit.frx":23F2
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   7995
      Width           =   11880
      _ExtentX        =   20955
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Test and Visits Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Searching word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu export 
         Caption         =   "Export Report"
      End
      Begin VB.Menu print 
         Caption         =   "Print"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmTestVisit"
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
Dim GPrint As sPrint
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

Private Sub CboSerBy_Click()
Text1.Text = ""
End Sub

Private Sub CboSerBy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text1.SetFocus
End If

End Sub

Private Sub export_Click()
'Early object binding
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
'Uncomment below for late object binding
'Dim oWord As Object
'Dim oDoc As Object
'Dim oRange As Object
Dim row As Integer
Dim col As Integer
Dim i As Integer
Dim n As Integer
Dim sTemp As String
Dim arr() As String
'i = FlexResult.Rows
'n = FlexResult.Cols
'MsgBox i
'MsgBox n

'ReDim arr(FlexResult.Rows - 1, FlexResult.Cols - 1)
ReDim arr(FlexResult.Rows - 1, FlexResult.Cols - 1)
'Create an instance of Word
Set oWord = CreateObject("Word.Application")

'Show Word to the user
oWord.Visible = True

'Add a new, blank document
Set oDoc = oWord.Documents.Add

'Get the current document's range object

'Store FlexGrid items to a two dimensional array
For row = 0 To FlexResult.Rows - 1
    n = 0
    For col = 0 To FlexResult.Cols - 1
        arr(i, n) = FlexResult.TextMatrix(row, col)
        n = n + 1
    Next
    i = i + 1
Next

'Store array items to a string
For i = LBound(arr, 1) To UBound(arr, 1)
    For n = LBound(arr, 2) To UBound(arr, 2)
        sTemp = sTemp & arr(i, n)
        If n = UBound(arr, 2) Then
           sTemp = sTemp & vbCrLf
        Else
           sTemp = sTemp & vbTab
        End If
    Next
Next

'get the current document's range object and move to end of document
Set oRange = oDoc.Bookmarks("\EndOfDoc").Range

oRange.Text = sTemp

'convert the text to a table and format the table
oRange.ConvertToTable vbTab, Format:=wdTableFormatColorful2

Set oRange = Nothing

End Sub

Private Sub Form_Load()
On Error Resume Next

Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
Set M = New sPrint
Set GPrint.ListViewName = List
GPrint.DrawHorizontalLines = True
GPrint.DrawVerticalLines = True
GPrint.DrawBorder = True
GPrint.BorderDistance = 2
GPrint.PosX = 300    'Value in Twips
GPrint.PosY = 1400  'Value in Twips
GPrint.HasPicture = True
Call MAINDB(MainPath & "\DATA\USECZ.mdb")
Call ROUTINE(MainPath & "\DATA\USROUTINE.mdb")
    CboSerBy.AddItem "Registration"
    CboSerBy.AddItem "Name"
    CboSerBy.AddItem "Sex"
    CboSerBy.AddItem "Diagnosis"
    CboSerBy.AddItem "Doctor"
    'CboSerBy.AddItem "Last Visit"
    'CboSerBy.AddItem "Next Visit"
    CboSerBy.ListIndex = 1
    If Right(MainPath, 1) = "\" Then
    Set DB = OpenDatabase(MainPath + "\DATA" & "USECZ.mdb")
Else
    Set DB = OpenDatabase(MainPath + "\DATA\USECZ.mdb")
End If
Set Rs = DB.OpenRecordset("USTestVisit")

End Sub

Private Sub Form_Unload(Cancel As Integer)
HideWithAnim
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Picture2.Visible = True
Wait (5)
On Error Resume Next
If CboSerBy.Text <> "" Then

Select Case CboSerBy.Text
    Case "Registration"
        FindPat Text1, List, DB, "USTestVisit", "OldID"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "OldID"
    Case "Sex"
        FindPat Text1, List, DB, "USTestVisit", "Sex"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "Sex"
    Case "Name"
        FindPat Text1, List, DB, "USTestVisit", "Name"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "Name"
        
    Case "Diagnosis"
        FindPat Text1, List, DB, "USTestVisit", "Treatment"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "Treatment"
    Case "Doctor"
        FindPat Text1, List, DB, "USTestVisit", "Doctor"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "Doctor"
    Case "Last Visit"
        FindPat Text1, List, DB, "USTestVisit", "LastVisit"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "LastVisit"
    Case "Next Visit"
        FindPat Text1, List, DB, "USTestVisit", "NextVisit"
        CheckRec Text1, FlexResult, DB, "USTestVisit", "NextVisit"""
    Case Else
End Select
Else
    Text1 = ""
    CboSerBy.ListIndex = 0
End If
End If
Picture2.Visible = False
End Sub

Private Sub USStyle1_Click()
Text2.Text = "1"
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
Printer.PrintQuality = vbPRPQHigh
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
        Printer.Print "Page = " & Text2.Text  '& '" Page."         & Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
        
        Printer.FontSize = 12: Printer.FontName = "Times New Roman"
        Printer.CurrentX = 500
        Printer.CurrentY = 12300: Printer.FontBold = True
        Printer.CurrentY = 450
        Printer.CurrentX = 400
        Printer.CurrentX = 12300
        Printer.Print "Date:- " & Format(Date, "dd/mm/yy")

        Printer.NewPage
        Text2.Text = Val(Text2.Text) + 1
Wend
        Printer.EndDoc
        GPrint.LastRowPrinted = False
        Me.ScaleMode = vbTwips
End Sub
Public Function FindPat(sTextbox As TextBox, sList As ListView, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
Label3.Caption = "0"
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
            sFlexGrid.Clear
            sFlexGrid.FormatString = "Old Registration    | New Registration     | Name                                                                                   |Address                                                                    | Diagnosis                                   | Last Visit               |  Next Visit              "
        Do While Not sTemp.EOF
            Set Li = List.ListItems.Add(, , sTemp!OldID)
            Li.SubItems(1) = sTemp!Type
            Li.SubItems(2) = sTemp!Name
            Li.SubItems(3) = sTemp!Age
            Li.SubItems(4) = sTemp!Sex
            Li.SubItems(5) = sTemp!TmpAddress
            Li.SubItems(6) = sTemp!perAddress
            Li.SubItems(7) = sTemp!Tel
            Li.SubItems(8) = sTemp!Reg
            Li.SubItems(9) = sTemp!Con
            Li.SubItems(10) = sTemp!AR
            Li.SubItems(11) = sTemp!TN
            Li.SubItems(12) = sTemp!fun
            Li.SubItems(13) = sTemp!sl
            Li.SubItems(14) = sTemp!OTHERS
            Li.SubItems(15) = sTemp!IDO
            Li.SubItems(16) = sTemp!GONIOS
            Li.SubItems(17) = sTemp!AscanKerometry
            Li.SubItems(18) = sTemp!PeRITYPE
            Li.SubItems(19) = sTemp!PeRIAMOUNT
            Li.SubItems(20) = sTemp!MINOROTType
            Li.SubItems(21) = sTemp!minorOTAMOUNT
            Li.SubItems(22) = sTemp!MAJOROTType
            Li.SubItems(23) = sTemp!MAJOROTAMOUNT
            Li.SubItems(24) = sTemp!SHIRMER
            Li.SubItems(25) = sTemp!RBS
            Li.SubItems(26) = sTemp!FFA
            Li.SubItems(27) = sTemp!BTCT
            Li.SubItems(28) = sTemp!DIAGNOSIS
            Li.SubItems(29) = sTemp!MEDICINES
            Li.SubItems(30) = sTemp!REMARK
            Li.SubItems(31) = sTemp!TOTAL
            Li.SubItems(32) = sTemp!PAID
            Li.SubItems(33) = sTemp!BALANCE
   
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
End Function

Private Sub USStyle2_Click()
Unload Me

End Sub

Private Sub USStyle3_Click()
FrmRegistration.Show
Unload Me
End Sub
