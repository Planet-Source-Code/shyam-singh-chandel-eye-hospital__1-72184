VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSystemSettings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "System Settings"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DSN for Access"
      TabPicture(0)   =   "Form2.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DSN for SQL"
      TabPicture(1)   =   "Form2.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   6495
         Begin VB.CommandButton cmdHelp 
            Caption         =   "?"
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   35
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "System DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   28
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "User DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "File DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   26
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   25
            Text            =   "SQL Server"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Text            =   "MIDSNSQL"
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   23
            Text            =   "pubs"
            Top             =   1560
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   4
            Left            =   1560
            TabIndex        =   22
            Text            =   "THIS IS A TEST"
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   21
            Text            =   "(local)"
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Driver:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DSN Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Database:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   31
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Server: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6495
         Begin VB.CommandButton cmdHelp 
            Caption         =   "?"
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option0 
            Caption         =   "File DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   15
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   14
            Text            =   "Microsoft Access Driver (*.mdb)"
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   13
            Text            =   "US EYE CLINIC ZONE"
            Top             =   600
            Width           =   3375
         End
         Begin VB.OptionButton Option0 
            Caption         =   "User DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton Option0 
            Caption         =   "System DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   11
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   10
            Text            =   "C:\MIBDD.MDB"
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   9
            Text            =   "US EYE CLINIC ZONE"
            Top             =   1680
            Width           =   3375
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Choose ..."
            Height          =   255
            Left            =   5040
            TabIndex        =   8
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Driver:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DSN Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Database:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1260
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   6735
      Begin VB.CommandButton cmdShowOM 
         Caption         =   "Show ODBC Manager"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "VerifyDSN"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DSN System Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   36
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "FrmSystemSettings"
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


Private Declare Function SQLManageDataSources Lib "ODBCCP32.DLL" (ByVal hWnd As Long) As Long   'Show ODBC Manager
Private Declare Function SQLValidDSN Lib "ODBCCP32.DLL" (ByVal lpszDSN As String) As Long

Private Sub cmdCreate_Click()
Dim oOdbc As New clsOdbc
Dim result As String
Dim strAttr As String
Dim m_tab As Integer
m_tab = SSTab1.Tab
If m_tab = 0 Then 'Type Access
    strAttr = strAttr & "DSN=" & Text1(1).Text & Chr$(0)
    strAttr = strAttr & "DBQ=" & Text1(2).Text & Chr$(0)
    strAttr = strAttr & "DESCRIPTION=" & Text1(3).Text & Chr$(0)
    If Option0(0) Then
        'Example for any driver
        oOdbc.ODBC_DRIVER_NAME = Text1(0).Text
        oOdbc.ODBC_ATTRIBUTES = strAttr
        result = oOdbc.ExecuteDSN(CreateUserDSN)
        'Or you can use:
        'result = oOdbc.ExecuteDSN(CreateUserDSN, MSAccess, strAttr)
    ElseIf Option0(1) Then
        result = oOdbc.ExecuteDSN(CreateSystemDSN, MSAccess, strAttr)
    ElseIf Option0(2) Then
        strAttr = "FIL=MS Access|DBQ=" & Text1(2).Text & "|UID=sa|Description=" & Text1(3).Text
        result = oOdbc.ExecuteFileDSN(Text1(1).Text, MSAccess, strAttr)
    End If
ElseIf m_tab = 1 Then 'type SQL Server
    strAttr = strAttr & "DSN=" & Text2(1).Text & Chr$(0)
    strAttr = strAttr & "SERVER=" & Text2(2).Text & Chr$(0)
    strAttr = strAttr & "DATABASE=" & Text2(3).Text & Chr$(0)
    strAttr = strAttr & "DESCRIPTION=" & Text2(4).Text & Chr$(0)
    If Option1(0) Then
        result = oOdbc.ExecuteDSN(CreateUserDSN, SQLServer, strAttr)
    ElseIf Option1(1) Then
        result = oOdbc.ExecuteDSN(CreateSystemDSN, SQLServer, strAttr)
    ElseIf Option1(2) Then
        strAttr = "SERVER=" & Text2(2).Text & "|DATABASE=" & Text2(3).Text & "|UID=sa"
        result = oOdbc.ExecuteFileDSN(Text2(1).Text, SQLServer, strAttr)
    End If
End If
If result <> "" Then MsgBox result, vbExclamation, "Error" Else MsgBox "Done"
Set oOdbc = Nothing
End Sub


Private Sub cmdHelp_Click(Index As Integer)
MsgBox "You need fill all values for create the DSN" & vbCrLf & _
        "After you can modify the DSN created" & vbCrLf & _
        "Finally you can delete the DSN created." & vbCrLf & vbCrLf & _
        "For Delete is necesary the Driver and DSN Name"
End Sub

Private Sub cmdModify_Click()
Dim oOdbc As New clsOdbc
Dim result As String
Dim strAttr As String
Dim m_tab As Integer
m_tab = SSTab1.Tab
If m_tab = 0 Then 'Type Access
    strAttr = strAttr & "DSN=" & Text1(1).Text & Chr$(0)
    strAttr = strAttr & "DBQ=" & Text1(2).Text & Chr$(0)
    strAttr = strAttr & "DESCRIPTION=" & Text1(3).Text & Chr$(0)
    If Option0(0) Then
        result = oOdbc.ExecuteDSN(ModifyUserDSN, MSAccess, strAttr)
    ElseIf Option0(1) Then
        result = oOdbc.ExecuteDSN(ModifySystemDSN, MSAccess, strAttr)
    ElseIf Option0(2) Then
        strAttr = "FIL=MS Access|DBQ=" & Text1(2).Text & "|UID=sa|Description=" & Text1(3).Text
        result = oOdbc.ExecuteFileDSN(Text1(1).Text, MSAccess, strAttr)
    End If
ElseIf m_tab = 1 Then 'type SQL Server
    strAttr = strAttr & "DSN=" & Text2(1).Text & Chr$(0)
    strAttr = strAttr & "SERVER=" & Text2(2).Text & Chr$(0)
    strAttr = strAttr & "DATABASE=" & Text2(3).Text & Chr$(0)
    strAttr = strAttr & "DESCRIPTION=" & Text2(4).Text & Chr$(0)
    If Option1(0) Then
        result = oOdbc.ExecuteDSN(ModifyUserDSN, SQLServer, strAttr)
    ElseIf Option1(1) Then
        result = oOdbc.ExecuteDSN(ModifySystemDSN, SQLServer, strAttr)
    ElseIf Option1(2) Then
        strAttr = "SERVER=" & Text2(2).Text & "|DATABASE=" & Text2(3).Text & "|UID=sa"
        result = oOdbc.ExecuteFileDSN(Text2(1).Text, SQLServer, strAttr)
    End If
End If
If result <> "" Then MsgBox result, vbExclamation, "Error" Else MsgBox "Done"
Set oOdbc = Nothing
End Sub

Private Sub cmdDelete_Click()
Dim oOdbc As New clsOdbc
Dim result As String
Dim strAttr As String
Dim m_tab As Integer
m_tab = SSTab1.Tab
If m_tab = 0 Then 'Type Access
    strAttr = strAttr & "DSN=" & Text1(1).Text & Chr$(0)
    If Option0(0) Then
        result = oOdbc.ExecuteDSN(DeleteUserDSN, MSAccess, strAttr)
    ElseIf Option0(1) Then
        result = oOdbc.ExecuteDSN(DeleteSystemDSN, MSAccess, strAttr)
    ElseIf Option0(2) Then 'Use regedit for delete
        strAttr = "FIL=MS Access|DBQ=" & Text1(2).Text & "|UID=sa|Description=" & Text1(3).Text
        result = oOdbc.ExecuteFileDSN(Text1(1).Text, MSAccess, strAttr)
    End If
ElseIf m_tab = 1 Then 'type SQL Server
    strAttr = strAttr & "DSN=" & Text2(1).Text & Chr$(0)
    If Option1(0) Then
        result = oOdbc.ExecuteDSN(DeleteUserDSN, SQLServer, strAttr)
    ElseIf Option1(1) Then
        result = oOdbc.ExecuteDSN(DeleteSystemDSN, SQLServer, strAttr)
    ElseIf Option1(2) Then 'User regedit for delete
        strAttr = "SERVER=" & Text2(2).Text & "|DATABASE=" & Text2(3).Text & "|UID=sa"
        result = oOdbc.ExecuteFileDSN(Text2(1).Text, SQLServer, strAttr)
    End If
End If
If result <> "" Then MsgBox result, vbExclamation, "Error" Else MsgBox "Done"
Set oOdbc = Nothing
End Sub

Private Sub cmdShowOM_Click()
    SQLManageDataSources (Me.hWnd)
End Sub

Private Sub cmdVerify_Click()
    Dim s_dsn As String
    If SSTab1.Tab = 0 Then
        s_dsn = Text1(1).Text
    ElseIf SSTab1.Tab = 1 Then
        s_dsn = Text2(1).Text
    ElseIf SSTab1.Tab = 2 Then
        's_dsn = Text3(1).Text
    End If
    If SQLValidDSN(s_dsn) Then MsgBox "DSN OK" Else MsgBox "DSN FALIDED"
End Sub

Private Sub Command4_Click()
    With CommonDialog1
        .Filter = "Access Database (*.mdb)|*.mdb"
        .ShowOpen
        If .FileName = "" Then
'            MsgBox ("You press cancel")
            Exit Sub
        End If
        Text1(2).Text = .FileName
    End With
End Sub

Private Sub Form_Load()
On Error Resume Next

Me.BackColor = GetSetting("USEYE", "Settings", "Back Color")
Text1(2).Text = MainPath & "\DATA\USECZ.mdb"

End Sub
