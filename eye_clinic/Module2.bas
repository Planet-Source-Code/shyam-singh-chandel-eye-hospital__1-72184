Attribute VB_Name = "CrtData"
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

  ' Sawan 9862151523

Option Explicit
Option Compare Text
Global RunSt As String
Global MainPath As String
Global RestoPath As String
Global StaffPath As String
Global CustomerPath As String
Global ItemsPath As String
Global PrintPath As String
Global scount As String
Global USER_LOG As String
Global SCRTIME As Integer
Global CName As String
Global CAddress As String
Global USMonth As String
Global USDate As String
Global USYear As String
Global USMedicine As String
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#If Win32 Then
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_FINDSTRINGEXACT = &H158
    Public Const LB_FINDSTRING = &H18F
    Public Const LB_FINDSTRINGEXACT = &H1A2
#Else
    Public Const WM_USER = &H400
    Public Const CB_FINDSTRING = WM_USER + 12
    Public Const CB_FINDSTRINGEXACT = WM_USER + 24
    Public Const LB_FINDSTRING = WM_USER + 16
    Public Const LB_FINDSTRINGEXACT = WM_USER + 35
#End If

Sub MAINDB(psDBfile As String)

    Dim t As Integer
    Dim MSG As String
    Dim DB As Database
    Dim dyna As Dynaset
    
    On Error GoTo ssc8
    Set DB = CreateDatabase(psDBfile, dbLangGeneral)

    t% = CreatTable(DB)
    t% = CreatTableM(DB)
   ' t% = CreatTable2(DB)
    
    If t% <> 0 Then Error (t%)

    DB.Close
    Set DB = Nothing
    Exit Sub

ssc8:

    If Err <> 3204 Then
        MSG = "Got error " & Err & " (" & ERROR$ & ") while trying to create user database!"
        MsgBox MSG, 64
    End If
    Exit Sub

End Sub

Sub ROUTINE(psDBfile As String)

    Dim t As Integer
    Dim MSG As String
    Dim DB As Database
    Dim dyna As Dynaset

    On Error GoTo ssc8
    Set DB = CreateDatabase(psDBfile, dbLangGeneral)

    t% = CreatTable2(DB)
        
    If t% <> 0 Then Error (t%)

    DB.Close
    Set DB = Nothing
    Exit Sub

ssc8:

    If Err <> 3204 Then
        MSG = "Got error " & Err & " (" & ERROR$ & ") while trying to create user database!"
        MsgBox MSG, 64
    End If
    Exit Sub

End Sub

Sub USERLOG(psDBfile As String)

    Dim t As Integer
    Dim MSG As String
    Dim DB As Database
    Dim dyna As Dynaset

    On Error GoTo ssc7
    Set DB = CreateDatabase(psDBfile, dbLangGeneral)

    t% = CreatTable3(DB)
    If t% <> 0 Then Error (t%)
    
    DB.Close
    Set DB = Nothing
    Exit Sub

ssc7:

    If Err <> 3204 Then
        MSG = "Got error " & Err & " (" & ERROR$ & ") while trying to create user database!"
        MsgBox MSG, 64
    End If
    Exit Sub

End Sub

Function CreatTable(DB As Database) As Integer

        On Error GoTo ssc
    Dim Qua As QueryDef
    Dim TD As New TableDef, fld() As New Field
    Dim idx() As New Index, i As Integer
    
    ReDim fld(1 To 44)
    ReDim idx(1 To 4)
    
    TD.Name = "USTestVisit"
        
    fld(1).Attributes = dbAutoIncrField
    For i = 1 To 44
        Select Case i
            Case 1:
                fld(i).Name = "ID"
                fld(i).Type = dbLong
           Case 2:
                fld(i).Name = "Remark"
                fld(i).Type = dbText
                fld(i).Size = 30
            Case 3:
                fld(i).Name = "Name"
                fld(i).Type = dbText
                fld(i).Size = 50
                
            Case 4:
                fld(i).Name = "TmpAddress"
                fld(i).Type = dbText
                fld(i).Size = 200
            Case 5:
                fld(i).Name = "PerAddress"
                fld(i).Type = dbText
                fld(i).Size = 200
            Case 6:
                fld(i).Name = "City"
                fld(i).Type = dbText
                fld(i).Size = 30
            Case 7:
                fld(i).Name = "Age"
                fld(i).Type = dbText
                fld(i).Size = 5
            Case 8:
                fld(i).Name = "Sex"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 9:
                fld(i).Name = "Tel"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 10:
                fld(i).Name = "Email"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 11:
                fld(i).Name = "Treatment"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 12:
                fld(i).Name = "LastVisit"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 13:
                fld(i).Name = "NextVisit"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 14:
                fld(i).Name = "Type"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 15:
                fld(i).Name = "OldID"
                fld(i).Type = dbText
                fld(i).Size = 20
           Case 16:
                fld(i).Name = "Doctor"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 17:
                fld(i).Name = "Reg"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 18:
                fld(i).Name = "Con"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 19:
                fld(i).Name = "AR"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 20:
                fld(i).Name = "TN"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 21:
                fld(i).Name = "Fun"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 22:
                fld(i).Name = "SL"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 23:
                fld(i).Name = "OTHERS"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 24:
                fld(i).Name = "IDO"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 25:
                fld(i).Name = "GONIOS"
                fld(i).Type = dbText
                fld(i).Size = 80
            Case 26:
                fld(i).Name = "AscanKerometry"
                fld(i).Type = dbText
                fld(i).Size = 80
            Case 27:
                fld(i).Name = "PERITYPE"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 28:
                fld(i).Name = "PERIAMOUNT"
                fld(i).Type = dbText
                fld(i).Size = 20
            
            Case 29:
                fld(i).Name = "MINOROTTYPE"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 30:
                fld(i).Name = "MINOROTAMOUNT"
                fld(i).Type = dbText
                fld(i).Size = 20
                
            Case 31:
                fld(i).Name = "MAJOROTTYPE"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 32:
                fld(i).Name = "MAJOROTAMOUNT"
                fld(i).Type = dbText
                fld(i).Size = 20
                
            Case 33:
                fld(i).Name = "SHIRMER"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 34:
                fld(i).Name = "RBS"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 35:
                fld(i).Name = "FFA"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 36:
                fld(i).Name = "BTCT"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 37:
                fld(i).Name = "DIAGNOSIS"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 38:
                fld(i).Name = "MEDICINES"
                fld(i).Type = dbText
                fld(i).Size = 200
            Case 39:
                fld(i).Name = "BALANCE"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 40:
                fld(i).Name = "TOTAL"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 41:
                fld(i).Name = "PAID"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 42:
                fld(i).Name = "YearF"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 43:
                fld(i).Name = "DateF"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 44:
                fld(i).Name = "MonthF"
                fld(i).Type = dbText
                fld(i).Size = 50
             End Select
      TD.Fields.Append fld(i)
    Next i

    
    idx(1).Name = "MonthF"
    idx(1).Fields = "MonthF"
    

    For i = 1 To 1
        TD.Indexes.Append idx(i)
    Next i

    DB.TableDefs.Append TD
    CreatTable = 0
    Exit Function

ssc:
    CreatTable = Err
    Exit Function

End Function

Function CreatTableM(DB As Database) As Integer

        On Error GoTo ssc4

    Dim TD As New TableDef, fld() As New Field
    Dim idx() As New Index, i As Integer
        
    ReDim fld(1 To 8)
    ReDim idx(1 To 8)
    
    TD.Name = "MEDICINE"
        
    fld(1).Attributes = dbAutoIncrField
    For i = 1 To 8
        Select Case i
            Case 1:
                fld(i).Name = "ID"
                fld(i).Type = dbLong
            Case 2:
                fld(i).Name = "REGNO"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 3:
                fld(i).Name = "MEDICINE"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 4:
                fld(i).Name = "QTY"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 5:
                fld(i).Name = "AMOUNT"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 6:
                fld(i).Name = "DATE"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 7:
                fld(i).Name = "MONTH"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 8:
                fld(i).Name = "YEAR"
                fld(i).Type = dbText
                fld(i).Size = 20
            End Select
      TD.Fields.Append fld(i)
    Next i
    
    idx(1).Name = "REGNO"
    idx(1).Fields = "REGNO"
    
    For i = 1 To 1
        TD.Indexes.Append idx(i)
    Next i

     DB.TableDefs.Append TD
    CreatTableM = 0
    Exit Function

ssc4:
    CreatTableM = Err
    Exit Function

End Function




Function CreatTable2(DB As Database) As Integer

        On Error GoTo ssc5

    Dim TD As New TableDef, fld() As New Field
    Dim idx() As New Index, i As Integer
        
    ReDim fld(1 To 19)
    ReDim idx(1 To 4)
    
    TD.Name = "AROUTINE"
        
    fld(1).Attributes = dbAutoIncrField
    For i = 1 To 19
        Select Case i
            Case 1:
                fld(i).Name = "ID"
                fld(i).Type = dbLong
            Case 2:
                fld(i).Name = "Name"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 3:
                fld(i).Name = "Address"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 4:
                fld(i).Name = "Tel"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 5:
                fld(i).Name = "Email"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 6:
                fld(i).Name = "Section"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 7:
                fld(i).Name = "MINOROTType"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 8:
                fld(i).Name = "MINOROTRates"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 9:
                fld(i).Name = "MAJOROTType"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 10:
                fld(i).Name = "MAJOROTRates"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 11:
                fld(i).Name = "PariType"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 12:
                fld(i).Name = "PariAMOUNT"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 13:
                fld(i).Name = "DOCTOR"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 14:
                fld(i).Name = "DOCTORCHARGES"
                fld(i).Type = dbText
                fld(i).Size = 10
            Case 15:
                fld(i).Name = "GIAGNOSIS"
                fld(i).Type = dbText
                fld(i).Size = 100
            Case 16:
                fld(i).Name = "SUBJECT"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 17:
                fld(i).Name = "A"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 18:
                fld(i).Name = "B"
                fld(i).Type = dbText
                fld(i).Size = 50
            Case 19:
                fld(i).Name = "C"
                fld(i).Type = dbText
                fld(i).Size = 50
            
            End Select
      TD.Fields.Append fld(i)
    Next i

    
    idx(1).Name = "Name"
    idx(1).Fields = "Name"
   
    For i = 1 To 1
        TD.Indexes.Append idx(i)
    Next i

    DB.TableDefs.Append TD
    CreatTable2 = 0
    Exit Function

ssc5:
    CreatTable2 = Err
    Exit Function

End Function
Function CreatTable3(DB As Database) As Integer

        On Error GoTo ssc4

    Dim TD As New TableDef, fld() As New Field
    Dim idx() As New Index, i As Integer
        
    ReDim fld(1 To 4)
    ReDim idx(1 To 4)
    
    TD.Name = "USERID"
        
    fld(1).Attributes = dbAutoIncrField
    For i = 1 To 4
        Select Case i
            Case 1:
                fld(i).Name = "ID"
                fld(i).Type = dbLong
            Case 2:
                fld(i).Name = "USER"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 3:
                fld(i).Name = "PASS"
                fld(i).Type = dbText
                fld(i).Size = 20
            Case 4:
                fld(i).Name = "PASS2"
                fld(i).Type = dbText
                fld(i).Size = 20
                           
            End Select
      TD.Fields.Append fld(i)
    Next i
    
    idx(1).Name = "USER"
    idx(1).Fields = "USER"
    
    For i = 1 To 1
        TD.Indexes.Append idx(i)
    Next i

     DB.TableDefs.Append TD
    CreatTable3 = 0
    Exit Function

ssc4:
    CreatTable3 = Err
    Exit Function

End Function
Public Function USERDB(DBFullPath As String) As Boolean
Dim DB As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set DB = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = DB.CreateTableDef("USERDB")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("USERNAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("PASS", dbText)
TD.Fields.Append f


DB.TableDefs.Append TD ''

USERDB = True
ErrorHandler:
If Not DB Is Nothing Then DB.Close

End Function
Public Function FindFirstMatch(ByVal ctlSearch As Control, ByVal SearchString As String, ByVal FirstRow As Integer, ByVal Exact As Boolean) As Integer

#If Win32 Then
    Dim Index As Long
#Else
    Dim Index As Integer
#End If

On Error Resume Next
If TypeOf ctlSearch Is ComboBox Then
    If Exact Then
        Index = SendMessage(ctlSearch.hWnd, CB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hWnd, CB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
ElseIf TypeOf ctlSearch Is ListBox Then
    If Exact Then
        Index = SendMessage(ctlSearch.hWnd, LB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hWnd, LB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
End If

FindFirstMatch = Index

End Function

Public Function FileExist(asPath As String) As Boolean
  On Error Resume Next
  
    If UCase(Dir(asPath)) = UCase(TrimPath(asPath)) Then
      FileExist = True
    Else
      FileExist = False
    End If
End Function
Public Function TrimPath(ByVal asPath As String) As String
 On Error Resume Next
     If Len(asPath) = 0 Then Exit Function
    Dim x As Integer
     Do
        x = InStr(asPath, "\")
        If x = 0 Then Exit Do
        asPath = Right(asPath, Len(asPath) - x)
    Loop
    TrimPath = asPath
End Function


