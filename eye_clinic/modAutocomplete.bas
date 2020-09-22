Attribute VB_Name = "modAutocomplete"
Public Function CheckRec(sTextbox As TextBox, sFlexGrid As MSFlexGrid, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
'Set AutoComplete function to FALSE
CheckRec = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
'Set OldLen as the sTextbox lenght
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
    'Here we got a bug!!
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
            sTemp.MoveFirst
            sFlexGrid.Clear
            sFlexGrid.FormatString = "Registration No   | Others     | Name                                                                                   |Address                                                                    | Diagnosis                                   | Last Visit               |  Next Visit              "
            Do While Not sTemp.EOF
                    sFlexGrid.AddItem sTemp!OldID
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 1) = sTemp!Type
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 2) = sTemp!Name
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 3) = sTemp!Address
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 4) = sTemp!Treatment
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 5) = sTemp!LastVisit
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 6) = sTemp!NextVisit
                sTemp.MoveNext
            Loop
        End If
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                CheckRec = True
    Else
        sFlexGrid.Clear
    End If
End If
sFlexGrid.FormatString = "Old Registration    | New Registration     | Name                                                                                   |Address                                                                    | Diagnosis                                   | Last Visit               |  Next Visit              "
End Function

Public Function Doctors(sTextbox As TextBox, sFlexGrid As MSFlexGrid, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
'Set AutoComplete function to FALSE
Doctors = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
'Set OldLen as the sTextbox lenght
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
    'Here we got a bug!!
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
            sTemp.MoveFirst
            sFlexGrid.Clear
            sFlexGrid.FormatString = "Old Registration    | New Registration     | Name                                                                                   |Address                                                                    | Diagnosis                                   | Last Visit               |  Next Visit              "
            Do While Not sTemp.EOF
                    sFlexGrid.AddItem sTemp!OldID
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 1) = sTemp!NewID
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 2) = sTemp!Name
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 3) = sTemp!Address
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 4) = sTemp!Treatment
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 5) = sTemp!LastVisit
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 6) = sTemp!NextVisit
                sTemp.MoveNext
            Loop
        End If
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                Doctors = True
    Else
        sFlexGrid.Clear
    End If
End If
sFlexGrid.FormatString = "Old Registration    | New Registration     | Name                                                                                   |Address                                                                    | Diagnosis                                   | Last Visit               |  Next Visit              "
End Function

