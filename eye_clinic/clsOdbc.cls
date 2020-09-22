VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOdbc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Private Declare Function SQLWriteFileDSN Lib "ODBCCP32.DLL" _
(ByVal lpszFileName As String, ByVal lpszAppName As String, _
ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Public Enum DSNType
    CreateUserDSN& = 1      ' Add User data source
    ModifyUserDSN& = 2      ' Configure existing DSN
    DeleteUserDSN& = 3      'Delete data source
'ODBC Version 2.5 & higher
    CreateSystemDSN& = 4    'Add system data source
    ModifySystemDSN& = 5    'Modify an existing system data source
    DeleteSystemDSN& = 6    'Remove an existing system data source
'ODBC Version 3.0
    'DeleteDefaultDSN& = 7 ' Remove the default data source. Experienced users only!
End Enum

Public Enum DSNDriver
    MSAccess = 1
    SQLServer = 2
End Enum

Private s_ODBC_Driver_Name    As String
Private s_ODBC_DSN_Attributes    As String

Public Property Let ODBC_DRIVER_NAME(ByVal sDriver As String)
   s_ODBC_Driver_Name = sDriver
End Property
Public Property Let ODBC_ATTRIBUTES(ByVal sAttrib As String)
   s_ODBC_DSN_Attributes = sAttrib
End Property

Private Function GetDriver(ByVal sDriver As Long) As String
   If sDriver = MSAccess Then
    GetDriver = "Microsoft Access Driver (*.mdb)"
   ElseIf sDriver = SQLServer Then
    GetDriver = "SQL Server"
   End If
End Function
Public Function ExecuteDSN(ByVal DSNType As DSNType, _
                            Optional ByVal sDriver As DSNDriver, _
                            Optional ByVal sAttributes As String) As String
On Error GoTo ExecuteDSNError
   Dim strDriver         As String
   Dim lRetVal           As Long
   Dim sAttrib           As String
   If (DSNType < CreateUserDSN) Or (DSNType > DeleteSystemDSN) Then
    ExecuteDSN = "Invalid DSN type"
    Exit Function
   End If
   
   If sDriver = 0 Then
        If (s_ODBC_Driver_Name = "") Then
            ExecuteDSN = "You need first initialize ODBC_DRIVER_NAME"
            Exit Function
        Else
            strDriver = s_ODBC_Driver_Name
        End If
   Else
        strDriver = GetDriver(sDriver)
        If strDriver = "" Then
            ExecuteDSN = "Unknow Driver" & vbCrLf & "You need specific a correct driver"
            Exit Function
        End If
   End If
   If sAttributes = "" Then  'If the user write your personalizate attributes
        If (s_ODBC_DSN_Attributes = "") Then
            ExecuteDSN = "You need first initialize ODBC_ATTRIBUTES"
            Exit Function
        Else
            sAttributes = s_ODBC_DSN_Attributes
        End If
   End If
   lRetVal = SQLConfigDataSource(0&, DSNType, strDriver, sAttributes)
   If lRetVal Then
     ExecuteDSN = "" 'Execute ok
   Else
     ExecuteDSN = "Error: " & vbCrLf & _
                    "Invalid attributes" & vbCrLf
   End If
ExecuteDSNExit:
   Exit Function
ExecuteDSNError:
    ExecuteDSN = "Error: " & Err.Number & vbCrLf & _
               "Source: " & Err.Source & vbCrLf & _
               "Description: " & Err.Description
   Resume ExecuteDSNExit
End Function


Public Function ExecuteFileDSN(ByVal dsn_name As String, _
                            Optional ByVal sDriver As DSNDriver, _
                            Optional ByVal strAttr As String) As String
    Dim clave As String
    Dim valor As String
    Dim strDriver As String
    Dim arrayAttr() As String
    clave = "ODBC"
    subclave = "DRIVER"
    strDriver = GetDriver(sDriver)
    valor = "SQL Server"
    Call SQLWriteFileDSN(dsn_name, clave, "DRIVER", strDriver)
    arrayAttr = Split(strAttr, "|")
    For i = LBound(arrayAttr) To UBound(arrayAttr)
        s_left = Left(arrayAttr(i), InStr(arrayAttr(i), "=") - 1)
        s_right = Right(arrayAttr(i), Len(arrayAttr(i)) - InStr(arrayAttr(i), "="))
        Call SQLWriteFileDSN(dsn_name, clave, s_left, s_right)
    Next
End Function
