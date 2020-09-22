Attribute VB_Name = "Folders"
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

Option Explicit
Private Const MAX_PATH = 260
Private Type BrowseInfo
  hWndOwner As Long
  'Handle of the user (ask with GetActiveWindow())
  pIDLRoot As Long
  'Adress of the IID-List.
  'It set the position of the first folder
  pszDisplayName As Long
  'Name of the selected folder.
  lpszTitle As Long
  'Displays the title of the dialog.
  ulFlags As Long
  'Flags, they show the effects of the dialog
  lpfnCallback As Long
  'Callback function
  lParam As Long
  'Gives the folder or a error back.
  iImage As Long
  'Displays the icon of a folder
End Type

'*********************************************************
' the following constants are the flags.

Private Const BIF_BROWSEFORCOMPUTER = &H1000
'Only computers are displayed.

Private Const BIF_BROWSEFORPRINTER = &H2000
'Only printers are displayed.

Private Const BIF_BROWSEINCLUDEFILES = &H4000
'The dialog will show files too.

Private Const BIF_DONTGOBELOWDOMAIN = &H2
'The dialog will not display networkfolders below a domain.

Private Const BIF_RETURNFSANCESTORS = &H8
'Only filesystemobjects are displayed.

Private Const BIF_RETURNONLYFSDIRS = &H1
'Only filesystemfolders are displayed.

Private Const BIF_STATUSTEXT = &H4
'The dialog will show a statusbar.

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
  (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias _
  "lstrcatA" (ByVal lpString1 As String, _
  ByVal lpString2 As String) As Long
'Get active window
Private Declare Function GetActiveWindow Lib "user32" () _
  As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
  (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
  (lpbi As BrowseInfo) As Long

Public Function BrowseForFolder(Prompt As String) As String
  
  Dim n As Integer
  Dim IDList As Long
  Dim result As Long
  Dim ThePath As String
  Dim BI As BrowseInfo

  'create filestructure
  With BI
    'Get handle of the active window
    .hWndOwner = GetActiveWindow()
    'Title of the dialog
    .lpszTitle = lstrcat(Prompt, "")
    'Only filesystemfolders are allowed.
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With

  'Show the dialog and give it to the IID-List
  IDList = SHBrowseForFolder(BI)

  'If IDList > 0, then edit the selected
  If IDList Then
    'Get memory
    ThePath = String$(MAX_PATH, 0)
    'convert IID-List to path
    result = SHGetPathFromIDList(IDList, ThePath)
    'delete memory for the IDList
    Call CoTaskMemFree(IDList)
    'delete all bytes behind Nullbyte
    n = InStr(ThePath, vbNullChar)
    If n Then ThePath = Left$(ThePath, n - 1)
  End If

  'Set callback
  BrowseForFolder = ThePath
End Function

Public Function CreateDirectory(ByVal psPath As String) As Integer

    Dim nCallDrive As String
    Dim pos1, pos2, pos3 As Integer
    Dim sDrive As String
    Dim sTemp As String

    On Error GoTo CreateDirErr
    
    sDrive = Left$(LTrim$(psPath), 3)
    nCallDrive = DriveExists(sDrive)
    If nCallDrive = "" Then
    sDrive = App.Path & "\"
    nCallDrive = DriveExists(sDrive)
       ' MsgBox "Drive Does not exist ", 16
       ' Exit Function
    End If

    If Right$(Trim$(psPath), 1) = "\" Then
      psPath = Left$(Trim$(psPath), Len(Trim$(psPath)) - 1)
    End If
    If Dir$(psPath, 16) = "" Then
        pos1 = 3
        ChDrive sDrive
        ChDir "\"
        Do
            pos2 = pos1
            pos1 = InStr(pos2 + 1, psPath, "\")
            If pos1 > 0 Then
                sTemp = Left$(psPath, pos1 - 1)
            Else
                sTemp = psPath
            End If
            MkDir sTemp
            ChDir sTemp
        Loop While pos1 > 0
    Else
        If InStr(Trim$(sDrive), ":") <> 0 Then
          ChDrive sDrive
        End If
        ChDir psPath
    End If
    psPath = LCase$(psPath)
    CreateDirectory = True
    Exit Function

CreateDirErr:
    If Err = 75 Or 76 Then Resume Next
    CreateDirectory = False
    Exit Function
    Resume  '   For debugging

End Function

Function DriveExists(ByVal drive As String) As String

    On Error Resume Next
    DriveExists = ""
    If Left$(drive, 2) <> "\\" Then
        DriveExists = Dir$(drive, 16)
    Else
        DriveExists = "\\"  '   Return anything other than ""
    End If

End Function

Public Sub SaveData(ByVal sFile, ByVal sFinalString)

On Error GoTo EditLab:

Dim nFile As Integer
nFile = FreeFile

Open sFile For Output As nFile
Print #nFile, sFinalString
Close #nFile

Exit Sub

EditLab:
    Close nFile
    MsgBox "Could not Save File.."
    MsgBox ERROR
    Exit Sub
End Sub



