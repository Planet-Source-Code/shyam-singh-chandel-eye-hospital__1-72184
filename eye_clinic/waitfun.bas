Attribute VB_Name = "waitfun"
Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
    End Type
Private Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
    End Type
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessageA Lib "user32" (lpMsg As MSG) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private mCancel As Boolean
Private Sub TimerProc()
    mCancel = True
End Sub
Public Sub Wait(mSecs As Long)
    Dim MyMsg As MSG
    Dim TimerID As Long
    TimerID = SetTimer(0, 0, mSecs, AddressOf TimerProc)
    mCancel = False
    Do While Not mCancel
        GetMessage MyMsg, 0, 0, 0
        TranslateMessage MyMsg
        DispatchMessageA MyMsg
    Loop
    KillTimer 0, TimerID
End Sub

