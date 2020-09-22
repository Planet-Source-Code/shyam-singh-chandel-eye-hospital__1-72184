VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'////                                                                    /////
'////     Developer: Shyam Singh Chandel                                 /////
'////     Jr. Technician (United News of India, Shillong)                /////
'////     URL http://groups.yahoo.com/group/sscsoft/                     /////
'////                                                                    /////
'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public HasPicture As Boolean: Public PicHeight As Long: Public PicWidth As Long
Public RowFrom As Long: Public RowTo As Long: Public RowHeight As Long: Public LastRowPrinted As Boolean
Public NumOfRowsPerPage As Integer: Dim lngLastRow As Long
Public DrawHorizontalLines As Boolean: Public DrawVerticalLines As Boolean: Dim bHeadPrinted As Boolean
Public BorderDistance As Long: Public DrawBorder As Boolean
Public PosX As Long: Public PosY As Long:
Public CurrentX As Long
Public CurrentY As Long
Dim objListView As Object:
Dim lngWidthHor As Long, lngWidthVer As Long, lngWidthBor As Long
Dim lngColorHor As Long, lngColorVer As Long, lngColorBor As Long

Private Const DT_END_ELLIPSIS = 32768: Const DT_MODIFYSTRING = 65536

Enum LineID
        VerticalLines = 1
        HorizontalLines = 2
        BorderLines = 3
End Enum
Dim M As sPrint

Private Sub Class_Initialize()
On Error Resume Next

PicHeight = 16: PicWidth = 16
BorderDistance = 5
DrawHorizontalLines = True: DrawVerticalLines = True: DrawBorder = True
lngWidthBor = 1: lngWidthHor = 1: lngWidthVer = 1
lngColorHor = vbBlack: lngColorVer = vbBlack: lngColorBor = vbBlack
NumOfRowsPerPage = 10
lngLastRow = 1
RowHeight = 250

End Sub

'  Prints the Head(Column header of ListView Control)

Public Sub PrintHead(Obj As Object)
On Error Resume Next

Dim lngKount As Long, rectBox As RECT, strTmp As String, lngBUpScalemode As Long
Dim intAlign As Integer, intIconNum As Integer, intSpace_Btn_Txt_n_Pic As Integer
Dim lngPosY As Long, lngPosX As Long, DX As Long, DY As Long
Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
Dim sngXFac As Single, sngYFac As Single, intPad As Integer

sngXFac = 1: sngYFac = 1
intPad = 3
Obj.ForeColor = objListView.ForeColor
lngNumOfColumns = objListView.ColumnHeaders.Count
lngNumOfSubItems = lngNumOfColumns - 1
lngRowHeight = Obj.ScaleY(RowHeight, vbTwips, vbPixels)
If TypeOf Obj Is Printer Then
            Obj.Print
            sngXFac = Screen.TwipsPerPixelX / Obj.TwipsPerPixelX
            sngYFac = Screen.TwipsPerPixelY / Obj.TwipsPerPixelY
            intPad = 3 * sngXFac
End If  'TypeOf Obj Is
lngBUpScalemode = Obj.ScaleMode
Obj.ScaleMode = vbPixels
intSpace_Btn_Txt_n_Pic = Obj.ScaleX(400, vbTwips, vbPixels)

lngPosX = Obj.ScaleX(PosX, vbTwips, vbPixels)
lngPosY = Obj.ScaleY(PosY, vbTwips, vbPixels)

If DrawBorder Then
lngPosX = lngPosX + (BorderDistance * sngXFac)
lngPosY = lngPosY + (BorderDistance * sngYFac)
End If

With rectBox
           .Left = lngPosX
           .Top = lngPosY
           
          .Bottom = lngRowHeight + .Top
          
            For lngKount = 1 To lngNumOfColumns
                        .Right = (objListView.ColumnHeaders(lngKount).Width * sngXFac) + .Left
                        strTmp = objListView.ColumnHeaders(lngKount).Text
                        intAlign = objListView.ColumnHeaders(lngKount).Alignment
                        If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
                        .Left = .Left + intPad: .Right = .Right - intPad
                       .Top = .Top + (2 * sngYFac)
                  
                                    If HasPicture = True And intAlign = 0 Then
                                                    intIconNum = objListView.ColumnHeaders(lngKount).Icon
                                                                    If intIconNum <> 0 And Sgn(.Right - .Left) <> -1 Then
                                                                                 DX = (.Right - .Left) + (4 * sngXFac): DY = PicWidth
                                                                                 If DX < DY Then DY = DX
                                                                                 Obj.PaintPicture objListView.ColumnHeaderIcons.ListImages(intIconNum).Picture, .Left, .Top, (DY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                                                                                .Left = .Left + intSpace_Btn_Txt_n_Pic
                                                                    End If 'intIconNum
                                    End If  'HasPicture
                        DrawText Obj.hDC, strTmp, Len(strTmp), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                       .Left = .Left - intPad: .Right = .Right + intPad
                       .Top = .Top - (2 * sngYFac)
                       .Left = .Right
            Next  'lngKount = 1 To lngNumOfColumns

            If DrawHorizontalLines Then
                        Obj.DrawWidth = lngWidthHor
                        Obj.Line (lngPosX, .Top)-(.Left, .Bottom - lngRowHeight), lngColorHor 'first line
                        Obj.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor 'second line
            End If  'DrawHorizontalLines
            If DrawBorder Then .Right = .Right + BorderDistance * sngXFac: .Bottom = .Bottom + BorderDistance * sngYFac

           If TypeOf Obj Is Printer Then
                        CurrentX = .Left * Obj.TwipsPerPixelX
                        CurrentY = .Top * Obj.TwipsPerPixelY
            Else
                        CurrentX = .Right * Screen.TwipsPerPixelX
                        CurrentY = .Bottom * Screen.TwipsPerPixelY
            End If

End With  'RectBox
bHeadPrinted = True
Obj.ScaleMode = lngBUpScalemode
End Sub

Public Sub PrintBody(Obj As Object)
On Error Resume Next

Dim intAlign As Integer, lngCol As Long
Dim lngRow As Long, rectBox As RECT, lngSubItems As Long, strText As String, lngBottom As Long
Dim lngColWidth As Long, lngBUpScalemode As Long, intIconNum As Integer, intSpace_Btn_Txt_n_Pic As Integer
Dim DX As Long, DY As Long, lngMaxRow As Long
Dim lngPosY As Long, lngPosX As Long
Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
Dim sngXFac As Single, sngYFac As Single, intPad As Integer


Obj.ForeColor = objListView.ForeColor
intPad = 3: sngXFac = 1: sngYFac = 1
lngNumOfColumns = objListView.ColumnHeaders.Count
lngNumOfSubItems = lngNumOfColumns - 1
lngRowHeight = Obj.ScaleY(RowHeight, vbTwips, vbPixels)

If TypeOf Obj Is Printer Then
            Obj.Print
            sngXFac = Screen.TwipsPerPixelX / Obj.TwipsPerPixelX
            sngYFac = Screen.TwipsPerPixelY / Obj.TwipsPerPixelY
            intPad = 3 * sngXFac
End If

lngBUpScalemode = Obj.ScaleMode
Obj.ScaleMode = vbPixels
intSpace_Btn_Txt_n_Pic = Obj.ScaleX(260, vbTwips, vbPixels)
Obj.DrawWidth = lngWidthHor
lngMaxRow = objListView.ListItems.Count
LastRowPrinted = False

lngPosX = Obj.ScaleX(PosX, vbTwips, vbPixels)
lngPosY = Obj.ScaleY(PosY, vbTwips, vbPixels)
       
 If DrawBorder Then
                lngPosX = lngPosX + (BorderDistance * sngXFac)
                lngPosY = lngPosY + (BorderDistance * sngYFac)
 End If

With rectBox
           
           .Left = lngPosX: .Top = lngPosY
            
            If bHeadPrinted Then .Top = lngPosY + lngRowHeight
                       
            For lngRow = RowFrom To RowTo
                        For lngCol = 1 To lngNumOfColumns 'objListView.ListItems(lngRow).ListSubItems.Count + 1
                                If lngCol = 1 Then
                                        strText = objListView.ListItems(lngRow).Text
                                        Obj.FontBold = objListView.ListItems(lngRow).Bold
                                        Obj.ForeColor = objListView.ListItems(lngRow).ForeColor
                                        
                                Else
                                        strText = objListView.ListItems(lngRow).SubItems(lngCol - 1)
                                        Obj.FontBold = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).Bold
                                        Obj.ForeColor = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).ForeColor
                                End If  'lngCol = 1
                                intAlign = objListView.ColumnHeaders(lngCol).Alignment
                                If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
                                .Right = .Left + (objListView.ColumnHeaders(lngCol).Width * sngXFac)
                                .Bottom = .Top + lngRowHeight
                                .Left = .Left + intPad: .Right = .Right - intPad
                                .Top = .Top + (2 * sngYFac)
                                
                                If HasPicture Then
                                                                           
                                            If lngCol = 1 Then intIconNum = objListView.ListItems(lngRow).SmallIcon Else intIconNum = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).ReportIcon
                                                    If intIconNum <> 0 Then
                                                                  DX = (.Right - .Left) + (4 * sngYFac): DY = PicWidth
                                                                  If DX < DY Then DY = DX
                                                                  Obj.PaintPicture objListView.SmallIcons.ListImages(intIconNum).Picture, .Left, .Top, (DY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                                                                 .Left = .Left + intSpace_Btn_Txt_n_Pic
                                                    End If  ' intIconNum <> 0
                                                    
                                End If  'HasPicture
                                
                                DrawText Obj.hDC, strText, Len(strText), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                                Obj.FontBold = False: Obj.ForeColor = objListView.ForeColor
                               .Left = .Left - intPad: .Right = .Right + intPad
                               .Top = .Top - (2 * sngYFac)
                                
                               .Left = .Right
                    
                        Next  'lngCol = 1 To lngNumOfColumns
                        If DrawHorizontalLines Then Obj.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor
                       
                       .Left = lngPosX: lngBottom = .Bottom
                       .Top = .Bottom
                        If lngRow = lngMaxRow Then LastRowPrinted = True: Exit For
            Next  'lngRow = RowFrom To RowTo
            
            If DrawHorizontalLines And (Not bHeadPrinted) Then Obj.Line (lngPosX, lngPosY)-(.Right, lngPosY), lngColorHor
            
            If DrawVerticalLines Then
                        Obj.DrawWidth = lngWidthVer
                        Obj.Line (lngPosX, lngPosY)-(lngPosX, lngBottom), lngColorVer
                        For lngCol = 1 To lngNumOfColumns
                                lngColWidth = lngColWidth + (objListView.ColumnHeaders(lngCol).Width * sngXFac)
                                Obj.Line (lngPosX + lngColWidth, lngPosY)-(lngPosX + lngColWidth, lngBottom), lngColorVer
                        Next  'lngCol = 1 To lngNumOfColumns
            End If  'DrawVerticalLines

            If DrawBorder Then
                       Obj.DrawWidth = lngWidthBor
                        lngPosY = lngPosY - (BorderDistance * sngYFac)
                        lngPosX = lngPosX - (BorderDistance * sngXFac)
                        DX = .Right + (BorderDistance * sngXFac)
                        DY = .Bottom + (BorderDistance * sngYFac)
                        
                        Obj.Line (DX, DY)-(DX, lngPosY), lngColorBor
                        Obj.Line -(lngPosX, lngPosY), lngColorBor
                        Obj.Line -(lngPosX, DY), lngColorBor
                        Obj.Line -(DX, DY), lngColorBor
                        .Right = .Right + (BorderDistance * sngXFac): .Bottom = .Bottom + (BorderDistance * sngYFac)
            End If  'DrawBorder
            If TypeOf Obj Is Printer Then
                    CurrentX = .Right * Obj.TwipsPerPixelX
                    CurrentY = .Bottom * Obj.TwipsPerPixelY
            Else
                    CurrentX = .Right * Screen.TwipsPerPixelX
                    CurrentY = .Bottom * Screen.TwipsPerPixelY
            End If
End With  'rectbox
bHeadPrinted = False
Obj.ScaleMode = lngBUpScalemode
End Sub

'Sets the ListViewControl
Public Property Set ListViewName(ByVal oNewValue As Object)
Set objListView = oNewValue
End Property

'LineWidth and Color setting, Printer can have non integer value such as 1.2, 1.5, 2.1 etc.
Public Sub SetLine(Obj As Object, ID As LineID, Optional Width As Single = 1#, Optional color As Long = vbBlack)
On Error Resume Next

Dim intWidth As Integer
Dim sngMul As Single
If Width = 1 Then intWidth = 1: GoTo DownThere
If TypeOf Obj Is Printer Then
            sngMul = 15 / Printer.TwipsPerPixelX  'Trying to make drawwidth similar in different resolution printer
            intWidth = Abs(sngMul * Width)
Else
            intWidth = Abs(Int(Width))
            If intWidth = 0 Then intWidth = 1
End If
DownThere:
Select Case ID
            Case HorizontalLines
                       lngWidthHor = intWidth
                       lngColorHor = color
            Case VerticalLines
                       lngWidthVer = intWidth
                       lngColorVer = color
            Case BorderLines
                       lngWidthBor = intWidth
                       lngColorBor = color
End Select
End Sub

'Set RowTo and RowFrom, to make printing several page easier.
Public Sub SetRows()

RowFrom = lngLastRow
RowTo = lngLastRow + (NumOfRowsPerPage - 1)
lngLastRow = RowTo + 1
If lngLastRow > objListView.ListItems.Count Then lngLastRow = 1
End Sub



