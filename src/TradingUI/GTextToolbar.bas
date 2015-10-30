Attribute VB_Name = "GTextToolbar"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GTextToolbar"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub gAddButtonImageToImageList( _
                ByVal pImageList As ImageList, _
                ByRef pButton As TWChartButtonInfo, _
                ByVal pPicture As PictureBox)
Const ProcName As String = "gAddButtonImageToImageList"
On Error GoTo Err

pImageList.ListImages.Add , IIf(pButton.Style <> tbrSeparator, pButton.Key, ""), getButtonImage(pButton.Caption, pPicture)

Exit Sub

Err:
If Err.Number = 35602 Then Resume Next  'Key is not unique in collection
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gAddButtonToToolbar( _
                ByVal pToolbar As Toolbar, _
                ByRef pButton As TWChartButtonInfo) As Button
Const ProcName As String = "gAddButtonToToolbar"
On Error GoTo Err

If pButton.Style <> tbrSeparator Then
    Set gAddButtonToToolbar = pToolbar.Buttons.Add(, pButton.Key, , pButton.Style, IIf(pButton.Style <> tbrSeparator, pButton.Key, Empty))
    With pToolbar.Buttons.Item(pButton.Key)
        .Enabled = pButton.Enabled
        .ToolTipText = pButton.ToolTipText
        .Value = pButton.Value
        .Tag = pButton
    End With
Else
    Set gAddButtonToToolbar = pToolbar.Buttons.Add(, , , pButton.Style)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gAdjustToolbarPictureSize(ByVal pText As String, ByVal pPicture As PictureBox)
Const ProcName As String = "gAdjustToolbarPictureSize"
On Error GoTo Err

pText = " " & pText & " "

pPicture.CurrentX = 0
pPicture.CurrentY = 0

Dim lHeight As Long
lHeight = pPicture.TextHeight(pText) + 2 * Screen.TwipsPerPixelY
If lHeight > pPicture.Height Then pPicture.Height = lHeight

Dim lWidth As Long
lWidth = pPicture.TextWidth(pText)
If lWidth > pPicture.Width Then pPicture.Width = lWidth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gCreateButtonInfo( _
                ByRef pInfo As TWChartButtonInfo, _
                ByVal pCaption As String, _
                ByVal pStyle As ButtonStyleConstants, _
                ByVal pValue As ValueConstants, _
                ByVal pTooltipText As String, _
                ByVal pEnabled As Boolean, _
                ByVal pChartIndex As Long)
Const ProcName As String = "gCreateButtonInfo"
On Error GoTo Err

pInfo.Caption = pCaption
pInfo.Enabled = pEnabled
pInfo.Key = GenerateGUIDString
pInfo.Style = pStyle
pInfo.ToolTipText = pTooltipText
pInfo.Value = pValue
pInfo.ChartIndex = pChartIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetButtonImageInImageList(ByVal pImageList As ImageList, ByRef pButton As TWChartButtonInfo, ByVal pPicture As PictureBox)
Const ProcName As String = "gSetButtonImageInImageList"
On Error GoTo Err

pImageList.ListImages.Remove pButton.Key
gAddButtonImageToImageList pImageList, pButton, pPicture

Exit Sub

Err:
If Err.Number = 35601 Then Resume Next  'Element not found
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetupToolbar(ByVal pToolbar As Toolbar, ByVal pBackColor As Long, ByVal pForeColor As Long, ByVal pImageList As ImageList, ByRef pButtons() As TWChartButtonInfo, ByVal pPicture As PictureBox)
Const ProcName As String = "gSetupToolbar"
On Error GoTo Err

pPicture.BackColor = pBackColor
pPicture.ForeColor = pForeColor
pPicture.Width = 0
pPicture.Height = 0

Set pToolbar.ImageList = Nothing
pToolbar.Buttons.Clear

pImageList.ListImages.Clear

Dim i As Long
For i = 0 To UBound(pButtons)
    gAdjustToolbarPictureSize pButtons(i).Caption, pPicture
Next

For i = 0 To UBound(pButtons)
    gAddButtonImageToImageList pImageList, pButtons(i), pPicture
    If i = 0 Then Set pToolbar.ImageList = pImageList   ' can't link image list to toolbar unless image list is non-empty
    gAddButtonToToolbar pToolbar, pButtons(i)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getButtonImage(ByVal pText As String, ByVal pPicture As PictureBox) As IPictureDisp
Const ProcName As String = "getButtonImage"
On Error GoTo Err

pText = " " & pText & " "

pPicture.Cls
If Not pPicture.AutoRedraw Then pPicture.AutoRedraw = True
pPicture.CurrentX = 0
pPicture.CurrentY = 0
pPicture.Print pText
Set getButtonImage = pPicture.Image

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





