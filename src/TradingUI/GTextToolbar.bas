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

Public Type TWButtonInfo
    caption         As String
    Key             As String
    Style           As ButtonStyleConstants
    value           As ValueConstants
    ToolTipText     As String
    Enabled         As Boolean
End Type

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
                ByRef pButton As TWButtonInfo, _
                ByVal pPicture As PictureBox)
Const ProcName As String = "gAddButtonImageToImageList"
On Error GoTo Err

pImageList.ListImages.Add , IIf(pButton.Style <> tbrSeparator, pButton.Key, ""), getButtonImage(pButton.caption, pPicture)

Exit Sub

Err:
If Err.Number = 35602 Then Resume Next  'Key is not unique in collection
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gAddButtonToToolbar( _
                ByVal pToolbar As Toolbar, _
                ByRef pButton As TWButtonInfo) As Button
Const ProcName As String = "gAddButtonToToolbar"
On Error GoTo Err

If pButton.Style <> tbrSeparator Then
    Dim lKey As String
    lKey = GenerateGUIDString
    Set gAddButtonToToolbar = pToolbar.Buttons.Add(, lKey, , pButton.Style, IIf(pButton.Style <> tbrSeparator, pButton.Key, Empty))
    With pToolbar.Buttons.Item(lKey)
        .Enabled = pButton.Enabled
        .ToolTipText = pButton.ToolTipText
        .value = pButton.value
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
                ByRef pInfo As TWButtonInfo, _
                ByVal pCaption As String, _
                ByVal pKey As String, _
                ByVal pStyle As ButtonStyleConstants, _
                ByVal pValue As ValueConstants, _
                ByVal pTooltipText As String, _
                ByVal pEnabled As Boolean)
Const ProcName As String = "gCreateButtonInfo"
On Error GoTo Err

pInfo.caption = pCaption
pInfo.Enabled = pEnabled
pInfo.Key = pKey
pInfo.Style = pStyle
pInfo.ToolTipText = pTooltipText
pInfo.value = pValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetButtonImageInImageList(ByVal pImageList As ImageList, ByRef pButton As TWButtonInfo, ByVal pPicture As PictureBox)
Const ProcName As String = "gSetButtonImageInImageList"
On Error GoTo Err

pImageList.ListImages.Remove pButton.Key
pImageList.ListImages.Add , pButton.Key, getButtonImage(pButton.caption, pPicture)

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetupToolbar(ByVal pToolbar As Toolbar, ByVal pBackColor As Long, ByVal pForeColor As Long, ByVal pImageList As ImageList, ByRef pButtons() As TWButtonInfo, ByVal pPicture As PictureBox)
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
    gAdjustToolbarPictureSize pButtons(i).caption, pPicture
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





