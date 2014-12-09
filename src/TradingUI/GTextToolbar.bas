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

Private Type TWButtonInfo
    Caption         As String
    Key             As String
    Style           As ButtonStyleConstants
    Value           As ValueConstants
    TooltipText     As String
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

Public Sub gAddButtonImageToImageList(ByVal pImageList As ImageList, ByRef pButton As TWButtonInfo, ByVal pPicture As PictureBox)
pImageList.ListImages.Add , IIf(pButton.Style <> tbrSeparator, pButton.Key, ""), GetImage(pButton.Caption, pPicture)
End Sub

Public Sub gAddButtonToToolbar(ByVal pToolbar As Toolbar, ByRef pButton As TWButtonInfo)
If pButton.Style <> tbrSeparator Then
    pToolbar.Buttons.Add , pButton.Key, , pButton.Style, IIf(pButton.Style <> tbrSeparator, pButton.Key, Empty)
    With pToolbar.Buttons.Item(pButton.Key)
        .Enabled = pButton.Enabled
        .TooltipText = pButton.TooltipText
        .Value = pButton.Value
    End With
Else
    pToolbar.Buttons.Add , , , pButton.Style
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





