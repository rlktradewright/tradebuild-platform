VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NativeTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITheme

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

Private Const ModuleName                            As String = "NativeTheme"

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

Private Property Get ITheme_AlertFont() As stdole.Font

End Property

Private Property Get ITheme_AlertForeColor() As stdole.OLE_COLOR
ITheme_AlertForeColor = &H80&
End Property

Private Property Get ITheme_Appearance() As TWUtilities40.AppearanceSettings
ITheme_Appearance = AppearanceFlat
End Property

Private Property Get ITheme_BackColor() As stdole.OLE_COLOR
ITheme_BackColor = SystemColorConstants.vbButtonFace
End Property

Private Property Get ITheme_BackColorAsk() As stdole.OLE_COLOR
ITheme_BackColorAsk = &HFAE968
End Property

Private Property Get ITheme_BackColorBid() As stdole.OLE_COLOR
ITheme_BackColorBid = &HD6C6F2
End Property

Private Property Get ITheme_BackColorTrade() As stdole.OLE_COLOR
ITheme_BackColorTrade = &HC1F7CA
End Property

Private Property Get ITheme_BaseCOlor() As stdole.OLE_COLOR
ITheme_BaseCOlor = SystemColorConstants.vbButtonFace
End Property

Private Property Get ITheme_BaseFont() As stdole.Font
Static sFont As StdFont
If sFont Is Nothing Then
    Set sFont = New StdFont
    sFont.Bold = False
    sFont.Italic = False
    sFont.Name = "Segoe UI"
    sFont.Size = 8
    sFont.Strikethrough = False
    sFont.Underline = False
End If
Set ITheme_BaseFont = sFont
End Property

Private Property Get ITheme_BorderStyle() As TWUtilities40.BorderStyleSettings
ITheme_BorderStyle = BorderStyleNone
End Property

Private Property Get ITheme_ButtonBackColor() As stdole.OLE_COLOR
ITheme_ButtonBackColor = SystemColorConstants.vbButtonFace
End Property

Private Property Get ITheme_ButtonFont() As stdole.Font
Static sFont As StdFont
If sFont Is Nothing Then
    Set sFont = New StdFont
    sFont.Bold = False
    sFont.Italic = False
    sFont.Name = "Segoe UI"
    sFont.Size = 8
    sFont.Strikethrough = False
    sFont.Underline = False
End If
Set ITheme_ButtonFont = sFont
End Property

Private Property Get ITheme_ButtonForeColor() As stdole.OLE_COLOR
ITheme_ButtonForeColor = SystemColorConstants.vbButtonText
End Property

Private Property Get ITheme_ComboFont() As stdole.Font

End Property

Private Property Get ITheme_CoolbarBackColor() As stdole.OLE_COLOR
ITheme_CoolbarBackColor = vbButtonFace
End Property

Private Property Get ITheme_DecreasedValueColor() As stdole.OLE_COLOR
ITheme_DecreasedValueColor = &H4444EB
End Property

Private Property Get ITheme_DefaultBorderColor() As stdole.OLE_COLOR
ITheme_DefaultBorderColor = &HFF7800
End Property

Private Property Get ITheme_DisabledBackColor() As stdole.OLE_COLOR
ITheme_DisabledBackColor = vbInactiveTitleBar
End Property

Private Property Get ITheme_DisabledFont() As stdole.Font

End Property

Private Property Get ITheme_DisabledForeColor() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_FocusBorderColor() As stdole.OLE_COLOR
ITheme_FocusBorderColor = &HA800&
End Property

Private Property Get ITheme_ForeColor() As stdole.OLE_COLOR
ITheme_ForeColor = SystemColorConstants.vbButtonText
End Property

Private Property Get ITheme_ForeColorAsk() As stdole.OLE_COLOR
ITheme_ForeColorAsk = vbBlack
End Property

Private Property Get ITheme_ForeColorBid() As stdole.OLE_COLOR
ITheme_ForeColorBid = vbBlack
End Property

Private Property Get ITheme_ForeColorTrade() As stdole.OLE_COLOR
ITheme_ForeColorTrade = vbBlack
End Property

Private Property Get ITheme_GridBackColorFixed() As stdole.OLE_COLOR
ITheme_GridBackColorFixed = &HF4F4F4
End Property

Private Property Get ITheme_GridFont() As stdole.Font
Static sFont As StdFont
If sFont Is Nothing Then
    Set sFont = New StdFont
    sFont.Bold = False
    sFont.Italic = False
    sFont.Name = "Segoe UI"
    sFont.Size = 10
    sFont.Strikethrough = False
    sFont.Underline = False
End If
Set ITheme_GridFont = sFont
End Property

Private Property Get ITheme_GridFontFixed() As stdole.Font

End Property

Private Property Get ITheme_GridForeColor() As stdole.OLE_COLOR
ITheme_GridForeColor = &H707070
End Property

Private Property Get ITheme_GridForeColorFixed() As stdole.OLE_COLOR
ITheme_GridForeColorFixed = &HA0A0A0
End Property

Private Property Get ITheme_GridLineColor() As stdole.OLE_COLOR
ITheme_GridLineColor = &H80000011
End Property

Private Property Get ITheme_GridLineColorFixed() As stdole.OLE_COLOR
ITheme_GridLineColorFixed = &HE0E0E0
End Property

Private Property Get ITheme_GridRowBackColorEven() As stdole.OLE_COLOR
ITheme_GridRowBackColorEven = &HEEEEEE
End Property

Private Property Get ITheme_GridRowBackColorOdd() As stdole.OLE_COLOR
ITheme_GridRowBackColorOdd = &HF8F8F8
End Property

Private Property Get ITheme_GridRowForeColorEven() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_GridRowForeColorOdd() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_IncreasedValueColor() As stdole.OLE_COLOR
ITheme_IncreasedValueColor = &HB7E43
End Property

Private Property Get ITheme_MouseoverBackColor() As stdole.OLE_COLOR
ITheme_MouseoverBackColor = &HFFF8E2
End Property

Private Property Get ITheme_MouseoverFont() As stdole.Font

End Property

Private Property Get ITheme_MouseOverForeColor() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_NegativeChangeBackColor() As stdole.OLE_COLOR
ITheme_NegativeChangeBackColor = &H4444EB
End Property

Private Property Get ITheme_NegativeChangeForeColor() As stdole.OLE_COLOR
ITheme_NegativeChangeForeColor = &HFFFFFF
End Property

Private Property Get ITheme_NonDefaultBorderColor() As stdole.OLE_COLOR
ITheme_NonDefaultBorderColor = vbButtonShadow
End Property

Private Property Get ITheme_PositiveChangeBackColor() As stdole.OLE_COLOR
ITheme_PositiveChangeBackColor = &HB7E43
End Property

Private Property Get ITheme_PositiveChangeForeColor() As stdole.OLE_COLOR
ITheme_PositiveChangeForeColor = &HFFFFFF
End Property

Private Property Get ITheme_PushedBackColor() As stdole.OLE_COLOR
ITheme_PushedBackColor = AdjustColorIntensity(ITheme_MouseoverBackColor, 0.95)
End Property

Private Property Get ITheme_PushedFont() As stdole.Font

End Property

Private Property Get ITheme_PushedForeColor() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_StatusBarBackColor() As stdole.OLE_COLOR
ITheme_StatusBarBackColor = vbButtonFace
End Property

Private Property Get ITheme_StatusBarForeColor() As stdole.OLE_COLOR
ITheme_StatusBarForeColor = vbButtonText
End Property

Private Property Get ITheme_TabstripBackColor() As stdole.OLE_COLOR
ITheme_TabstripBackColor = vbButtonFace
End Property

Private Property Get ITheme_TabstripForeColor() As stdole.OLE_COLOR
ITheme_TabstripForeColor = vbButtonText
End Property

Private Property Get ITheme_TextBackColor() As stdole.OLE_COLOR
ITheme_TextBackColor = SystemColorConstants.vbWindowBackground
End Property

Private Property Get ITheme_TextFont() As stdole.Font

End Property

Private Property Get ITheme_TextForeColor() As stdole.OLE_COLOR
ITheme_TextForeColor = SystemColorConstants.vbWindowText
End Property

Private Property Get ITheme_ToolbarBackColor() As stdole.OLE_COLOR
ITheme_ToolbarBackColor = vbButtonFace
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




