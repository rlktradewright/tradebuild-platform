VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlackTheme"
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

Private Const ModuleName                            As String = "BlackTheme"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITheme Interface Members
'@================================================================================

Private Property Get ITheme_AlertFont() As stdole.Font

End Property

Private Property Get ITheme_AlertForeColor() As stdole.OLE_COLOR
ITheme_AlertForeColor = &HA8F0&
End Property

Private Property Get ITheme_Appearance() As TWUtilities40.AppearanceSettings
ITheme_Appearance = AppearanceFlat
End Property

Private Property Get ITheme_BackColor() As stdole.OLE_COLOR
ITheme_BackColor = &H404040
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
ITheme_BaseCOlor = &H0&
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
ITheme_ButtonBackColor = &H505050
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
ITheme_ButtonForeColor = &HE0E0E0
End Property

Private Property Get ITheme_ComboFont() As stdole.Font

End Property

Private Property Get ITheme_CoolbarBackColor() As stdole.OLE_COLOR
ITheme_CoolbarBackColor = &H404040
End Property

Private Property Get ITheme_DecreasedValueColor() As stdole.OLE_COLOR
ITheme_DecreasedValueColor = &H4444EB
End Property

Private Property Get ITheme_DefaultBorderColor() As stdole.OLE_COLOR
ITheme_DefaultBorderColor = &HF0FF00
End Property

Private Property Get ITheme_DisabledBackColor() As stdole.OLE_COLOR
ITheme_DisabledBackColor = &H454545
End Property

Private Property Get ITheme_DisabledFont() As stdole.Font

End Property

Private Property Get ITheme_DisabledForeColor() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_FocusBorderColor() As stdole.OLE_COLOR
ITheme_FocusBorderColor = &HFF00&
End Property

Private Property Get ITheme_ForeColor() As stdole.OLE_COLOR
ITheme_ForeColor = &HE0E0E0
End Property

Private Property Get ITheme_ForeColorAsk() As stdole.OLE_COLOR
ITheme_ForeColorAsk = &H202020
End Property

Private Property Get ITheme_ForeColorBid() As stdole.OLE_COLOR
ITheme_ForeColorBid = &H202020
End Property

Private Property Get ITheme_ForeColorTrade() As stdole.OLE_COLOR
ITheme_ForeColorTrade = &H202020
End Property

Private Property Get ITheme_GridBackColorFixed() As stdole.OLE_COLOR
ITheme_GridBackColorFixed = &H404040
End Property

Private Property Get ITheme_GridFont() As stdole.Font

End Property

Private Property Get ITheme_GridFontFixed() As stdole.Font

End Property

Private Property Get ITheme_GridForeColor() As stdole.OLE_COLOR
ITheme_GridForeColor = &HE0E0E0
End Property

Private Property Get ITheme_GridForeColorFixed() As stdole.OLE_COLOR
ITheme_GridForeColorFixed = ColorConstants.vbYellow
End Property

Private Property Get ITheme_GridLineColor() As stdole.OLE_COLOR
ITheme_GridLineColor = &H444444
End Property

Private Property Get ITheme_GridLineColorFixed() As stdole.OLE_COLOR
ITheme_GridLineColorFixed = &H505050
End Property

Private Property Get ITheme_GridRowBackColorEven() As stdole.OLE_COLOR
ITheme_GridRowBackColorEven = &H606060
End Property

Private Property Get ITheme_GridRowBackColorOdd() As stdole.OLE_COLOR
ITheme_GridRowBackColorOdd = &H808080
End Property

Private Property Get ITheme_GridRowForeColorEven() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_GridRowForeColorOdd() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_IncreasedValueColor() As stdole.OLE_COLOR
ITheme_IncreasedValueColor = &H72D813
End Property

Private Property Get ITheme_MouseoverBackColor() As stdole.OLE_COLOR
ITheme_MouseoverBackColor = &HD28400
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
ITheme_NonDefaultBorderColor = &H808080
End Property

Private Property Get ITheme_PositiveChangeBackColor() As stdole.OLE_COLOR
ITheme_PositiveChangeBackColor = &HB7E43
End Property

Private Property Get ITheme_PositiveChangeForeColor() As stdole.OLE_COLOR
ITheme_PositiveChangeForeColor = &HFFFFFF
End Property

Private Property Get ITheme_PushedBackColor() As stdole.OLE_COLOR
ITheme_PushedBackColor = &HFFA50F
End Property

Private Property Get ITheme_PushedFont() As stdole.Font

End Property

Private Property Get ITheme_PushedForeColor() As stdole.OLE_COLOR

End Property

Private Property Get ITheme_StatusBarBackColor() As stdole.OLE_COLOR
ITheme_StatusBarBackColor = &HD7BD97
End Property

Private Property Get ITheme_StatusBarForeColor() As stdole.OLE_COLOR
ITheme_StatusBarForeColor = &HC0C0C0
End Property

Private Property Get ITheme_TabstripBackColor() As stdole.OLE_COLOR
ITheme_TabstripBackColor = &H303030
End Property

Private Property Get ITheme_TabstripForeColor() As stdole.OLE_COLOR
ITheme_TabstripForeColor = &HE8A460
End Property

Private Property Get ITheme_TextBackColor() As stdole.OLE_COLOR
ITheme_TextBackColor = &H403010
End Property

Private Property Get ITheme_TextFont() As stdole.Font

End Property

Private Property Get ITheme_TextForeColor() As stdole.OLE_COLOR
ITheme_TextForeColor = &HF0F0F0
End Property

Private Property Get ITheme_ToolbarBackColor() As stdole.OLE_COLOR
ITheme_ToolbarBackColor = &H404040
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

