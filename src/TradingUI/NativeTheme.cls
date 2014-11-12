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
Attribute VB_Exposed = False
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

Private Property Get ITheme_BackColor() As stdole.OLE_COLOR
ITheme_BackColor = SystemColorConstants.vbButtonFace
End Property

Private Property Get ITheme_BaseCOlor() As stdole.OLE_COLOR
ITheme_BaseCOlor = SystemColorConstants.vbButtonFace
End Property

Private Property Get ITheme_ForeColor() As stdole.OLE_COLOR
ITheme_ForeColor = SystemColorConstants.vbButtonText
End Property

Private Property Get ITheme_TextBackColor() As stdole.OLE_COLOR
ITheme_TextBackColor = SystemColorConstants.vbWindowBackground
End Property

Private Property Get ITheme_TextForeColor() As stdole.OLE_COLOR
ITheme_TextForeColor = SystemColorConstants.vbWindowText
End Property

Private Property Get ITheme_GridBackColorFixed() As stdole.OLE_COLOR
ITheme_GridBackColorFixed = &HF4F4F4
End Property

Private Property Get ITheme_DecreasedValueColor() As stdole.OLE_COLOR
ITheme_DecreasedValueColor = &H4444EB
End Property

Private Property Get ITheme_GridForeColor() As stdole.OLE_COLOR
ITheme_GridForeColor = &H707070
End Property

Private Property Get ITheme_GridForeColorFixed() As stdole.OLE_COLOR
ITheme_GridForeColorFixed = &HA0A0A0
End Property

Private Property Get ITheme_GridGridColor() As stdole.OLE_COLOR
ITheme_GridGridColor = &H80000011
End Property

Private Property Get ITheme_GridGridColorFixed() As stdole.OLE_COLOR
ITheme_GridGridColorFixed = &HE0E0E0
End Property

Private Property Get ITheme_IncreasedValueColor() As stdole.OLE_COLOR
ITheme_IncreasedValueColor = &HB7E43
End Property

Private Property Get ITheme_NegativeChangeBackColor() As stdole.OLE_COLOR
ITheme_NegativeChangeBackColor = &H4444EB
End Property

Private Property Get ITheme_NegativeChangeForeColor() As stdole.OLE_COLOR
ITheme_NegativeChangeForeColor = &HFFFFFF
End Property

Private Property Get ITheme_PositiveChangeBackColor() As stdole.OLE_COLOR
ITheme_PositiveChangeBackColor = &HB7E43
End Property

Private Property Get ITheme_PositiveChangeForeColor() As stdole.OLE_COLOR
ITheme_PositiveChangeForeColor = &HFFFFFF
End Property

Private Property Get ITheme_GridRowBackColorEven() As stdole.OLE_COLOR
ITheme_GridRowBackColorEven = &HEEEEEE
End Property

Private Property Get ITheme_GridRowBackColorOdd() As stdole.OLE_COLOR
ITheme_GridRowBackColorOdd = &HF8F8F8
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




