VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITheme"
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

Private Const ModuleName                            As String = "ITheme"

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

Public Property Get BaseCOlor() As OLE_COLOR
End Property

Public Property Get BackColor() As OLE_COLOR
End Property

Public Property Get ForeColor() As OLE_COLOR
End Property

Public Property Get TextBackColor() As OLE_COLOR
End Property

Public Property Get TextForeColor() As OLE_COLOR
End Property

Public Property Get GridBackColorFixed() As OLE_COLOR
End Property

Public Property Get DecreasedValueColor() As OLE_COLOR
End Property

Public Property Get GridForeColor() As OLE_COLOR
End Property

Public Property Get GridForeColorFixed() As OLE_COLOR
End Property

Public Property Get GridGridColor() As OLE_COLOR
End Property

Public Property Get GridGridColorFixed() As OLE_COLOR
End Property

Public Property Get IncreasedValueColor() As OLE_COLOR
End Property

Public Property Get NegativeChangeBackColor() As OLE_COLOR
End Property

Public Property Get NegativeChangeForeColor() As OLE_COLOR
End Property

Public Property Get PositiveChangeBackColor() As OLE_COLOR
End Property

Public Property Get PositiveChangeForeColor() As OLE_COLOR
End Property

Public Property Get GridRowBackColorEven() As OLE_COLOR
End Property

Public Property Get GridRowBackColorOdd() As OLE_COLOR
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




