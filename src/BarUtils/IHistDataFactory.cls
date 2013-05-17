VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IHistDataFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "IHistDataFactory"

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

Public Property Get TimePeriodValidator() As ITimePeriodValidator
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateBarWriter( _
                ByVal pContract As Contract, _
                ByVal pTimePeriod As TimePeriod) As TaskController
End Function

Public Function FetchBars( _
                ByVal pBarSpecifier As BarDataSpecifier, _
                ByVal pListener As IBarListener, _
                Optional ByVal pCookie As Variant, _
                Optional ByVal pNotifyProgress As Boolean = False, _
                Optional ByVal pReturnBarsObject As Boolean = True) As TaskController
End Function

Public Function SupportsWrite() As Boolean
End Function

Public Function SupportsAsyncWrite() As Boolean
End Function

'@================================================================================
' Helper Functions
'@================================================================================




