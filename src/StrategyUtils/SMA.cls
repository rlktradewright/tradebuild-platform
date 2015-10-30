VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMA"
Attribute VB_GlobalNameSpace = True
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

Private Const ModuleName                            As String = "SMA"

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

Public Function SMA( _
                Optional ByVal Ref As Long, _
                Optional ByVal SMAid As ResourceIdentifier) As Variant
Const ProcName As String = "SMA"
On Error GoTo Err

SMA = getSMA(SMAid).ma(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getSMA( _
                ByVal SMA As ResourceIdentifier) As CmnStudiesLib27.SMA
Const ProcName As String = "getBB"
On Error GoTo Err

If SMA Is Nothing Then Set SMA = gResourceContext.PrimarySMA
Assert Not SMA Is Nothing, "No SMA currently defined"

Dim lObj As Object
Set lObj = SMA.Resource
AssertArgument TypeOf lObj Is CmnStudiesLib27.SMA, "ResourceIdentifier does not refer to an SMA study"

Set getSMA = lObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




