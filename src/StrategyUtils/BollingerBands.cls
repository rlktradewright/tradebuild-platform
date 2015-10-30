VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BollingerBands"
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

Private Const ModuleName                            As String = "BollingerBands"

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

Public Function BollingerBottom( _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerBottom"
On Error GoTo Err

BollingerBottom = getBB(BollingerBands).Bottom(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerCentre( _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerCentre"
On Error GoTo Err

BollingerCentre = getBB(BollingerBands).Centre(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerInBottomBand( _
                ByVal Value As Double, _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerInBottomBand"
On Error GoTo Err

BollingerInBottomBand = getBB(BollingerBands).InBottomBand(Value, Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerInCentreBand( _
                ByVal Value As Double, _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerInCentreBand"
On Error GoTo Err

BollingerInCentreBand = getBB(BollingerBands).InCentreBand(Value, Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerInTopBand( _
                ByVal Value As Double, _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerInTopBand"
On Error GoTo Err

BollingerInTopBand = getBB(BollingerBands).InTopBand(Value, Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerSpread( _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerSpread"
On Error GoTo Err

BollingerSpread = getBB(BollingerBands).Spread(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BollingerTop( _
                Optional ByVal Ref As Long, _
                Optional ByVal BollingerBands As ResourceIdentifier) As Variant
Const ProcName As String = "BollingerTop"
On Error GoTo Err

BollingerTop = getBB(BollingerBands).Top(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBB( _
                ByVal BB As ResourceIdentifier) As CmnStudiesLib27.BollingerBands
Const ProcName As String = "getBB"
On Error GoTo Err

If BB Is Nothing Then Set BB = gResourceContext.PrimaryBollingerBands
Assert Not BB Is Nothing, "No BollingerBands currently defined"

Dim lObj As Object
Set lObj = BB.Resource
AssertArgument TypeOf lObj Is CmnStudiesLib27.BollingerBands, "ResourceIdentifier does not refer to a BollingerBands study"

Set getBB = lObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




