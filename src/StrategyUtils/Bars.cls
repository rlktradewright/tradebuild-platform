VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars"
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

Private Const ModuleName                            As String = "Bars"

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

Public Function BarNumber( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "CloseValue"
On Error GoTo Err

BarNumber = getBars(Timeframe).Bar(Ref).BarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CloseValue( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "CloseValue"
On Error GoTo Err

CloseValue = getBars(Timeframe).CloseValue(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CurrentBarNumber( _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "CurrentBarNumber"
On Error GoTo Err

CurrentBarNumber = getBars(Timeframe).CurrentBarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HigherHighValue( _
                ByVal Rank As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "HigherHighValue"
On Error GoTo Err

HigherHighValue = getBars(Timeframe).HigherHighValue(Rank, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HighestCloseValue( _
                ByVal NumberOfBars As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "HighestCloseValue"
On Error GoTo Err

HighestCloseValue = getBars(Timeframe).HighestCloseValue(NumberOfBars, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HighestHighValue( _
                ByVal NumberOfBars As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "HighestHighValue"
On Error GoTo Err

HighestHighValue = getBars(Timeframe).HighestHighValue(NumberOfBars, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HighValue( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "HighValue"
On Error GoTo Err

HighValue = getBars(Timeframe).HighValue(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsDownBar( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "IsDownBar"
On Error GoTo Err

IsDownBar = getBars(Timeframe).Down(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsInsideBar( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "IsInsideBar"
On Error GoTo Err

IsInsideBar = getBars(Timeframe).Inside(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsOutsideBar( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "IsOutsideBar"
On Error GoTo Err

IsOutsideBar = getBars(Timeframe).Outside(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsUpBar( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "IsUpBar"
On Error GoTo Err

IsUpBar = getBars(Timeframe).Up(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LowerLowValue( _
                ByVal Rank As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "LowerLowValue"
On Error GoTo Err

LowerLowValue = getBars(Timeframe).LowerLowValue(Rank, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LowestCloseValue( _
                ByVal NumberOfBars As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "LowestCloseValue"
On Error GoTo Err

LowestCloseValue = getBars(Timeframe).LowestCloseValue(NumberOfBars, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LowestLowValue( _
                ByVal NumberOfBars As Long, _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "LowestLowValue"
On Error GoTo Err

LowestLowValue = getBars(Timeframe).LowestLowValue(NumberOfBars, Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LowValue( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "LowValue"
On Error GoTo Err

LowValue = getBars(Timeframe).LowValue(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NumberOfBars( _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "NumberOfBars"
On Error GoTo Err

NumberOfBars = getBars(Timeframe).Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OpenValue( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "OpenValue"
On Error GoTo Err

OpenValue = getBars(Timeframe).OpenValue(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TickVolume( _
                Optional ByVal Ref As Long, _
                Optional ByVal Timeframe As ResourceIdentifier) As Variant
Const ProcName As String = "TickVolume"
On Error GoTo Err

TickVolume = getBars(Timeframe).TickVolume(Ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getBars( _
                Optional ByVal Timeframe As ResourceIdentifier) As BarUtils27.Bars
Const ProcName As String = "getBars"
On Error GoTo Err

If Timeframe Is Nothing Then Set Timeframe = gResourceContext.PrimaryTimeframe
Assert Not Timeframe Is Nothing, "No Timeframe is currently defined"

Dim lObj As Object
Set lObj = Timeframe.Resource
AssertArgument TypeOf lObj Is Timeframe, "ResourceIdentifier does not refer to a Timeframe"

Dim tf As Timeframe
Set tf = lObj

Set getBars = tf.BarStudy.BarsFuture.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


