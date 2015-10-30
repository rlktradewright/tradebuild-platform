VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Swing"
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

Private Const ModuleName                            As String = "Swing"

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

Public Function PotentialSwingHigh( _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "PotentialSwingHigh"
On Error GoTo Err

PotentialSwingHigh = getSwing(Swing).PotentialSwingHighPoint.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PotentialSwingHighBarNumber( _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "PotentialSwingHighBarNumber"
On Error GoTo Err

PotentialSwingHighBarNumber = getSwing(Swing).PotentialSwingHighPoint.BarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PotentialSwingHighTime( _
                Optional ByVal Swing As ResourceIdentifier) As Date
Const ProcName As String = "PotentialSwingHighTime"
On Error GoTo Err

PotentialSwingHighTime = getSwing(Swing).PotentialSwingHighPoint.Timestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PotentialSwingLow( _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "PotentialSwingLow"
On Error GoTo Err

PotentialSwingLow = getSwing(Swing).PotentialSwingLowPoint.Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PotentialSwingLowBarNumber( _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "PotentialSwingLowBarNumber"
On Error GoTo Err

PotentialSwingLowBarNumber = getSwing(Swing).PotentialSwingLowPoint.BarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PotentialSwingLowTime( _
                Optional ByVal Swing As ResourceIdentifier) As Date
Const ProcName As String = "PotentialSwingLowTime"
On Error GoTo Err

PotentialSwingLowTime = getSwing(Swing).PotentialSwingLowPoint.Timestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingHigh( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "SwingHigh"
On Error GoTo Err

SwingHigh = getSwing(Swing).SwingHighPoint(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingHighBarNumber( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "SwingHighBarNumber"
On Error GoTo Err

SwingHighBarNumber = getSwing(Swing).SwingHighPoint(Ref).BarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingHighTime( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Date
Const ProcName As String = "SwingHighTime"
On Error GoTo Err

SwingHighTime = getSwing(Swing).SwingHighPoint(Ref).Timestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingLow( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "SwingLow"
On Error GoTo Err

SwingLow = getSwing(Swing).SwingLowPoint(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingLowBarNumber( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Variant
Const ProcName As String = "SwingLowBarNumber"
On Error GoTo Err

SwingLowBarNumber = getSwing(Swing).SwingLowPoint(Ref).BarNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SwingLowTime( _
                Optional ByVal Ref As Long = 0, _
                Optional ByVal Swing As ResourceIdentifier) As Date
Const ProcName As String = "SwingLowTime"
On Error GoTo Err

SwingLowTime = getSwing(Swing).SwingLowPoint(Ref).Timestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getSwing( _
                ByVal Swing As ResourceIdentifier) As CmnStudiesLib27.Swing
Const ProcName As String = "getSwing"
On Error GoTo Err

If Swing Is Nothing Then Set Swing = gResourceContext.PrimarySwing
Assert Not Swing Is Nothing, "No Swing is currently defined"

Dim lObj As Object
Set lObj = Swing.Resource
AssertArgument TypeOf lObj Is CmnStudiesLib27.Swing, "ResourceIdentifier does not refer to a Swing study"

Set getSwing = lObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




