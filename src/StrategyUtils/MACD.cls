VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACD"
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

Private Const ModuleName                            As String = "MACD"

Private Const ValueMACD                             As String = "MACD"
Private Const ValueMACDHist                         As String = "MACD hist"
Private Const ValueMACDLowerBalance                 As String = "MACD lower balance"
Private Const ValueMACDSignal                       As String = "MACD signal"
Private Const ValueStrength                         As String = "Strength"
Private Const ValueStrengthCount                    As String = "Strength count"
Private Const ValueMACDUpperBalance                 As String = "MACD upper balance"

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

Public Property Get MACDValueMACD() As String
MACDValueMACD = ValueMACD
End Property

Public Property Get MACDValueMACDHist() As String
MACDValueMACDHist = ValueMACDHist
End Property

Public Property Get MACDValueMACDLowerBalance() As String
MACDValueMACDLowerBalance = ValueMACDLowerBalance
End Property

Public Property Get MACDValueMACDSignal() As String
MACDValueMACDSignal = ValueMACDSignal
End Property

Public Property Get MACDValueStrength() As String
MACDValueStrength = ValueStrength
End Property

Public Property Get MACDValueStrengthCount() As String
MACDValueStrengthCount = ValueStrengthCount
End Property

Public Property Get MACDValueMACDUpperBalance() As String
MACDValueMACDUpperBalance = ValueMACDUpperBalance
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function MACDHist( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDHist"
On Error GoTo Err

MACDHist = getMACD(MACD).MACDHist(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDHistDown( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDHistDown"
On Error GoTo Err

MACDHistDown = getMACD(MACD).MACDHistDown(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDHistUp( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDHistUp"
On Error GoTo Err

MACDHistUp = getMACD(MACD).MACDHistUp(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDNoStrength( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDNoStrength"
On Error GoTo Err

MACDNoStrength = (getMACD(MACD).strength(Ref).Value = CmnStudiesLib27.MACDNoStrength)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDSignal( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDSignal"
On Error GoTo Err

MACDSignal = getMACD(MACD).MACDSignal(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDStrengthConfirmed( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDStrengthConfirmed"
On Error GoTo Err

MACDStrengthConfirmed = (getMACD(MACD).strength(Ref).Value = CmnStudiesLib27.MACDConfirmedStrength)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDStrengthCount( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDStrengthCount"
On Error GoTo Err

MACDStrengthCount = getMACD(MACD).strengthCount(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDStrengthPotential( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDStrengthPotential"
On Error GoTo Err

MACDStrengthPotential = (getMACD(MACD).strength(Ref).Value = CmnStudiesLib27.MACDPotentialStrength)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDValue( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDValue"
On Error GoTo Err

MACDValue = getMACD(MACD).MACD(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDValueDown( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDValueDown"
On Error GoTo Err

MACDValueDown = getMACD(MACD).MACDDown(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDValueUp( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDValueUp"
On Error GoTo Err

MACDValueUp = getMACD(MACD).MACDUp(Ref).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDWeaknessConfirmed( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDWeaknessConfirmed"
On Error GoTo Err

MACDWeaknessConfirmed = (getMACD(MACD).strength(Ref).Value = CmnStudiesLib27.MACDConfirmedWeakness)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MACDWeaknessPotential( _
                Optional ByVal Ref As Long, _
                Optional ByVal MACD As ResourceIdentifier) As Variant
Const ProcName As String = "MACDWeaknessPotential"
On Error GoTo Err

MACDWeaknessPotential = (getMACD(MACD).strength(Ref).Value = CmnStudiesLib27.MACDPotentialWeakness)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getMACD( _
                ByVal MACD As ResourceIdentifier) As CmnStudiesLib27.MACD
Const ProcName As String = "getMACD"
On Error GoTo Err

If MACD Is Nothing Then Set MACD = gResourceContext.PrimaryMACD
Assert Not MACD Is Nothing, "No MACD is currently defined"

Dim lObj As Object
Set lObj = MACD.Resource
AssertArgument TypeOf lObj Is CmnStudiesLib27.MACD, "ResourceIdentifier does not refer to a MACD study"

Set getMACD = lObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




