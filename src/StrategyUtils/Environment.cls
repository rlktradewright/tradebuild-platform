VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
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

Private Const ModuleName                            As String = "Environment"

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

Public Property Get ContractIdentifier() As String
Const ProcName As String = "ContractIdentifier"
On Error GoTo Err

ContractIdentifier = getContract.Specifier.ToString

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinutesToEndOfTradingSession() As Double
MinutesToEndOfTradingSession = (gTradingContext.Session.CurrentSessionEndTime - Timestamp) * 1440#
End Property

Public Property Get TickSize() As Double
TickSize = getContract.TickSize
End Property

Public Property Get TickValue() As Double
TickValue = getContract.TickValue
End Property

Public Property Get Timestamp() As Date
Timestamp = gTradingContext.Timestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ChartStudyValue( _
                ByVal Study As ResourceIdentifier, _
                ByVal ValueName As String, _
                Optional ByVal Timeframe As ResourceIdentifier)
Const ProcName As String = "ChartStudyValue"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"
Assert TypeOf Study.Resource Is IStudy

If Timeframe Is Nothing Then Set Timeframe = gResourceContext.PrimaryTimeframe

Assert TypeOf Timeframe.Resource Is Timeframe

gInitialisationContext.ChartStudyValue Study.Resource, ValueName, Timeframe.Resource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DefineParameter( _
                ByVal Name As String, _
                Optional ByVal DefaultValue As String)
Const ProcName As String = "DefineParameter"
On Error GoTo Err

Assert Not gStrategyRunner.DefaultParameters Is Nothing, "Can't define parameters at this point"
AssertArgument Name <> "", "Parameter name must be supplied"

gStrategyRunner.DefaultParameters.SetParameterValue Name, DefaultValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetParameterValue( _
                ByVal Name As String, _
                Optional ByVal DefaultValue As String) As String
Const ProcName As String = "GetParameterValue"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing Or gTradingContext Is Nothing, "Can't get parameters at this point"
AssertArgument Name <> "", "Parameter name must be supplied"

GetParameterValue = gStrategyRunner.Parameters.GetParameterValue(Name, DefaultValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LogTradeMessage(ByVal Message As String)
Const ProcName As String = "LogTradeMessage"
On Error GoTo Err

gTradingContext.LogTradeReason Message

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Function getContract() As IContract
If Not gInitialisationContext Is Nothing Then
    Set getContract = gInitialisationContext.Contract
ElseIf Not gTradingContext Is Nothing Then
    Set getContract = gTradingContext.Contract
Else
    Assert False, "No contract identifier available at this point"
End If
End Function


