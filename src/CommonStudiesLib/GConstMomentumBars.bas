Attribute VB_Name = "GConstMomentumBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstMomentumBars"

Public Const ConstMomentumBarsInputOpenInterest As String = "Open interest"
Public Const ConstMomentumBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstMomentumBarsInputPrice As String = "Price"
Public Const ConstMomentumBarsInputPriceUcase As String = "PRICE"

Public Const ConstMomentumBarsInputTotalVolume As String = "Total volume"
Public Const ConstMomentumBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstMomentumBarsInputTickVolume As String = "Tick volume"
Public Const ConstMomentumBarsInputTickVolumeUcase As String = "TICK VOLUME"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================


Private mDefaultParameters As Parameters
Private mStudyDefinition As StudyDefinition

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

'@================================================================================
' Procedures
'@================================================================================


Public Property Let defaultParameters(ByVal Value As Parameters)
' create a clone of the default parameters supplied by the caller
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set mDefaultParameters = Value.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = New Parameters
    mDefaultParameters.SetParameterValue ConstMomentumBarsParamTicksPerBar, 10
End If

' now create a clone of the default parameters for the caller
Set defaultParameters = mDefaultParameters.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyDefinition() As StudyDefinition
Dim inputDef As StudyInputDefinition
Dim valueDef As StudyValueDefinition
Dim paramDef As StudyParameterDefinition
Const ProcName As String = "StudyDefinition"
On Error GoTo Err

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = gCreateBarStudyDefinition( _
                                ConstMomentumBarsStudyName, _
                                ConstMomentumBarsStudyShortName, _
                                "Constant Momentum bars " & _
                                "divide price movement into periods (bars) of equal price movement. " & _
                                "For each period the open, high, low and close price values " & _
                                "are determined.", _
                                ConstMomentumBarsInputPrice, _
                                ConstMomentumBarsInputTotalVolume, _
                                ConstMomentumBarsInputTickVolume, _
                                ConstMomentumBarsInputOpenInterest)
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstMomentumBarsParamTicksPerBar)
    paramDef.Description = "The number of ticks movement in each constant momentum bar"
    paramDef.ParameterType = ParameterTypeInteger

End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================
















