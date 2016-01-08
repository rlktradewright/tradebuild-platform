Attribute VB_Name = "GConstVolumeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstVolumeBars"

Public Const ConstVolBarsInputOpenInterest As String = "Open interest"
Public Const ConstVolBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstVolBarsInputPrice As String = "Price"
Public Const ConstVolBarsInputPriceUcase As String = "PRICE"

Public Const ConstVolBarsInputTotalVolume As String = "Total volume"
Public Const ConstVolBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstVolBarsInputTickVolume As String = "Tick volume"
Public Const ConstVolBarsInputTickVolumeUcase As String = "TICK VOLUME"

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
    mDefaultParameters.SetParameterValue ConstVolumeBarsParamVolPerBar, 1000
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
                                ConstVolumeBarsStudyName, _
                                ConstVolumeBarsStudyShortName, _
                                "Constant Momentum bars " & _
                                "Constant volume bars " & _
                                "divide price movement into periods (bars) of equal volume. " & _
                                "For each period the open, high, low and close price values " & _
                                "are determined.", _
                                ConstVolBarsInputPrice, _
                                ConstVolBarsInputTotalVolume, _
                                ConstVolBarsInputTickVolume, _
                                ConstVolBarsInputOpenInterest)
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstVolumeBarsParamVolPerBar)
    paramDef.Description = "The volume in each constant volume bar"
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














