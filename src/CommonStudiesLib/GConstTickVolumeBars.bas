Attribute VB_Name = "GConstTickVolumeBars"
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

Private Const ModuleName                            As String = "GConstTickVolumeBars"

Public Const ConstTickVolumeBarsInputOpenInterest As String = "Open interest"
Public Const ConstTickVolumeBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstTickVolumeBarsInputPrice As String = "Price"
Public Const ConstTickVolumeBarsInputPriceUcase As String = "PRICE"

Public Const ConstTickVolumeBarsInputTotalVolume As String = "Total volume"
Public Const ConstTickVolumeBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstTickVolumeBarsInputTickVolume As String = "Tick volume"
Public Const ConstTickVolumeBarsInputTickVolumeUcase As String = "TICK VOLUME"

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
    mDefaultParameters.SetParameterValue ConstTickVolumeBarsParamTicksPerBar, 100
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
                                ConstTickVolumeBarsStudyName, _
                                ConstTickVolumeBarsStudyShortName, _
                                "Constant Tick Volume bars " & _
                                "divide price movement into periods (bars) with equal numbers of trades. " & _
                                "For each period the open, high, low and close price values " & _
                                "are determined.", _
                                ConstTickVolumeBarsInputPrice, _
                                ConstTickVolumeBarsInputTotalVolume, _
                                ConstTickVolumeBarsInputTickVolume, _
                                ConstTickVolumeBarsInputOpenInterest)
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTickVolumeBarsParamTicksPerBar)
    paramDef.Description = "The number of trades in each constant tick volume bar"
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


















