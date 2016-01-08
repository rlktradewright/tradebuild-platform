Attribute VB_Name = "GConstTimeBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GConstTimeBars"

Public Const ConstTimeBarsInputOpenInterest As String = "Open interest"
Public Const ConstTimeBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const ConstTimeBarsInputPrice As String = "Price"
Public Const ConstTimeBarsInputPriceUcase As String = "PRICE"

Public Const ConstTimeBarsInputTotalVolume As String = "Total volume"
Public Const ConstTimeBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const ConstTimeBarsInputTickVolume As String = "Tick volume"
Public Const ConstTimeBarsInputTickVolumeUcase As String = "TICK VOLUME"

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
    mDefaultParameters.SetParameterValue ConstTimeBarsParamBarLength, 5
    mDefaultParameters.SetParameterValue ConstTimeBarsParamTimeUnits, _
                                        TimePeriodUnitsToString(TimePeriodMinute)
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

ReDim ar(6) As Variant

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = gCreateBarStudyDefinition( _
                                ConstTimeBarsStudyName, _
                                ConstTimeBarsStudyShortName, _
                                "Constant time bars " & _
                                "divide price movement into periods (bars) of equal time. " & _
                                "For each period the open, high, low and close price values " & _
                                "are determined.", _
                                ConstTimeBarsInputPrice, _
                                ConstTimeBarsInputTotalVolume, _
                                ConstTimeBarsInputTickVolume, _
                                ConstTimeBarsInputOpenInterest)
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamBarLength)
    paramDef.Description = "The number of time units in each constant time bar"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(ConstTimeBarsParamTimeUnits)
    paramDef.Description = "The time units that the constant time bars are measured in"
    paramDef.ParameterType = ParameterTypeString
    ar(0) = TimePeriodUnitsToString(TimePeriodSecond)
    ar(1) = TimePeriodUnitsToString(TimePeriodMinute)
    ar(2) = TimePeriodUnitsToString(TimePeriodHour)
    ar(3) = TimePeriodUnitsToString(TimePeriodDay)
    ar(4) = TimePeriodUnitsToString(TimePeriodWeek)
    ar(5) = TimePeriodUnitsToString(TimePeriodMonth)
    ar(6) = TimePeriodUnitsToString(TimePeriodYear)
    paramDef.PermittedValues = ar
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================












