Attribute VB_Name = "GUserDefinedBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GUserDefinedBars"

Public Const UserDefinedBarsInputValue  As String = "Value"
Public Const UserDefinedBarsInputValueUCase  As String = "VALUE"

Public Const UserDefinedBarsInputBarNumber  As String = "Bar number"
Public Const UserDefinedBarsInputBarNumberUCase  As String = "BAR NUMBER"

Public Const UserDefinedBarsValueBar As String = "Bar"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================


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
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Assert False, "Study has no parameters", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set defaultParameters = New Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyDefinition() As StudyDefinition
Const ProcName As String = "StudyDefinition"
On Error GoTo Err

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = UserDefinedBarsStudyName
    mStudyDefinition.NeedsBars = False
    mStudyDefinition.ShortName = UserDefinedBarsStudyShortName
    mStudyDefinition.Description = "User-defined bars " & _
                        "divide value movement into periods (bars) of duration " & _
                        "determined by the program that supplies the values. " & _
                        "For each period the open, high, low and close values " & _
                        "are determined."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom
    
    
    Dim inputDef As StudyInputDefinition
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(UserDefinedBarsInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Value"
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(UserDefinedBarsInputBarNumber)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Bar number"
    
    Dim valueDef As StudyValueDefinition
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(UserDefinedBarsValueBar)
    valueDef.Description = "The user-defined bars"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IncludeInChart = True
    valueDef.ValueMode = ValueModeBar
    valueDef.ValueStyle = gCreateBarStyle
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpen)
    valueDef.Description = "Bar open Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHigh)
    valueDef.Description = "Bar high Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueLow)
    valueDef.Description = "Bar low Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueClose)
    valueDef.Description = "Bar close Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.IsDefault = True
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&H80&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueTickVolume)
    valueDef.Description = "Bar tick volume"
    valueDef.DefaultRegion = StudyValueDefaultRegionCustom
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H800000, DisplayMode:=DataPointDisplayModeHistogram)
    valueDef.ValueType = ValueTypeInteger
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHL2)
    valueDef.Description = "Bar H+L/2 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHLC3)
    valueDef.Description = "Bar H+L+C/3 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOHLC4)
    valueDef.Description = "Bar O+H+L+C/4 Value"
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000)
    valueDef.ValueType = ValueTypeReal
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================














