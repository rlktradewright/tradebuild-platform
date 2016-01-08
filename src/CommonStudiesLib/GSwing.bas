Attribute VB_Name = "GSwing"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GSwing"

Public Const SwingInputValue As String = "Input"

Public Const SwingParamIncludeImplicitSwingPoints As String = "Include implicit swing points"
Public Const SwingParamMinimumSwingTicks As String = "Minimum swing (ticks)"

Public Const SwingValueSwingHighLine As String = "Swing high line"
Public Const SwingValueSwingLowLine As String = "Swing low line"
Public Const SwingValueSwingLine As String = "Swing line"
Public Const SwingValueSwingPoint As String = "Swing point"
Public Const SwingValueSwingHighPoint As String = "Swing high point"
Public Const SwingValueSwingLowPoint As String = "Swing low point"

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
    mDefaultParameters.SetParameterValue SwingParamMinimumSwingTicks, "10"
    mDefaultParameters.SetParameterValue SwingParamIncludeImplicitSwingPoints, "Yes"
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
    Set mStudyDefinition = New StudyDefinition
    mStudyDefinition.name = SwingName
    mStudyDefinition.ShortName = SwingShortName
    mStudyDefinition.Description = "Determines the significant swing points of " & _
                                    "the underlying. For a move to be considered a swing, " & _
                                    "it must move at least the distance specified in the " & _
                                    "Minimum swing (ticks) parameter."
    mStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionUnderlying
    
    
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(SwingInputValue)
    inputDef.InputType = InputTypeReal
    inputDef.Description = "Input value"
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingPoint)
    valueDef.Description = "Swing points"
    valueDef.IsDefault = True
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlack, DataPointDisplayModePoint, Layer:=LayerDataPoints + 60, Linethickness:=5, PointStyle:=PointSquare)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingHighPoint)
    valueDef.Description = "Swing high points"
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbBlue, DataPointDisplayModePoint, Layer:=LayerDataPoints + 60, Linethickness:=5, PointStyle:=PointSquare)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingLowPoint)
    valueDef.Description = "Swing low points"
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeNone
    valueDef.ValueStyle = gCreateDataPointStyle(vbRed, DataPointDisplayModePoint, Layer:=LayerDataPoints + 60, Linethickness:=5, PointStyle:=PointSquare)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingLine)
    valueDef.Description = "Swing point lines"
    valueDef.IncludeInChart = True
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeLine
    valueDef.ValueStyle = gCreateLineStyle(ArrowEndColor:=&H808080, ArrowEndFillColor:=vbYellow, ArrowEndStyle:=ArrowClosed, Color:=&H808080, Layer:=LayerLines)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingHighLine)
    valueDef.Description = "Swing high point lines"
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeLine
    valueDef.ValueStyle = gCreateLineStyle(ArrowEndColor:=vbBlue, ArrowEndFillColor:=vbBlue, ArrowEndStyle:=ArrowClosed, Color:=vbBlue, Layer:=LayerLines)
    valueDef.ValueType = ValueTypeReal
    
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(SwingValueSwingLowLine)
    valueDef.Description = "Swing low point lines"
    valueDef.IsDefault = False
    valueDef.DefaultRegion = StudyValueDefaultRegionDefault
    valueDef.ValueMode = ValueModeLine
    valueDef.ValueStyle = gCreateLineStyle(ArrowEndColor:=vbRed, ArrowEndFillColor:=vbRed, ArrowEndStyle:=ArrowClosed, Color:=vbRed, Layer:=LayerLines)
    valueDef.ValueType = ValueTypeReal
    
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamMinimumSwingTicks)
    paramDef.Description = "The minimum number of ticks bar clearance from a high/low to " & _
                            "establish a new swing"
    paramDef.ParameterType = ParameterTypeInteger

    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(SwingParamIncludeImplicitSwingPoints)
    paramDef.Description = "Indicates whether to include implied swing points"
    paramDef.ParameterType = ParameterTypeBoolean
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================





