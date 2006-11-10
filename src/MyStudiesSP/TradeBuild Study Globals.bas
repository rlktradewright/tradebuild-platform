Option Explicit


'================================================================================
' Constants
'================================================================================

'
' TODO define constants for the names of the study's input values, eg:
'
Public Const MyStudyInputPrice As String = "Price"

'
' TODO define constants for the names of the study's parameters, eg:
'
Public Const MyStudyParamPeriods As String = ParamPeriods

'
' TODO define constants for the names of the study's output values, eg:
'
Public Const MyStudyValueOut As String = "Out"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mDefaultParameters As IParameters
Private mStudyDefinition As IStudyDefinition

'================================================================================
' Procedures
'================================================================================

Public Property Let commonServiceConsumer( _
                ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property


Public Property Let defaultParameters(ByVal value As IParameters)
' create a clone of the default parameters supplied by the caller
Set mDefaultParameters = value.Clone
End Property

Public Property Get defaultParameters() As IParameters
If mDefaultParameters Is Nothing Then
    Set mDefaultParameters = mCommonServiceConsumer.NewParameters
'
' TODO set the default values for the parameters, eg
'
    mDefaultParameters.setParameterValue MyStudyParamPeriods, 14
End If

' now return a clone of the default parameters for the caller, to
' prevent the caller changing ours
Set defaultParameters = mDefaultParameters.Clone
End Property

Public Property Get studyDefinition() As TradeBuildSP.IStudyDefinition
Dim inputDef As IStudyInputDefinition
Dim valueDef As IStudyValueDefinition
Dim paramDef As IStudyParameterDefinition

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = mCommonServiceConsumer.NewStudyDefinition
'
' TODO set the name and short name to the relevant constants defined in Globals, eg
'
    mStudyDefinition.name = MyStudyName
    mStudyDefinition.shortName = MyStudyShortName
'
' TODO set the study description, eg
'
    mStudyDefinition.Description = "MyStudy identifies the precise instant where " & _
                                    "a major market move is about to begin"
'
' TODO set the chart region where the study values will be displayed by default (the
' user can override this):
'
'   DefaultRegionNone           means that the values will be displayed in the same
'                               region as the underlying data for the study (so for
'                               example if the user chooses to select volume the input,
'                               then the study will be displayed in the Volume region)
'
'   DefaultRegionCustom         means that the values will be displayed in their own
'                               chart region, regardless of where the input is coming
'                               from
'
'   DefaultRegionCustomPrice    means that the values will be displayed in the Price
'                               region
'
'   DefaultRegionCustomVolume   means that the values will be displayed in the Volume
'                               region
'
    mStudyDefinition.defaultRegion = StudyDefaultRegions.DefaultRegionCustom
    
'
' TODO setup the study input definitions, eg
'
' (NB: repeat this block as many times as there are inputs)
    Set inputDef = mStudyDefinition.StudyInputDefinitions.Add(MyStudyInputPrice)
    inputDef.name = MyStudyInputPrice   ' NB: this name will be passed to the study
                                        ' when an input is notified so that it can
                                        ' identify which input it has to deal with
    inputDef.inputType = InputTypeDouble
    inputDef.Description = "Input"
    
'
' TODO setup the study parameter definitions, eg
'
' (NB: repeat this block as many times as there are parameters, or delete it if
' the study has no parameters)
    Set paramDef = mStudyDefinition.StudyParameterDefinitions.Add(MyStudyParamPeriods)
    paramDef.name = MyStudyParamPeriods
    paramDef.Description = "The number of periods used to calculate the study value"
    paramDef.parameterType = ParameterTypeInteger

'
' TODO setup the study output value definitions, eg
'
' (NB: repeat this block as many times as there are output values)
    Set valueDef = mStudyDefinition.StudyValueDefinitions.Add(MyStudyValueOut)
    valueDef.name = MyStudyValueOut
    valueDef.Description = "The MyStudy value"
    valueDef.isDefault = True   ' only one value should be defined as default. If
                                ' a study only has one output, it should be defined
                                ' as default. If it has more than one, set the most
                                ' important or commonly used one as default, or don't
                                ' set a default at all
    valueDef.defaultRegion = DefaultRegionNone
        '
        ' DefaultRegionNone           means that the values will be displayed in the
        '                             same region as determined by
        '                             StudyDefinition.defaultRegion
        '
        ' DefaultRegionCustom         means that the values will be displayed in their own
        '                             chart region
        '
        ' DefaultRegionCustomPrice    means that the values will be displayed in the Price
        '                             region
        '
        ' DefaultRegionCustomVolume   means that the values will be displayed in the Volume
        '                             region
        '
    valueDef.maximumValue = 100 ' the maximum that this value can reach - if there is
                                ' no maximum, do not set this property at all
    valueDef.minimumValue = 0   ' the minimum that this value can reach - if there is
                                ' no minimum, do not set this property at all
    valueDef.valueType = ValueTypeDouble
    valueDef.multipleValuesPerBar = False
                                ' set this to true if your study needs to be able to
                                ' generate more than one value in a single bar - for
                                ' example if your study value consists of price
                                ' turning points, a single bar could be both a high
                                ' and a low
    
    
End If

' return a clone to prevent the application changing our definition
Set studyDefinition = mStudyDefinition.Clone
End Property

'================================================================================
' Helper Function
'================================================================================


