VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyUtils"

'@================================================================================
' Enums
'@================================================================================

Public Enum ReplayNumbers
    ReplayAll = -1
    ReplayAsRequired = -2
End Enum

Public Enum StudyDefaultRegions
    StudyDefaultRegionNone = 0
    StudyDefaultRegionCustom
    StudyDefaultRegionUnderlying
End Enum

Public Enum StudyValueDefaultRegions
    StudyValueDefaultRegionNone = 0
    StudyValueDefaultRegionCustom
    StudyValueDefaultRegionUnderlying
    StudyValueDefaultRegionDefault
End Enum

Public Enum StudyInputTypes
    InputTypeNone = 0
    InputTypeInteger = 1
    InputTypeReal = 2
    InputTypeString = 3
    InputTypeDate = 4
    InputTypeBoolean = 5
End Enum

Public Enum StudyParameterTypes
    ParameterTypeNone = 0
    ParameterTypeInteger = 1
    ParameterTypeReal = 2
    ParameterTypeString = 3
    ParameterTypeDate = 4
    ParameterTypeBoolean = 5
End Enum

Public Enum StudyValueModes
    ValueModeNone = 0
    ValueModeLine = 1
    ValueModeBar = 2
    ValueModeText = 3
End Enum

Public Enum StudyValueTypes
    ValueTypeNone = 0
    ValueTypeInteger = 1
    ValueTypeReal = 2
    ValueTypeString = 3
    ValueTypeDate = 4
    ValueTypeBoolean = 5
End Enum

Public Enum TaskDiscriminators
    TaskAddStudy
    TaskAddStudyValueListener
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type StudyValueEventData
    Source              As Object
    ValueName           As String
    sVal                As SValue
End Type

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

Public Property Get BuiltInStudyLibraryName() As String
BuiltInStudyLibraryName = BuiltInStudyLibName
End Property

Public Property Get BuiltInStudyLibraryProgId() As String
BuiltInStudyLibraryProgId = BuiltInStudyLibProgId
End Property

Public Property Get ConstTickVolumeBarsStudyName() As String
ConstTickVolumeBarsStudyName = Globals.ConstTickVolumeBarsStudyName
End Property

Public Property Get ConstTimeBarsStudyName() As String
ConstTimeBarsStudyName = Globals.ConstTimeBarsStudyName
End Property

Public Property Get ConstTimeBarsParamBarLength() As String
ConstTimeBarsParamBarLength = Globals.ConstTimeBarsParamBarLength
End Property

Public Property Get ConstTimeBarsParamTimeUnits() As String
ConstTimeBarsParamTimeUnits = Globals.ConstTimeBarsParamTimeUnits
End Property

Public Property Get ConstVolumeBarsStudyName() As String
ConstVolumeBarsStudyName = Globals.ConstVolumeBarsStudyName
End Property

Public Property Get ConstVolumeBarsParamVolPerBar() As String
ConstVolumeBarsParamVolPerBar = Globals.ConstVolumeBarsParamVolPerBar
End Property

Public Property Get ConstMomentumBarsStudyName() As String
ConstMomentumBarsStudyName = Globals.ConstMomentumBarsStudyName
End Property

Public Property Get ConstMomentumBarsParamTicksPerBar() As String
ConstMomentumBarsParamTicksPerBar = Globals.ConstMomentumBarsParamTicksPerBar
End Property

Public Property Get InputNameAsk() As String
InputNameAsk = Globals.AskInputName
End Property

Public Property Get InputNameBid() As String
InputNameBid = Globals.BidInputName
End Property

Public Property Get InputNameOpenInterest() As String
InputNameOpenInterest = Globals.OpenInterestInputName
End Property

Public Property Get InputNameTickVolume() As String
InputNameTickVolume = Globals.TickVolumeInputName
End Property

Public Property Get InputNameTrade() As String
InputNameTrade = Globals.TradeInputName
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = Globals.VolumeInputName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateBarStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "CreateBarStudy"
On Error GoTo Err

Select Case pTimePeriod.Units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    Set CreateBarStudy = setupConstantTimeBarsStudy(pTimePeriod, pStudyBase, pInitialBarsFuture)
Case TimePeriodTickMovement
    Set CreateBarStudy = setupConstantMomentumBarsStudy(pTimePeriod, pStudyBase, pInitialBarsFuture)
Case TimePeriodTickVolume
    Set CreateBarStudy = setupConstantTickVolumeBarsStudy(pTimePeriod, pStudyBase, pInitialBarsFuture)
Case TimePeriodVolume
    Set CreateBarStudy = setupConstantVolumeBarsStudy(pTimePeriod, pStudyBase, pInitialBarsFuture)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateCacheReplayTask( _
                ByVal pStudyManager As StudyManager, _
                ByVal pValueCache As valueCache, _
                ByVal pTarget As Object, _
                ByVal pSourceStudy As IStudy, _
                ByVal pNumberOfValuesToReplay As Long, _
                ByVal pDiscriminator As Long) As CacheReplayTask
Const ProcName As String = "CreateCacheReplayTask"
On Error GoTo Err

Set CreateCacheReplayTask = New CacheReplayTask
CreateCacheReplayTask.Initialise pStudyManager, _
                            pValueCache, _
                            pTarget, _
                            pSourceStudy, _
                            pNumberOfValuesToReplay, _
                            pDiscriminator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyPoint( _
                ByVal X As Date, _
                ByVal Y As Double) As StudyPoint
Const ProcName As String = "CreateStudyPoint"
On Error GoTo Err

Set CreateStudyPoint = New StudyPoint
CreateStudyPoint.X = X
CreateStudyPoint.Y = Y

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DefaultStudyValueName() As String
Const ProcName As String = "DefaultStudyValueName"
On Error GoTo Err

DefaultStudyValueName = DefaultStudyValueNameStr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MovingAverageStudyValueName() As String
Const ProcName As String = "MovingAverageStudyValueName"
On Error GoTo Err

MovingAverageStudyValueName = MovingAverageStudyValueNameStr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SetDefaultStudyLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Const ProcName As String = "SetDefaultStudyLibraryConfig"
On Error GoTo Err

Dim currSLsList As ConfigurationSection
Dim currSL As ConfigurationSection

Set currSLsList = configdata.AddConfigurationSection(ConfigNameStudyLibraries)

Set currSLsList = configdata.AddConfigurationSection(ConfigNameStudyLibraries, , StudyLibrariesRenderer)

Set currSL = currSLsList.AddConfigurationSection(ConfigNameStudyLibrary & "(" & BuiltInStudyLibraryName & ")")

currSL.SetAttribute AttributeNameEnabled, "True"
currSL.SetAttribute AttributeNameStudyLibraryBuiltIn, "True"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function setupBarsStudy( _
                ByVal pStudyName As String, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pParams As Parameters, _
                ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupBarsStudy"
On Error GoTo Err

ReDim InputValueNames(3) As String
InputValueNames(0) = InputNameTrade
InputValueNames(1) = InputNameVolume
InputValueNames(2) = InputNameTickVolume
InputValueNames(3) = InputNameOpenInterest

Dim lBarStudy As IBarStudy
Set lBarStudy = pStudyBase.StudyManager.AddStudy(pStudyName, _
                                        pStudyBase.BaseStudy, _
                                        InputValueNames, _
                                        pParams)
If Not pInitialBarsFuture Is Nothing Then lBarStudy.InitialBarsFuture = pInitialBarsFuture
Set setupBarsStudy = lBarStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantMomentumBarsStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupConstantMomentumBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstMomentumBarsParamTicksPerBar, pTimePeriod.Length

Set setupConstantMomentumBarsStudy = setupBarsStudy(ConstMomentumBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantTickVolumeBarsStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupConstantTickVolumeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstTickVolumeBarsParamTicksPerBar, pTimePeriod.Length

Set setupConstantTickVolumeBarsStudy = setupBarsStudy(ConstTickVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantTimeBarsStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupConstantTimeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstTimeBarsParamBarLength, pTimePeriod.Length
lParams.SetParameterValue ConstTimeBarsParamTimeUnits, _
                        TimePeriodUnitsToString(pTimePeriod.Units)

Set setupConstantTimeBarsStudy = setupBarsStudy(ConstTimeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantVolumeBarsStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupConstantVolumeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstVolumeBarsParamVolPerBar, pTimePeriod.Length

Set setupConstantVolumeBarsStudy = setupBarsStudy(ConstVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


