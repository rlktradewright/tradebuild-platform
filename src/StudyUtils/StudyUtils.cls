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

Public Property Get ConstMomentumBarsParamTicksPerBar() As String
ConstMomentumBarsParamTicksPerBar = Globals.ConstMomentumBarsParamTicksPerBar
End Property

Public Property Get ConstMomentumBarsStudyName() As String
ConstMomentumBarsStudyName = Globals.ConstMomentumBarsStudyName
End Property

Public Property Get ConstMomentumBarsStudyShortName() As String
ConstMomentumBarsStudyShortName = Globals.ConstMomentumBarsStudyShortName
End Property

Public Property Get ConstTickVolumeBarsParamTicksPerBar() As String
ConstTickVolumeBarsParamTicksPerBar = Globals.ConstTickVolumeBarsParamTicksPerBar
End Property

Public Property Get ConstTickVolumeBarsStudyName() As String
ConstTickVolumeBarsStudyName = Globals.ConstTickVolumeBarsStudyName
End Property

Public Property Get ConstTickVolumeBarsStudyShortName() As String
ConstTickVolumeBarsStudyShortName = Globals.ConstTickVolumeBarsStudyShortName
End Property

Public Property Get ConstTimeBarsParamBarLength() As String
ConstTimeBarsParamBarLength = Globals.ConstTimeBarsParamBarLength
End Property

Public Property Get ConstTimeBarsParamTimeUnits() As String
ConstTimeBarsParamTimeUnits = Globals.ConstTimeBarsParamTimeUnits
End Property

Public Property Get ConstTimeBarsStudyName() As String
ConstTimeBarsStudyName = Globals.ConstTimeBarsStudyName
End Property

Public Property Get ConstTimeBarsStudyShortName() As String
ConstTimeBarsStudyShortName = Globals.ConstTimeBarsStudyShortName
End Property

Public Property Get ConstVolumeBarsParamVolPerBar() As String
ConstVolumeBarsParamVolPerBar = Globals.ConstVolumeBarsParamVolPerBar
End Property

Public Property Get ConstVolumeBarsStudyShortName() As String
ConstVolumeBarsStudyShortName = Globals.ConstVolumeBarsStudyShortName
End Property

Public Property Get ConstVolumeBarsStudyName() As String
ConstVolumeBarsStudyName = Globals.ConstVolumeBarsStudyName
End Property

Public Property Get InputNameAsk() As String
InputNameAsk = Globals.AskInputName
End Property

Public Property Get InputNameBarNumber() As String
InputNameBarNumber = Globals.BarNumberInputName
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

Public Property Get InputNameValue() As String
InputNameValue = Globals.ValueInputName
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = Globals.VolumeInputName
End Property

Public Property Get UserDefinedBarsStudyName() As String
UserDefinedBarsStudyName = Globals.UserDefinedBarsStudyName
End Property

Public Property Get UserDefinedBarsStudyShortName() As String
UserDefinedBarsStudyShortName = Globals.UserDefinedBarsStudyShortName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BarStudyValueOpen() As String
BarStudyValueOpen = Globals.BarStudyValueOpen
End Function

Public Function BarStudyValueHigh() As String
BarStudyValueHigh = Globals.BarStudyValueHigh
End Function

Public Function BarStudyValueLow() As String
BarStudyValueLow = Globals.BarStudyValueLow
End Function

Public Function BarStudyValueClose() As String
BarStudyValueClose = Globals.BarStudyValueClose
End Function

Public Function BarStudyValueVolume() As String
BarStudyValueVolume = Globals.BarStudyValueVolume
End Function

Public Function BarStudyValueTickVolume() As String
BarStudyValueTickVolume = Globals.BarStudyValueTickVolume
End Function

Public Function BarStudyValueOpenInterest() As String
BarStudyValueOpenInterest = Globals.BarStudyValueOpenInterest
End Function

Public Function BarStudyValueHL2() As String
BarStudyValueHL2 = Globals.BarStudyValueHL2
End Function

Public Function BarStudyValueHLC3() As String
BarStudyValueHLC3 = Globals.BarStudyValueHLC3
End Function

Public Function BarStudyValueOHLC4() As String
BarStudyValueOHLC4 = Globals.BarStudyValueOHLC4
End Function

Public Function CreateBarStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "CreateBarStudy"
On Error GoTo Err

If pTimePeriod.Length = 0 Then
    Set CreateBarStudy = setupUserDefinedBarsStudy(pTimePeriod, pStudyBase, pInitialBarsFuture)
Else
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
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarStudySupport( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionFuture As IFuture, _
                ByVal pPriceTickSize As Double) As BarStudySupport
Const ProcName As String = "CreateBarStudySupport"
On Error GoTo Err

Set CreateBarStudySupport = New BarStudySupport
CreateBarStudySupport.Initialise pTimePeriod, pSessionFuture, pPriceTickSize

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

Public Function CreateStudyBaseForDoubleInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pQuantum As Double = 0#, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForTickDataInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument pQuantum >= 0, "pQuantum is negative"

Dim lStudyBase As New StudyBaseForDoubleInput
lStudyBase.Initialise pStudyManager, pQuantum, pName

Set CreateStudyBaseForDoubleInput = lStudyBase

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForIntegerInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForIntegerInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"

Dim lStudyBase As New StudyBaseForIntegerInput
lStudyBase.Initialise pStudyManager, pName

Set CreateStudyBaseForIntegerInput = lStudyBase

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForNullInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForNullInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"

Dim lStudyBase As New StudyBaseForNullInput
lStudyBase.Initialise pStudyManager, pName

Set CreateStudyBaseForNullInput = lStudyBase

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForTickDataInput( _
                ByVal pStudyManager As StudyManager, _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture) As IStudyBase
Const ProcName As String = "CreateStudyBaseForTickDataInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument Not pTickSource Is Nothing, "pTickSource Is Nothing"
AssertArgument Not pContractFuture Is Nothing, "pContractFuture Is Nothing"

Dim lStudyBase As New StudyBaseForTickDataInput
lStudyBase.Initialise pStudyManager, pContractFuture

pTickSource.AddGenericTickListener lStudyBase

Set CreateStudyBaseForTickDataInput = lStudyBase

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForUserBarsInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pQuantum As Double = 0#, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForUserBarsInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument pQuantum >= 0, "pQuantum is negative"

Dim lStudyBase As New StudyBaseForUserBarsInput
lStudyBase.Initialise pStudyManager, pQuantum, pName

Set CreateStudyBaseForUserBarsInput = lStudyBase

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

Private Function setupBarStudy( _
                ByVal pStudyName As String, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pParams As Parameters, _
                ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupBarStudy"
On Error GoTo Err

If pStudyName = UserDefinedBarsStudyName Then
    ReDim InputValueNames(1) As String
    InputValueNames(0) = InputNameValue
    InputValueNames(1) = InputNameBarNumber
Else
    ReDim InputValueNames(3) As String
    InputValueNames(0) = InputNameTrade
    InputValueNames(1) = InputNameVolume
    InputValueNames(2) = InputNameTickVolume
    InputValueNames(3) = InputNameOpenInterest
End If

Dim lBarStudy As IBarStudy
Set lBarStudy = pStudyBase.StudyManager.AddStudy(pStudyName, _
                                        pStudyBase.BaseStudy, _
                                        InputValueNames, _
                                        pParams)
If Not pInitialBarsFuture Is Nothing Then lBarStudy.InitialBarsFuture = pInitialBarsFuture
Set setupBarStudy = lBarStudy

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

Set setupConstantMomentumBarsStudy = setupBarStudy(ConstMomentumBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

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

Set setupConstantTickVolumeBarsStudy = setupBarStudy(ConstTickVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

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

Set setupConstantTimeBarsStudy = setupBarStudy(ConstTimeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

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

Set setupConstantVolumeBarsStudy = setupBarStudy(ConstVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupUserDefinedBarsStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "setupUserDefinedBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters

Set setupUserDefinedBarsStudy = setupBarStudy(UserDefinedBarsStudyName, pStudyBase, lParams, pInitialBarsFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


