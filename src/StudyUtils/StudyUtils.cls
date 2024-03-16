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
BuiltInStudyLibraryName = GStudyUtils.BuiltInStudyLibraryName
End Property

Public Property Get BuiltInStudyLibraryProgId() As String
BuiltInStudyLibraryProgId = GStudyUtils.BuiltInStudyLibraryProgId
End Property

Public Property Get ConstMomentumBarsParamTicksPerBar() As String
ConstMomentumBarsParamTicksPerBar = GStudyUtils.ConstMomentumBarsParamTicksPerBar
End Property

Public Property Get ConstMomentumBarsStudyName() As String
ConstMomentumBarsStudyName = GStudyUtils.ConstMomentumBarsStudyName
End Property

Public Property Get ConstMomentumBarsStudyShortName() As String
ConstMomentumBarsStudyShortName = GStudyUtils.ConstMomentumBarsStudyShortName
End Property

Public Property Get ConstTickVolumeBarsParamTicksPerBar() As String
ConstTickVolumeBarsParamTicksPerBar = GStudyUtils.ConstTickVolumeBarsParamTicksPerBar
End Property

Public Property Get ConstTickVolumeBarsStudyName() As String
ConstTickVolumeBarsStudyName = GStudyUtils.ConstTickVolumeBarsStudyName
End Property

Public Property Get ConstTickVolumeBarsStudyShortName() As String
ConstTickVolumeBarsStudyShortName = GStudyUtils.ConstTickVolumeBarsStudyShortName
End Property

Public Property Get ConstTimeBarsParamBarLength() As String
ConstTimeBarsParamBarLength = GStudyUtils.ConstTimeBarsParamBarLength
End Property

Public Property Get ConstTimeBarsParamTimeUnits() As String
ConstTimeBarsParamTimeUnits = GStudyUtils.ConstTimeBarsParamTimeUnits
End Property

Public Property Get ConstTimeBarsStudyName() As String
ConstTimeBarsStudyName = GStudyUtils.ConstTimeBarsStudyName
End Property

Public Property Get ConstTimeBarsStudyShortName() As String
ConstTimeBarsStudyShortName = GStudyUtils.ConstTimeBarsStudyShortName
End Property

Public Property Get ConstVolumeBarsParamVolPerBar() As String
ConstVolumeBarsParamVolPerBar = GStudyUtils.ConstVolumeBarsParamVolPerBar
End Property

Public Property Get ConstVolumeBarsStudyShortName() As String
ConstVolumeBarsStudyShortName = GStudyUtils.ConstVolumeBarsStudyShortName
End Property

Public Property Get ConstVolumeBarsStudyName() As String
ConstVolumeBarsStudyName = GStudyUtils.ConstVolumeBarsStudyName
End Property

Public Property Get InputNameAsk() As String
InputNameAsk = GStudyUtils.InputNameAsk
End Property

Public Property Get InputNameBarNumber() As String
InputNameBarNumber = GStudyUtils.InputNameBarNumber
End Property

Public Property Get InputNameBid() As String
InputNameBid = GStudyUtils.InputNameBid
End Property

Public Property Get InputNameOpenInterest() As String
InputNameOpenInterest = GStudyUtils.InputNameOpenInterest
End Property

Public Property Get InputNameTickVolume() As String
InputNameTickVolume = GStudyUtils.InputNameTickVolume
End Property

Public Property Get InputNameTrade() As String
InputNameTrade = GStudyUtils.InputNameTrade
End Property

Public Property Get InputNameValue() As String
InputNameValue = GStudyUtils.InputNameValue
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = GStudyUtils.InputNameVolume
End Property

Public Property Get UserDefinedBarsStudyName() As String
UserDefinedBarsStudyName = GStudyUtils.UserDefinedBarsStudyName
End Property

Public Property Get UserDefinedBarsStudyShortName() As String
UserDefinedBarsStudyShortName = GStudyUtils.UserDefinedBarsStudyShortName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BarStudyValueBar() As String
BarStudyValueBar = GStudyUtils.BarStudyValueBar
End Function

Public Function BarStudyValueOpen() As String
BarStudyValueOpen = GStudyUtils.BarStudyValueOpen
End Function

Public Function BarStudyValueHigh() As String
BarStudyValueHigh = GStudyUtils.BarStudyValueHigh
End Function

Public Function BarStudyValueLow() As String
BarStudyValueLow = GStudyUtils.BarStudyValueLow
End Function

Public Function BarStudyValueClose() As String
BarStudyValueClose = GStudyUtils.BarStudyValueClose
End Function

Public Function BarStudyValueVolume() As String
BarStudyValueVolume = GStudyUtils.BarStudyValueVolume
End Function

Public Function BarStudyValueTickVolume() As String
BarStudyValueTickVolume = GStudyUtils.BarStudyValueTickVolume
End Function

Public Function BarStudyValueOpenInterest() As String
BarStudyValueOpenInterest = GStudyUtils.BarStudyValueOpenInterest
End Function

Public Function BarStudyValueHL2() As String
BarStudyValueHL2 = GStudyUtils.BarStudyValueHL2
End Function

Public Function BarStudyValueHLC3() As String
BarStudyValueHLC3 = GStudyUtils.BarStudyValueHLC3
End Function

Public Function BarStudyValueOHLC4() As String
BarStudyValueOHLC4 = GStudyUtils.BarStudyValueOHLC4
End Function

Public Function CreateBarStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pIncludeDataOutsideSession As Boolean, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "CreateBarStudy"
On Error GoTo Err

Set CreateBarStudy = GStudyUtils.CreateBarStudy( _
                                    pAllowInitialBars, _
                                    pTimePeriod, _
                                    pStudyBase, _
                                    pIncludeDataOutsideSession, _
                                    pInitialBarsFuture)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarStudySupport( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal pPriceTickSize As Double) As BarStudySupport
Const ProcName As String = "CreateBarStudySupport"
On Error GoTo Err

Set CreateBarStudySupport = GStudyUtils.CreateBarStudySupport(pTimePeriod, pSession, pPriceTickSize)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateCacheReplayTask( _
                ByVal pStudyManager As StudyManager, _
                ByVal pValueCache As ValueCache, _
                ByVal pTarget As Object, _
                ByVal pSourceStudy As IStudy, _
                ByVal pNumberOfValuesToReplay As Long, _
                ByVal pDiscriminator As Long) As CacheReplayTask
Const ProcName As String = "CreateCacheReplayTask"
On Error GoTo Err

Set CreateCacheReplayTask = GStudyUtils.CreateCacheReplayTask(pStudyManager, _
                            pValueCache, _
                            pTarget, _
                            pSourceStudy, _
                            pNumberOfValuesToReplay, _
                            pDiscriminator)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForDoubleInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pQuantum As Double = 0#, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForDoubleInput"
On Error GoTo Err

Set CreateStudyBaseForDoubleInput = GStudyUtils.CreateStudyBaseForDoubleInput( _
                                                pStudyManager, _
                                                pQuantum, _
                                                pName)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForIntegerInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForIntegerInput"
On Error GoTo Err

Set CreateStudyBaseForIntegerInput = GStudyUtils.CreateStudyBaseForIntegerInput( _
                                                pStudyManager, _
                                                pName)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForNullInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForNullInput"
On Error GoTo Err

Set CreateStudyBaseForNullInput = GStudyUtils.CreateStudyBaseForNullInput( _
                                                pStudyManager, _
                                                pName)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForTickDataInput( _
                ByVal pStudyManager As StudyManager, _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture) As IStudyBase
Const ProcName As String = "CreateStudyBaseForTickDataInput"
On Error GoTo Err

Set CreateStudyBaseForTickDataInput = GStudyUtils.CreateStudyBaseForTickDataInput( _
                                                pStudyManager, _
                                                pTickSource, _
                                                pContractFuture)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForTickDataInputWithContract( _
                ByVal pStudyManager As StudyManager, _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContract As IContract) As IStudyBase
Const ProcName As String = "CreateStudyBaseForTickDataInputWithContract"
On Error GoTo Err

Set CreateStudyBaseForTickDataInputWithContract = GStudyUtils.CreateStudyBaseForTickDataInputWithContract( _
                                                pStudyManager, _
                                                pTickSource, _
                                                pContract)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyPoint( _
                ByVal X As Date, _
                ByVal Y As Double) As StudyPoint
Const ProcName As String = "CreateStudyPoint"
On Error GoTo Err

Set CreateStudyPoint = GStudyUtils.CreateStudyPoint(X, Y)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function DefaultStudyValueName() As String
Const ProcName As String = "DefaultStudyValueName"
On Error GoTo Err

DefaultStudyValueName = GStudyUtils.DefaultStudyValueName

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MovingAverageStudyValueName() As String
Const ProcName As String = "MovingAverageStudyValueName"
On Error GoTo Err

MovingAverageStudyValueName = GStudyUtils.MovingAverageStudyValueName

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SetDefaultStudyLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Const ProcName As String = "SetDefaultStudyLibraryConfig"
On Error GoTo Err

GStudyUtils.SetDefaultStudyLibraryConfig configdata

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

