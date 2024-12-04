Attribute VB_Name = "GStudyUtils"
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

Private Const ModuleName                As String = "GStudyUtils"

Public Const AskInputName                       As String = "Ask"
Public Const BidInputName                       As String = "Bid"
Public Const OpenInterestInputName              As String = "Open interest"
Public Const TickVolumeInputName                As String = "Tick Volume"
Public Const TradeInputName                     As String = "Trade"
Public Const VolumeInputName                    As String = "Total Volume"
Public Const ValueInputName                     As String = "Value"
Public Const BarNumberInputName                 As String = "Bar number"

Public Const AttributeNameEnabled               As String = "Enabled"
Public Const AttributeNameStudyLibraryBuiltIn   As String = "BuiltIn"
Public Const AttributeNameStudyLibraryProgId    As String = "ProgId"

Public Const BuiltInStudyLibProgId              As String = "CmnStudiesLib27.StudyLib"
Public Const BuiltInStudyLibName                As String = "BuiltIn"

Public Const ConfigNameStudyLibraries           As String = "StudyLibraries"
Public Const ConfigNameStudyLibrary             As String = "StudyLibrary"

Public Const ConstTickVolumeBarsStudyName       As String = "Constant Tick Volume bars"
Public Const ConstTickVolumeBarsStudyShortName  As String = "CTV Bars"
Public Const ConstTickVolumeBarsParamTicksPerBar As String = "Ticks per bar"

Public Const ConstTimeBarsStudyName             As String = "Constant Time Bars"
Public Const ConstTimeBarsStudyShortName        As String = "Bars"
Public Const ConstTimeBarsParamBarLength        As String = "Bar length"
Public Const ConstTimeBarsParamTimeUnits        As String = "Time units"

Public Const ConstVolumeBarsStudyName           As String = "Constant Volume bars"
Public Const ConstVolumeBarsStudyShortName      As String = "CV Bars"
Public Const ConstVolumeBarsParamVolPerBar      As String = "Volume per bar"

Public Const ConstMomentumBarsStudyName         As String = "Constant Momentum Bars"
Public Const ConstMomentumBarsStudyShortName    As String = "CM Bars"
Public Const ConstMomentumBarsParamTicksPerBar  As String = "Ticks move per bar"

Public Const UserDefinedBarsStudyName           As String = "User-defined Bars"
Public Const UserDefinedBarsStudyShortName      As String = "UD Bars"

Public Const DefaultStudyValueNameStr           As String = "$DEFAULT"
Public Const MovingAverageStudyValueNameStr     As String = "MA"

' sub-Value names for study values in bar mode
Public Const BarStudyValueBar                   As String = "Bar"
Public Const BarStudyValueOpen                  As String = "Open"
Public Const BarStudyValueHigh                  As String = "High"
Public Const BarStudyValueLow                   As String = "Low"
Public Const BarStudyValueClose                 As String = "Close"
Public Const BarStudyValueVolume                As String = "Volume"
Public Const BarStudyValueTickVolume            As String = "Tick Volume"
Public Const BarStudyValueOpenInterest          As String = "Open Interest"
Public Const BarStudyValueHL2                   As String = "(H+L)/2"
Public Const BarStudyValueHLC3                  As String = "(H+L+C)/3"
Public Const BarStudyValueOHLC4                 As String = "(O+H+L+C)/4"

Public Const StudyLibrariesRenderer             As String = "StudiesUI27.StudyLibConfigurer"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

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

Public Property Get InputNameAsk() As String
InputNameAsk = AskInputName
End Property

Public Property Get InputNameBarNumber() As String
InputNameBarNumber = BarNumberInputName
End Property

Public Property Get InputNameBid() As String
InputNameBid = BidInputName
End Property

Public Property Get InputNameOpenInterest() As String
InputNameOpenInterest = OpenInterestInputName
End Property

Public Property Get InputNameTickVolume() As String
InputNameTickVolume = TickVolumeInputName
End Property

Public Property Get InputNameTrade() As String
InputNameTrade = TradeInputName
End Property

Public Property Get InputNameValue() As String
InputNameValue = ValueInputName
End Property

Public Property Get InputNameVolume() As String
InputNameVolume = VolumeInputName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateBarStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pIncludeDataOutsideSession As Boolean, _
                Optional ByVal pInitialBarsFuture As IFuture) As IBarStudy
Const ProcName As String = "CreateBarStudy"
On Error GoTo Err

AssertArgument pAllowInitialBars Or pInitialBarsFuture Is Nothing, "Either pAllowInitialBars must be true or pInitialBarsFuture must be Nothing"

If pTimePeriod.Units = TimePeriodNone Or pTimePeriod.Length = 0 Then
    Set CreateBarStudy = setupUserDefinedBarsStudy(pAllowInitialBars, pTimePeriod, pStudyBase, pInitialBarsFuture, pIncludeDataOutsideSession)
Else
    Select Case pTimePeriod.Units
    Case TimePeriodSecond, _
            TimePeriodMinute, _
            TimePeriodHour, _
            TimePeriodDay, _
            TimePeriodWeek, _
            TimePeriodMonth, _
            TimePeriodYear
        Set CreateBarStudy = setupConstantTimeBarsStudy(pAllowInitialBars, pTimePeriod, pStudyBase, pInitialBarsFuture, pIncludeDataOutsideSession)
    Case TimePeriodTickMovement
        Set CreateBarStudy = setupConstantMomentumBarsStudy(pAllowInitialBars, pTimePeriod, pStudyBase, pInitialBarsFuture, pIncludeDataOutsideSession)
    Case TimePeriodTickVolume
        Set CreateBarStudy = setupConstantTickVolumeBarsStudy(pAllowInitialBars, pTimePeriod, pStudyBase, pInitialBarsFuture, pIncludeDataOutsideSession)
    Case TimePeriodVolume
        Set CreateBarStudy = setupConstantVolumeBarsStudy(pAllowInitialBars, pTimePeriod, pStudyBase, pInitialBarsFuture, pIncludeDataOutsideSession)
    End Select
End If

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

Set CreateBarStudySupport = New BarStudySupport
CreateBarStudySupport.Initialise pTimePeriod, pSession, pPriceTickSize

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

Set CreateCacheReplayTask = New CacheReplayTask
CreateCacheReplayTask.Initialise pStudyManager, _
                            pValueCache, _
                            pTarget, _
                            pSourceStudy, _
                            pNumberOfValuesToReplay, _
                            pDiscriminator

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForDecimalInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pQuantum As BoxedDecimal, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForDecimalInput"
On Error GoTo Err

If pQuantum Is Nothing Then Set pQuantum = DecimalZero

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument pQuantum.DecimalValue >= 0, "pQuantum is negative"

Dim lStudyBase As New StudyBaseForDecimalInput
lStudyBase.Initialise pStudyManager, pQuantum, pName

Set CreateStudyBaseForDecimalInput = lStudyBase

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForDoubleInput( _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pQuantum As BoxedDecimal, _
                Optional ByVal pName As String) As IStudyBase
Const ProcName As String = "CreateStudyBaseForDoubleInput"
On Error GoTo Err

If pQuantum Is Nothing Then Set pQuantum = DecimalZero

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument pQuantum.DecimalValue >= 0, "pQuantum is negative"

Dim lStudyBase As New StudyBaseForDoubleInput
lStudyBase.Initialise pStudyManager, pQuantum, pName

Set CreateStudyBaseForDoubleInput = lStudyBase

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
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
GStudies.HandleUnexpectedError ProcName, ModuleName
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
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStudyBaseForTickDataInput( _
                ByVal pStudyManager As StudyManager, _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture) As IStudyBase
Const ProcName As String = "CreateStudyBaseForTickDataInput"
On Error GoTo Err

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument Not pContractFuture Is Nothing, "pContractFuture Is Nothing"

Dim lStudyBase As New StudyBaseForTickDataInput
lStudyBase.Initialise pStudyManager, pContractFuture

If Not pTickSource Is Nothing Then pTickSource.AddGenericTickListener lStudyBase

Set CreateStudyBaseForTickDataInput = lStudyBase

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

AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"
AssertArgument Not pContract Is Nothing, "pContract Is Nothing"

Dim lStudyBase As New StudyBaseForTickDataInput
lStudyBase.InitialiseWithContract pStudyManager, pContract

If Not pTickSource Is Nothing Then pTickSource.AddGenericTickListener lStudyBase

Set CreateStudyBaseForTickDataInputWithContract = lStudyBase

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
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
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function DefaultStudyValueName() As String
Const ProcName As String = "DefaultStudyValueName"
On Error GoTo Err

DefaultStudyValueName = DefaultStudyValueNameStr

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MovingAverageStudyValueName() As String
Const ProcName As String = "MovingAverageStudyValueName"
On Error GoTo Err

MovingAverageStudyValueName = MovingAverageStudyValueNameStr

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
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
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function setupBarStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pStudyName As String, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pParams As Parameters, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupBarStudy"
On Error GoTo Err

If pStudyName = UserDefinedBarsStudyName Then
    ReDim InputValueNames(4) As String
    InputValueNames(0) = DefaultStudyValueName
    InputValueNames(1) = InputNameVolume
    InputValueNames(2) = InputNameTickVolume
    InputValueNames(3) = InputNameOpenInterest
    InputValueNames(4) = InputNameBarNumber
Else
    ReDim InputValueNames(3) As String
    InputValueNames(0) = DefaultStudyValueName
    InputValueNames(1) = InputNameVolume
    InputValueNames(2) = InputNameTickVolume
    InputValueNames(3) = InputNameOpenInterest
End If

GStudies.Logger.Log "Adding study: " & pStudyName & " to " & pStudyBase.BaseStudy.Name, ProcName, ModuleName, LogLevelHighDetail
Dim lBarStudy As IBarStudy
Set lBarStudy = pStudyBase.StudyManager.AddStudy(pStudyName, _
                                        pStudyBase.BaseStudy, _
                                        InputValueNames, _
                                        pIncludeDataOutsideSession, _
                                        pParams)
lBarStudy.AllowInitialBars = pAllowInitialBars
If Not pInitialBarsFuture Is Nothing Then lBarStudy.InitialBarsFuture = pInitialBarsFuture
Set setupBarStudy = lBarStudy

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantMomentumBarsStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupConstantMomentumBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstMomentumBarsParamTicksPerBar, pTimePeriod.Length

Set setupConstantMomentumBarsStudy = setupBarStudy(pAllowInitialBars, ConstMomentumBarsStudyName, pStudyBase, lParams, pInitialBarsFuture, pIncludeDataOutsideSession)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantTickVolumeBarsStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupConstantTickVolumeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstTickVolumeBarsParamTicksPerBar, pTimePeriod.Length

Set setupConstantTickVolumeBarsStudy = setupBarStudy(pAllowInitialBars, ConstTickVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture, pIncludeDataOutsideSession)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantTimeBarsStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupConstantTimeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstTimeBarsParamBarLength, pTimePeriod.Length
lParams.SetParameterValue ConstTimeBarsParamTimeUnits, _
                        TimePeriodUnitsToString(pTimePeriod.Units)

Set setupConstantTimeBarsStudy = setupBarStudy(pAllowInitialBars, ConstTimeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture, pIncludeDataOutsideSession)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupConstantVolumeBarsStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupConstantVolumeBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ConstVolumeBarsParamVolPerBar, pTimePeriod.Length

Set setupConstantVolumeBarsStudy = setupBarStudy(pAllowInitialBars, ConstVolumeBarsStudyName, pStudyBase, lParams, pInitialBarsFuture, pIncludeDataOutsideSession)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupUserDefinedBarsStudy( _
                ByVal pAllowInitialBars As Boolean, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pInitialBarsFuture As IFuture, _
                ByVal pIncludeDataOutsideSession As Boolean) As IBarStudy
Const ProcName As String = "setupUserDefinedBarsStudy"
On Error GoTo Err

Dim lParams As New Parameters

Set setupUserDefinedBarsStudy = setupBarStudy(pAllowInitialBars, UserDefinedBarsStudyName, pStudyBase, lParams, pInitialBarsFuture, pIncludeDataOutsideSession)

Exit Function

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Function




