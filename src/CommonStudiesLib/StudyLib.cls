VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyLib"
Attribute VB_GlobalNameSpace = False
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
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IStudyLibrary

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyLib"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String

Private mStudyNames() As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = "TradeWright Common Studies Library"

ReDim mStudyNames(0) As String
addStudyName AccDistName
addStudyName AtrName
addStudyName BbName
addStudyName ConstMomentumBarsName
addStudyName ConstTickVolBarsName
addStudyName ConstTimeBarsName
addStudyName ConstVolBarsName
addStudyName DoncName
addStudyName EmaName
addStudyName FIName
addStudyName MacdName
addStudyName PsName
addStudyName RsiName
addStudyName SStochName
addStudyName SmaName
addStudyName SdName
addStudyName StochName
addStudyName SwingName
End Sub

'@================================================================================
' IStudyLibrary Interface Members
'@================================================================================

Private Function IStudyLibrary_createStudy( _
                ByVal name As String) As IStudy
Const ProcName As String = "IStudyLibrary_createStudy"
On Error GoTo Err

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyLibrary_createStudy = New accDist
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyLibrary_createStudy = New ATR
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyLibrary_createStudy = New BollingerBands
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set IStudyLibrary_createStudy = New ConstMomentumBars
Case UCase$(ConstTickVolBarsName), UCase$(ConstTickVolBarsShortName)
    Set IStudyLibrary_createStudy = New ConstTickVolumeBars
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set IStudyLibrary_createStudy = New ConstTimeBars
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set IStudyLibrary_createStudy = New ConstVolumeBars
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyLibrary_createStudy = New DonchianChannels
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyLibrary_createStudy = New EMA
Case UCase$(FIName), UCase$(FIShortName)
    Set IStudyLibrary_createStudy = New ForceIndex
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyLibrary_createStudy = New MACD
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyLibrary_createStudy = New ParabolicStop
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyLibrary_createStudy = New RSI
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyLibrary_createStudy = New SMA
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyLibrary_createStudy = New SlowStochastic
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyLibrary_createStudy = New StandardDeviation
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyLibrary_createStudy = New Stochastic
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyLibrary_createStudy = New Swing
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudyLibrary_getImplementedStudyNames() As String()
Const ProcName As String = "IStudyLibrary_getImplementedStudyNames"
On Error GoTo Err

IStudyLibrary_getImplementedStudyNames = mStudyNames

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudyLibrary_getStudyDefaultParameters( _
                ByVal name As String) As Parameters
Const ProcName As String = "IStudyLibrary_getStudyDefaultParameters"
On Error GoTo Err

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GAccDist.defaultParameters
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GATR.defaultParameters
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GBollingerBands.defaultParameters
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GConstMomentumBars.defaultParameters
Case UCase$(ConstTickVolBarsName), UCase$(ConstTickVolBarsShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GConstTickVolumeBars.defaultParameters
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GConstTimeBars.defaultParameters
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GConstVolumeBars.defaultParameters
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyLibrary_getStudyDefaultParameters = GDonchianChannels.defaultParameters
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GEMA.defaultParameters
Case UCase$(FIName), UCase$(FIShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GForceIndex.defaultParameters
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GRSI.defaultParameters
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GSMA.defaultParameters
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GMACD.defaultParameters
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GParabolicStop.defaultParameters
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GStandardDeviation.defaultParameters
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GSlowStochastic.defaultParameters
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GStochastic.defaultParameters
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyLibrary_getStudyDefaultParameters = GSwing.defaultParameters
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudyLibrary_getStudyDefinition( _
                ByVal name As String) As StudyDefinition
Const ProcName As String = "IStudyLibrary_getStudyDefinition"
On Error GoTo Err

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyLibrary_getStudyDefinition = GAccDist.StudyDefinition
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyLibrary_getStudyDefinition = GATR.StudyDefinition
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyLibrary_getStudyDefinition = GBollingerBands.StudyDefinition
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set IStudyLibrary_getStudyDefinition = GConstMomentumBars.StudyDefinition
Case UCase$(ConstTickVolBarsName), UCase$(ConstTickVolBarsShortName)
    Set IStudyLibrary_getStudyDefinition = GConstTickVolumeBars.StudyDefinition
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set IStudyLibrary_getStudyDefinition = GConstTimeBars.StudyDefinition
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set IStudyLibrary_getStudyDefinition = GConstVolumeBars.StudyDefinition
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyLibrary_getStudyDefinition = GDonchianChannels.StudyDefinition
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyLibrary_getStudyDefinition = GEMA.StudyDefinition
Case UCase$(FIName), UCase$(FIShortName)
    Set IStudyLibrary_getStudyDefinition = GForceIndex.StudyDefinition
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyLibrary_getStudyDefinition = GRSI.StudyDefinition
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyLibrary_getStudyDefinition = GSMA.StudyDefinition
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyLibrary_getStudyDefinition = GMACD.StudyDefinition
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyLibrary_getStudyDefinition = GParabolicStop.StudyDefinition
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyLibrary_getStudyDefinition = GStandardDeviation.StudyDefinition
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyLibrary_getStudyDefinition = GSlowStochastic.StudyDefinition
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyLibrary_getStudyDefinition = GStochastic.StudyDefinition
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyLibrary_getStudyDefinition = GSwing.StudyDefinition
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStudyLibrary_initialise( _
                ByVal libManager As StudyLibraryManager, _
                ByVal handle As Long)
Const ProcName As String = "IStudyLibrary_initialise"
On Error GoTo Err

' nothing to do

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IStudyLibrary_name( _
                ByVal RHS As String)
Const ProcName As String = "IStudyLibrary_name"
On Error GoTo Err

mName = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudyLibrary_name() As String
Const ProcName As String = "IStudyLibrary_name"
On Error GoTo Err

IStudyLibrary_name = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudyLibrary_SetStudyDefaultParameters( _
                ByVal name As String, _
                ByVal params As Parameters) As Boolean

Const ProcName As String = "IStudyLibrary_SetStudyDefaultParameters"
On Error GoTo Err

IStudyLibrary_SetStudyDefaultParameters = True

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    GAccDist.defaultParameters = params
Case UCase$(AtrName), UCase$(AtrShortName)
    GATR.defaultParameters = params
Case UCase$(BbName), UCase$(BbShortName)
    GBollingerBands.defaultParameters = params
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    GConstMomentumBars.defaultParameters = params
Case UCase$(ConstTickVolBarsName), UCase$(ConstTickVolBarsShortName)
    GConstTickVolumeBars.defaultParameters = params
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    GConstTimeBars.defaultParameters = params
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    GConstVolumeBars.defaultParameters = params
Case UCase$(DoncName), UCase$(DoncName)
    GDonchianChannels.defaultParameters = params
Case UCase$(EmaName), UCase$(EmaShortName)
    GEMA.defaultParameters = params
Case UCase$(FIName), UCase$(FIShortName)
    GForceIndex.defaultParameters = params
Case UCase$(RsiName), UCase$(RsiShortName)
    GRSI.defaultParameters = params
Case UCase$(SmaName), UCase$(SmaShortName)
    GSMA.defaultParameters = params
Case UCase$(MacdName), UCase$(MacdShortName)
    GMACD.defaultParameters = params
Case UCase$(PsName), UCase$(PsShortName)
    GParabolicStop.defaultParameters = params
Case UCase$(SdName), UCase$(SdShortName)
    GStandardDeviation.defaultParameters = params
Case UCase$(SStochName), UCase$(SStochShortName)
    GSlowStochastic.defaultParameters = params
Case UCase$(StochName), UCase$(StochShortName)
    GStochastic.defaultParameters = params
Case UCase$(SwingName), UCase$(SwingShortName)
    GSwing.defaultParameters = params
Case Else
    IStudyLibrary_SetStudyDefaultParameters = False
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub IStudyLibrary_terminate()
' nothing to do
Const ProcName As String = "IStudyLibrary_terminate"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let name(ByVal Value As String)
Const ProcName As String = "name"
On Error GoTo Err

mName = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get name() As String
Const ProcName As String = "name"
On Error GoTo Err

name = mName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addStudyName(ByVal name As String)
Const ProcName As String = "addStudyName"
On Error GoTo Err

If mStudyNames(0) <> "" Then
    ReDim Preserve mStudyNames(UBound(mStudyNames) + 1) As String
End If
mStudyNames(UBound(mStudyNames)) = name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

