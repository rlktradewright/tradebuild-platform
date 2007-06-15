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

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements StudyLibrary

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mHandle As Long

Private mStudyNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = "TradeWright Common Studies Library"

ReDim mStudyNames(0) As String
addStudyName AccDistName
addStudyName AtrName
addStudyName BbName
addStudyName ConstMomentumBarsName
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

'================================================================================
' StudyLibrary Interface Members
'================================================================================

Private Function StudyLibrary_createStudy( _
                ByVal name As String) As Study
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set StudyLibrary_createStudy = New accDist
Case UCase$(AtrName), UCase$(AtrShortName)
    Set StudyLibrary_createStudy = New ATR
Case UCase$(BbName), UCase$(BbShortName)
    Set StudyLibrary_createStudy = New BollingerBands
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set StudyLibrary_createStudy = New ConstMomentumBars
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set StudyLibrary_createStudy = New ConstTimeBars
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set StudyLibrary_createStudy = New ConstVolumeBars
Case UCase$(DoncName), UCase$(DoncName)
    Set StudyLibrary_createStudy = New DonchianChannels
Case UCase$(EmaName), UCase$(EmaShortName)
    Set StudyLibrary_createStudy = New EMA
Case UCase$(FIName), UCase$(FIShortName)
    Set StudyLibrary_createStudy = New ForceIndex
Case UCase$(MacdName), UCase$(MacdShortName)
    Set StudyLibrary_createStudy = New MACD
Case UCase$(PsName), UCase$(PsShortName)
    Set StudyLibrary_createStudy = New ParabolicStop
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_createStudy = New RSI
Case UCase$(SmaName), UCase$(SmaShortName)
    Set StudyLibrary_createStudy = New SMA
Case UCase$(SStochName), UCase$(SStochShortName)
    Set StudyLibrary_createStudy = New SlowStochastic
Case UCase$(SdName), UCase$(SdShortName)
    Set StudyLibrary_createStudy = New StandardDeviation
Case UCase$(StochName), UCase$(StochShortName)
    Set StudyLibrary_createStudy = New Stochastic
Case UCase$(SwingName), UCase$(SwingShortName)
    Set StudyLibrary_createStudy = New Swing
End Select
End Function

Private Function StudyLibrary_getImplementedStudyNames() As String()
StudyLibrary_getImplementedStudyNames = mStudyNames
End Function

Private Function StudyLibrary_getStudyDefaultParameters( _
                ByVal name As String) As Parameters
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set StudyLibrary_getStudyDefaultParameters = GAccDist.defaultParameters
Case UCase$(AtrName), UCase$(AtrShortName)
    Set StudyLibrary_getStudyDefaultParameters = GATR.defaultParameters
Case UCase$(BbName), UCase$(BbShortName)
    Set StudyLibrary_getStudyDefaultParameters = GBollingerBands.defaultParameters
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set StudyLibrary_getStudyDefaultParameters = GConstMomentumBars.defaultParameters
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set StudyLibrary_getStudyDefaultParameters = GConstTimeBars.defaultParameters
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set StudyLibrary_getStudyDefaultParameters = GConstVolumeBars.defaultParameters
Case UCase$(DoncName), UCase$(DoncName)
    Set StudyLibrary_getStudyDefaultParameters = GDonchianChannels.defaultParameters
Case UCase$(EmaName), UCase$(EmaShortName)
    Set StudyLibrary_getStudyDefaultParameters = GEMA.defaultParameters
Case UCase$(FIName), UCase$(FIShortName)
    Set StudyLibrary_getStudyDefaultParameters = GForceIndex.defaultParameters
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_getStudyDefaultParameters = GRSI.defaultParameters
Case UCase$(SmaName), UCase$(SmaShortName)
    Set StudyLibrary_getStudyDefaultParameters = GSMA.defaultParameters
Case UCase$(MacdName), UCase$(MacdShortName)
    Set StudyLibrary_getStudyDefaultParameters = GMACD.defaultParameters
Case UCase$(PsName), UCase$(PsShortName)
    Set StudyLibrary_getStudyDefaultParameters = GParabolicStop.defaultParameters
Case UCase$(SdName), UCase$(SdShortName)
    Set StudyLibrary_getStudyDefaultParameters = GStandardDeviation.defaultParameters
Case UCase$(SStochName), UCase$(SStochShortName)
    Set StudyLibrary_getStudyDefaultParameters = GSlowStochastic.defaultParameters
Case UCase$(StochName), UCase$(StochShortName)
    Set StudyLibrary_getStudyDefaultParameters = GStochastic.defaultParameters
Case UCase$(SwingName), UCase$(SwingShortName)
    Set StudyLibrary_getStudyDefaultParameters = GSwing.defaultParameters
End Select
End Function

Private Function StudyLibrary_getStudyDefinition( _
                ByVal name As String) As StudyDefinition
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set StudyLibrary_getStudyDefinition = GAccDist.StudyDefinition
Case UCase$(AtrName), UCase$(AtrShortName)
    Set StudyLibrary_getStudyDefinition = GATR.StudyDefinition
Case UCase$(BbName), UCase$(BbShortName)
    Set StudyLibrary_getStudyDefinition = GBollingerBands.StudyDefinition
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    Set StudyLibrary_getStudyDefinition = GConstMomentumBars.StudyDefinition
Case UCase$(ConstTimeBarsName), UCase$(ConstTimeBarsShortName)
    Set StudyLibrary_getStudyDefinition = GConstTimeBars.StudyDefinition
Case UCase$(ConstVolBarsName), UCase$(ConstVolBarsShortName)
    Set StudyLibrary_getStudyDefinition = GConstVolumeBars.StudyDefinition
Case UCase$(DoncName), UCase$(DoncName)
    Set StudyLibrary_getStudyDefinition = GDonchianChannels.StudyDefinition
Case UCase$(EmaName), UCase$(EmaShortName)
    Set StudyLibrary_getStudyDefinition = GEMA.StudyDefinition
Case UCase$(FIName), UCase$(FIShortName)
    Set StudyLibrary_getStudyDefinition = GForceIndex.StudyDefinition
Case UCase$(RsiName), UCase$(RsiShortName)
    Set StudyLibrary_getStudyDefinition = GRSI.StudyDefinition
Case UCase$(SmaName), UCase$(SmaShortName)
    Set StudyLibrary_getStudyDefinition = GSMA.StudyDefinition
Case UCase$(MacdName), UCase$(MacdShortName)
    Set StudyLibrary_getStudyDefinition = GMACD.StudyDefinition
Case UCase$(PsName), UCase$(PsShortName)
    Set StudyLibrary_getStudyDefinition = GParabolicStop.StudyDefinition
Case UCase$(SdName), UCase$(SdShortName)
    Set StudyLibrary_getStudyDefinition = GStandardDeviation.StudyDefinition
Case UCase$(SStochName), UCase$(SStochShortName)
    Set StudyLibrary_getStudyDefinition = GSlowStochastic.StudyDefinition
Case UCase$(StochName), UCase$(StochShortName)
    Set StudyLibrary_getStudyDefinition = GStochastic.StudyDefinition
Case UCase$(SwingName), UCase$(SwingShortName)
    Set StudyLibrary_getStudyDefinition = GSwing.StudyDefinition
End Select
End Function

Private Sub StudyLibrary_initialise( _
                ByVal libManager As StudyLibraryManager, _
                ByVal handle As Long)
Set gLibraryManager = libManager
mHandle = handle
End Sub

Private Property Let StudyLibrary_name( _
                ByVal RHS As String)
mName = RHS
End Property

Private Property Get StudyLibrary_name() As String
StudyLibrary_name = mName
End Property

Private Function StudyLibrary_setStudyDefaultParameters( _
                ByVal name As String, _
                ByVal params As Parameters) As Boolean

StudyLibrary_setStudyDefaultParameters = True

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    GAccDist.defaultParameters = params
Case UCase$(AtrName), UCase$(AtrShortName)
    GATR.defaultParameters = params
Case UCase$(BbName), UCase$(BbShortName)
    GBollingerBands.defaultParameters = params
Case UCase$(ConstMomentumBarsName), UCase$(ConstMomentumBarsShortName)
    GConstMomentumBars.defaultParameters = params
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
    StudyLibrary_setStudyDefaultParameters = False
End Select

End Function

Private Sub StudyLibrary_terminate()
' nothing to do
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub addStudyName(ByVal name As String)
If mStudyNames(0) <> "" Then
    ReDim Preserve mStudyNames(UBound(mStudyNames) + 1) As String
End If
mStudyNames(UBound(mStudyNames)) = name
End Sub

