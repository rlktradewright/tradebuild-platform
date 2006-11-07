VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyServiceProvider"
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

Implements TradeBuildSP.ICommonServiceProvider
Implements TradeBuildSP.IStudyServiceProvider

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
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mLogLevel As TradeBuildSP.LogLevels

Private mStudyNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = "TradeWright TradeBuild Built-in Studies Service Provider"

ReDim mStudyNames(0) As String
addStudyName AccDistName
addStudyName AtrName
addStudyName BbName
addStudyName DoncName
addStudyName EmaName
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
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
With details
    .Comments = App.Comments
    .EXEName = App.EXEName
    .FileDescription = App.FileDescription
    .LegalCopyright = App.LegalCopyright
    .LegalTrademarks = App.LegalTrademarks
    .Path = App.Path
    .ProductName = App.ProductName
    .Vendor = App.CompanyName
    .VersionMajor = App.Major
    .VersionMinor = App.Minor
    .VersionRevision = App.Revision
End With
ICommonServiceProvider_Details = details
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer
GAccDist.commonServiceConsumer = commonServiceConsumer
GATR.commonServiceConsumer = commonServiceConsumer
GEMA.commonServiceConsumer = commonServiceConsumer
GDonchianChannels.commonServiceConsumer = commonServiceConsumer
GSMA.commonServiceConsumer = commonServiceConsumer
GMACD.commonServiceConsumer = commonServiceConsumer
GParabolicStop.commonServiceConsumer = commonServiceConsumer
GStandardDeviation.commonServiceConsumer = commonServiceConsumer
GBollingerBands.commonServiceConsumer = commonServiceConsumer
GRSI.commonServiceConsumer = commonServiceConsumer
GSwing.commonServiceConsumer = commonServiceConsumer
GStochastic.commonServiceConsumer = commonServiceConsumer
GSlowStochastic.commonServiceConsumer = commonServiceConsumer

mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.Studies

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IStudyServiceProvider Interface Members
'================================================================================

Private Function IStudyServiceProvider_createStudy( _
                ByVal name As String) As TradeBuildSP.IStudy
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyServiceProvider_createStudy = New accDist
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyServiceProvider_createStudy = New ATR
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyServiceProvider_createStudy = New BollingerBands
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyServiceProvider_createStudy = New DonchianChannels
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyServiceProvider_createStudy = New EMA
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyServiceProvider_createStudy = New MACD
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyServiceProvider_createStudy = New ParabolicStop
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_createStudy = New RSI
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyServiceProvider_createStudy = New SMA
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyServiceProvider_createStudy = New SlowStochastic
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyServiceProvider_createStudy = New StandardDeviation
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyServiceProvider_createStudy = New Stochastic
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyServiceProvider_createStudy = New Swing
End Select

End Function

Private Function IStudyServiceProvider_getStudyDefaultParameters( _
                ByVal name As String) As TradeBuildSP.IParameters
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GAccDist.defaultParameters
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GATR.defaultParameters
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GBollingerBands.defaultParameters
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GDonchianChannels.defaultParameters
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GEMA.defaultParameters
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GRSI.defaultParameters
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GSMA.defaultParameters
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GMACD.defaultParameters
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GParabolicStop.defaultParameters
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GStandardDeviation.defaultParameters
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GSlowStochastic.defaultParameters
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GStochastic.defaultParameters
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GSwing.defaultParameters
End Select

End Function

Private Property Get IStudyServiceProvider_implementedStudyNames() As String()
IStudyServiceProvider_implementedStudyNames = mStudyNames
End Property

Private Function IStudyServiceProvider_setStudyDefaultParameters( _
                ByVal name As String, _
                ByVal parameters As TradeBuildSP.IParameters) As Boolean

IStudyServiceProvider_setStudyDefaultParameters = True

Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    GAccDist.defaultParameters = parameters
Case UCase$(AtrName), UCase$(AtrShortName)
    GATR.defaultParameters = parameters
Case UCase$(BbName), UCase$(BbShortName)
    GBollingerBands.defaultParameters = parameters
Case UCase$(DoncName), UCase$(DoncName)
    GDonchianChannels.defaultParameters = parameters
Case UCase$(EmaName), UCase$(EmaShortName)
    GEMA.defaultParameters = parameters
Case UCase$(RsiName), UCase$(RsiShortName)
    GRSI.defaultParameters = parameters
Case UCase$(SmaName), UCase$(SmaShortName)
    GSMA.defaultParameters = parameters
Case UCase$(MacdName), UCase$(MacdShortName)
    GMACD.defaultParameters = parameters
Case UCase$(PsName), UCase$(PsShortName)
    GParabolicStop.defaultParameters = parameters
Case UCase$(SdName), UCase$(SdShortName)
    GStandardDeviation.defaultParameters = parameters
Case UCase$(SStochName), UCase$(SStochShortName)
    GSlowStochastic.defaultParameters = parameters
Case UCase$(StochName), UCase$(StochShortName)
    GStochastic.defaultParameters = parameters
Case UCase$(SwingName), UCase$(SwingShortName)
    GSwing.defaultParameters = parameters
Case Else
    IStudyServiceProvider_setStudyDefaultParameters = False
End Select

End Function

Private Function IStudyServiceProvider_studyDefinition( _
                ByVal name As String) As TradeBuildSP.IStudyDefinition
Select Case UCase$(name)
Case UCase$(AccDistName), UCase$(AccDistShortName)
    Set IStudyServiceProvider_studyDefinition = GAccDist.studyDefinition
Case UCase$(AtrName), UCase$(AtrShortName)
    Set IStudyServiceProvider_studyDefinition = GATR.studyDefinition
Case UCase$(BbName), UCase$(BbShortName)
    Set IStudyServiceProvider_studyDefinition = GBollingerBands.studyDefinition
Case UCase$(DoncName), UCase$(DoncName)
    Set IStudyServiceProvider_studyDefinition = GDonchianChannels.studyDefinition
Case UCase$(EmaName), UCase$(EmaShortName)
    Set IStudyServiceProvider_studyDefinition = GEMA.studyDefinition
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_studyDefinition = GRSI.studyDefinition
Case UCase$(SmaName), UCase$(SmaShortName)
    Set IStudyServiceProvider_studyDefinition = GSMA.studyDefinition
Case UCase$(MacdName), UCase$(MacdShortName)
    Set IStudyServiceProvider_studyDefinition = GMACD.studyDefinition
Case UCase$(PsName), UCase$(PsShortName)
    Set IStudyServiceProvider_studyDefinition = GParabolicStop.studyDefinition
Case UCase$(SdName), UCase$(SdShortName)
    Set IStudyServiceProvider_studyDefinition = GStandardDeviation.studyDefinition
Case UCase$(SStochName), UCase$(SStochShortName)
    Set IStudyServiceProvider_studyDefinition = GSlowStochastic.studyDefinition
Case UCase$(StochName), UCase$(StochShortName)
    Set IStudyServiceProvider_studyDefinition = GStochastic.studyDefinition
Case UCase$(SwingName), UCase$(SwingShortName)
    Set IStudyServiceProvider_studyDefinition = GSwing.studyDefinition
End Select

End Function

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

