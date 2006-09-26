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
addStudyName AtrName
addStudyName BbName
addStudyName DoncName
addStudyName EmaName
addStudyName MacdName
addStudyName SmaName
addStudyName SdName
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
GATR.commonServiceConsumer = commonServiceConsumer
GEMA.commonServiceConsumer = commonServiceConsumer
GDonchianChannels.commonServiceConsumer = commonServiceConsumer
GSMA.commonServiceConsumer = commonServiceConsumer
GMACD.commonServiceConsumer = commonServiceConsumer
GStandardDeviation.commonServiceConsumer = commonServiceConsumer
GBollingerBands.commonServiceConsumer = commonServiceConsumer
GSwing.commonServiceConsumer = commonServiceConsumer

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
                ByVal name As String, _
                ByVal serviceConsumer As TradeBuildSP.IStudyServiceConsumer) As TradeBuildSP.IStudy
Select Case UCase$(name)
Case UCase$(AtrName)
    Dim lAtr As New ATR
    lAtr.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lAtr
Case UCase$(BbName)
    Dim lBb As New BollingerBands
    lBb.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lBb
Case UCase$(DoncName)
    Dim lDonc As New DonchianChannels
    lDonc.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lDonc
Case UCase$(EmaName)
    Dim lEMA As New EMA
    lEMA.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lEMA
Case UCase$(SmaName)
    Dim lSMA As New SMA
    lSMA.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lSMA
Case UCase$(MacdName)
    Dim lMACD As New MACD
    lMACD.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lMACD
Case UCase$(SdName)
    Dim lSD As New StandardDeviation
    lSD.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lSD
Case UCase$(SwingName)
    Dim lSwing As New Swing
    lSwing.initialise mCommonServiceConsumer, serviceConsumer
    Set IStudyServiceProvider_createStudy = lSwing
End Select

End Function

Private Function IStudyServiceProvider_getStudyDefaultParameters( _
                ByVal name As String) As TradeBuildSP.IParameters
Select Case UCase$(name)
Case UCase$(AtrName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GATR.defaultParameters
Case UCase$(BbName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GBollingerBands.defaultParameters
Case UCase$(DoncName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GDonchianChannels.defaultParameters
Case UCase$(EmaName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GEMA.defaultParameters
Case UCase$(SmaName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GSMA.defaultParameters
Case UCase$(MacdName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GMACD.defaultParameters
Case UCase$(SdName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GStandardDeviation.defaultParameters
Case UCase$(SwingName)
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
Case UCase$(AtrName)
    GATR.defaultParameters = parameters
Case UCase$(BbName)
    GBollingerBands.defaultParameters = parameters
Case UCase$(DoncName)
    GDonchianChannels.defaultParameters = parameters
Case UCase$(EmaName)
    GEMA.defaultParameters = parameters
Case UCase$(SmaName)
    GSMA.defaultParameters = parameters
Case UCase$(MacdName)
    GMACD.defaultParameters = parameters
Case UCase$(SdName)
    GStandardDeviation.defaultParameters = parameters
Case UCase$(SwingName)
    GSwing.defaultParameters = parameters
Case Else
    IStudyServiceProvider_setStudyDefaultParameters = False
End Select

End Function

Private Function IStudyServiceProvider_studyDefinition( _
                ByVal name As String) As TradeBuildSP.IStudyDefinition
Select Case UCase$(name)
Case UCase$(AtrName)
    Set IStudyServiceProvider_studyDefinition = GATR.studyDefinition
Case UCase$(BbName)
    Set IStudyServiceProvider_studyDefinition = GBollingerBands.studyDefinition
Case UCase$(DoncName)
    Set IStudyServiceProvider_studyDefinition = GDonchianChannels.studyDefinition
Case UCase$(EmaName)
    Set IStudyServiceProvider_studyDefinition = GEMA.studyDefinition
Case UCase$(SmaName)
    Set IStudyServiceProvider_studyDefinition = GSMA.studyDefinition
Case UCase$(MacdName)
    Set IStudyServiceProvider_studyDefinition = GMACD.studyDefinition
Case UCase$(SdName)
    Set IStudyServiceProvider_studyDefinition = GStandardDeviation.studyDefinition
Case UCase$(SwingName)
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

