VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartLoader"

'@================================================================================
' Member variables
'@================================================================================

Private mChartConfig                                As ConfigurationSection
Private mChartForms                                 As ChartForms
Private mTickers                                    As Tickers
Private mBarFormatterLibManager                     As BarFormatterLibManager
Private mTimePeriodValidator                        As ITimePeriodValidator
Private mOwner                                      As Variant
Private mTheme                                      As ITheme

Private mFutureBuilder                              As New FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State = MarketDataSourceStates.MarketDataSourceStateReady Or _
    ev.State = MarketDataSourceStates.MarketDataSourceStatePaused Or _
    ev.State = MarketDataSourceStates.MarketDataSourceStateRunning _
Then
    If Not mChartForms.CreateChartFromConfig( _
                            mChartConfig, _
                            mTickers, _
                            mBarFormatterLibManager, _
                            mTimePeriodValidator) Then
        mFutureBuilder.Value = mChartConfig
        mFutureBuilder.Complete
    Else
        mFutureBuilder.Value = Nothing
        mFutureBuilder.Complete
    End If
End If

Dim lTicker As Ticker: Set lTicker = ev.Source
lTicker.RemoveStateChangeListener Me

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

'@================================================================================
' Methods
'@================================================================================

Friend Function LoadChart( _
                ByVal pChartConfig As ConfigurationSection, _
                ByVal pChartForms As ChartForms, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme) As IFuture
Const ProcName As String = "LoadChart"
On Error GoTo Err

Set LoadChart = mFutureBuilder.Future

Set mChartConfig = pChartConfig
Set mChartForms = pChartForms
Set mChartConfig = pChartConfig
Set mTickers = pTickers
Set mBarFormatterLibManager = pBarFormatterLibManager
Set mTimePeriodValidator = pTimePeriodValidator
gSetVariant mOwner, pOwner
Set mTheme = pTheme

Dim lTickerKey As String
lTickerKey = mChartConfig.GetSetting(ConfigSettingDataSourceKey, "")

Dim lTicker As Ticker
If lTickerKey = "" Then
    Set lTicker = mTickers.GetTicker(mChartConfig.InstanceQualifier)
Else
    Set lTicker = mTickers.GetTicker(lTickerKey)
End If

If lTicker Is Nothing Then
    mFutureBuilder.Value = mChartConfig
    mFutureBuilder.Complete
    Exit Function
End If

If Not (lTicker.State = MarketDataSourceStateReady Or _
        lTicker.State = MarketDataSourceStatePaused Or _
        lTicker.State = MarketDataSourceStateRunning) Then
    lTicker.AddStateChangeListener Me
    Exit Function
End If

If Not mChartForms.CreateChartFromConfig(mChartConfig, mTickers, mBarFormatterLibManager, mTimePeriodValidator) Then
    mFutureBuilder.Value = mChartConfig
    mFutureBuilder.Complete
    Exit Function
End If

mFutureBuilder.Value = Nothing
mFutureBuilder.Complete

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




