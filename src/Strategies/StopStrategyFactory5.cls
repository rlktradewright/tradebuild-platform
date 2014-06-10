VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopStrategyFactory5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IStrategyFactory

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

Private Const ModuleName                            As String = "StopStrategyFactory5"

'@================================================================================
' Member variables
'@================================================================================

Private mTimeframe As Timeframe
Private mLongTimeframe As Timeframe
Private mSwing As Swing
Private mMACDStudy As MACD
Private mListenerKey As String

Private mStopTimeframe As Timeframe
Private mStopMA As EMA
Private mBoll As BollingerBands

Private mLongBoll As BollingerBands
Private mLongBollPeriods As Long

Private mBreakevenTimeframe As Timeframe

Private mTimePeriod As Integer
Private mTimePeriodUnit As String
Private mLongTimePeriod As Integer
Private mLongTimePeriodUnit As String

Private mLongMA As SMA
Private mLongMAPeriods As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStopStrategyFactory Interface Members
'@================================================================================

Private Function IStrategyFactory_CreateStrategy( _
                ByVal pTradingContext As TradingContext) As IPositionMgmtStrategy
Const ProcName As String = "IStrategyFactory_CreateStrategy"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStrategyFactory_Initialise( _
                ByVal pContext As InitialisationContext)
Const ProcName As String = "IStrategyFactory_Initialise"
On Error GoTo Err

Dim barsID As String
Dim longBarsID As String
Dim stopBarsID As String
Dim breakevenBarsID As String
Dim studyID As String
Dim params As Parameters

Set mTimeframe = pContext.AddTimeframe(GetTimePeriod(mTimePeriod, mTimePeriodUnit), barsID)

If mMinimumSwing <> 0 Then
    Set params = New cParameters
    params.setParam "valuetype", "$default"
    params.setParam "numberOfSwingHighs", 5
    params.setParam "numberOfSwingLows", 5
    params.setParam "minimumswing", mMinimumSwing
    params.setParam "includeimplicitswingpoints", "N"
    Set mSwing = pContext.AddStudy("studies2.CSwingStudy", _
                            studyID, _
                            barsID, _
                            params)
End If

If mLongTimePeriod <> 0 And _
    (mLongMAPeriods <> 0 Or mLongBollPeriods <> 0) _
Then
    Set mLongTimeframe = pContext.addPeriod(mLongTimePeriod, _
                                    mLongTimePeriodUnit, _
                                    longBarsID)
    
    If mLongMAPeriods <> 0 Then
        Set params = New cParameters
        params.setParam "valuetype", "lastprice"
        params.setParam "periods", mLongMAPeriods
        Set mLongMA = pContext.AddStudy("studies2.cSimpleMovingAverageStudy", _
                                studyID, _
                                longBarsID, _
                                params)
        mLongMA.numberOfHistoricValues = 2
    End If
    
    If mLongBollPeriods <> 0 Then
        Set params = New cParameters
        params.setParam "valuetype", "lastprice"
        params.setParam "periods", mLongBollPeriods
        params.setParam "deviations", 2
        Set mLongBoll = pContext.AddStudy("studies2.cBollingerBandStudy", _
                                studyID, _
                                longBarsID, _
                                params)
        mLongBoll.numberOfHistoricValues = 2
    End If
End If

If mConfirmedCountForBreakEven <> 0 Or _
    mConfirmedCountForSwing <> 0 Or _
    mConfirmedCountForTrailing <> 0 Or _
    mConfirmedCountForExit <> 0 _
Then
    Set params = New cParameters
    params.setParam "valuetype", "lastprice"
    params.setParam "shortperiods", 5
    params.setParam "shortmatype", "studies2.cExponentialMovingAverageStudy"
    params.setParam "longperiods", 34
    params.setParam "longmatype", "studies2.cExponentialMovingAverageStudy"
    params.setParam "smoothingperiods", 5
    params.setParam "smoothingmatype", "studies2.cSimpleMovingAverageStudy"
    Set mMACDStudy = pContext.AddStudy("studies2.cMACDStudy", _
                            studyID, _
                            barsID, _
                            params)
    mMACDStudy.numberOfHistoricValues = 10
End If

If mBackstopMABarLength > 0 And _
    mBackstopMAPeriods > 0 _
Then
    Set mStopTimeframe = pContext.AddTimeframe(GetTimePeriod(mBackstopMABarLength, _
                                                            mParams.GetParameterValue(ParamTimePeriodUnit)), _
                                                10)

    Set params = New Parameters
    params.SetParameterValue "Periods", mBackstopMAPeriods
    params.SetParameterValue "source", "HL/2"
    lInputValueNames(0) = "HL/2"
    Set mStopMA = pContext.AddStudy("ExponentialMovingAverage", _
                            mStopTimeframe, _
                            lInputValueNames, _
                            params, _
                            2)

    Set params = New Parameters
    params.SetParameterValue "Periods", 34
    params.SetParameterValue "Standard deviations", 2
    lInputValueNames(0) = "$default"
    Set mStopBoll = pContext.AddStudy("BollingerBands", _
                            mStopTimeframe, _
                            lInputValueNames, _
                            params, _
                            2)
End If
    
If mParams.GetParameterValue(ParamBreakEvenBarLength) > 0 Then
    Set mBreakevenTimeframe = pContext.AddTimeframe(GetTimePeriod(mParams.GetParameterValue(ParamBreakEvenBarLength), _
                                                                mParams.GetParameterValue(ParamTimePeriodUnit)), _
                                                    2)
End If




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

'@================================================================================
' Helper Functions
'@================================================================================




