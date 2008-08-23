VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements CollectionChangeListener

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

Private Const ProjectName                   As String = "gbd"
Private Const ModuleName                    As String = "Processor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTB As TradeBuildAPI
Attribute mTB.VB_VarHelpID = -1

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mTimeframe As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mPriceFormatString As String

Private mFromDate As Date
Private mToDate As Date
Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

Private mBarStudy As Study
Private mBars As bars

Private mBar As Bar

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTB = TradeBuildAPI
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As TWUtilities30.CollectionChangeEvent)

outputBar
Set mBar = ev.affectedItem

End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_Notification(ev As TWUtilities30.NotificationEvent)
gCon.writeLineToConsole "Notification " & ev.eventCode & _
                        ": " & ev.eventMessage
End Sub

Private Sub mTicker_stateChange(ev As TWUtilities30.StateChangeEvent)
If ev.state = TickerStates.TickerStateReady Then
    loaddata
End If
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
outputBar
Set gProcessor = Nothing
gCon.writeLineToConsole "Completed"
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEvent)
gCon.writeLineToConsole "Error " & ev.errorCode & _
                        ": " & ev.errorMessage
End Sub

Private Sub mTB_Notification(ev As TWUtilities30.NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Dim spe As ServiceProviderError
    Set spe = mTB.GetServiceProviderError
    gCon.writeLineToConsole "Service provider error " & spe.errorCode & _
                            ": " & spe.message & _
                            " (" & spe.serviceProviderName & ")"
Else
    gCon.writeLineToConsole "Notification " & ev.eventCode & _
                            ": " & ev.eventMessage
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub startData( _
                ByVal spec As ContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal number As Long, _
                ByVal barLength As Long, _
                ByVal barUnits As TimePeriodUnits, _
                ByVal sessionOnly As Boolean)
                
Set mTicker = TradeBuildAPI.Tickers.Add(TickerOptions.TickerOptUseExchangeTimeZone)
mTicker.loadTicker spec

mFromDate = fromDate
mToDate = toDate
mNumber = number
mBarLength = barLength
mBarUnits = barUnits
mSessionOnly = sessionOnly
                
End Sub

Friend Sub stopData()
mTicker.stopTicker
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub loaddata()

mPriceFormatString = mTicker.Contract.priceFormatString

Set mTimeframe = mTicker.Timeframes.addHistorical(GetTimePeriod(mBarLength, mBarUnits), _
                                        "", _
                                        mNumber, _
                                        mFromDate, _
                                        mToDate, _
                                        Not mSessionOnly)
Set mBarStudy = mTimeframe.tradeStudy
Set mBars = mTimeframe.tradeStudy.bars
mBars.addCollectionChangeListener Me

End Sub

Private Sub outputBar()
Dim sv As SValue

If mBar Is Nothing Then Exit Sub

gCon.writeString FormatTimestamp(mBar.TimeStamp, TimestampDateAndTimeISO8601)
gCon.writeString ","
gCon.writeString Format(mBar.openValue, mPriceFormatString)
gCon.writeString ","
gCon.writeString Format(mBar.highValue, mPriceFormatString)
gCon.writeString ","
gCon.writeString Format(mBar.lowValue, mPriceFormatString)
gCon.writeString ","
gCon.writeString Format(mBar.closeValue, mPriceFormatString)
gCon.writeString ","
gCon.writeString mBar.volume
gCon.writeString ","
gCon.writeString mBar.tickVolume
gCon.writeString ","
gCon.writeLine mBar.openInterest
End Sub

