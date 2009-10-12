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
Implements LogListener

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

Private mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private mFromDate As Date
Private mToDate As Date
Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

Private mBarStudy As Study
Private mBars As Bars

Private mBar As Bar

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTB = TradeBuildAPI
If gLogToConsole Then GetLogger("").AddLogListener Me
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As TWUtilities30.CollectionChangeEvent)

outputBar
Set mBar = ev.affectedItem

End Sub

'@================================================================================
' LogListener Interface Members
'@================================================================================

Private Sub LogListener_Finish()

End Sub

Private Sub LogListener_Notify(ByVal Logrec As TWUtilities30.LogRecord)
Static formatter As LogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter
gCon.writeLineToConsole formatter.FormatRecord(Logrec)
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEvent)
gCon.writeErrorLine "Error " & ev.errorCode & _
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
    gCon.writeLineToConsole ev.eventMessage
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Ticker( _
                ByVal value As Ticker)
Set mTicker = value
End Property

Friend Property Let Timeframe( _
                ByVal value As Timeframe)
Set mBars = value.tradeStudy.Bars
mBars.AddCollectionChangeListener Me
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
outputBar
mBars.RemoveCollectionChangeListener Me
GetLogger("").RemoveLogListener Me
End Sub

Friend Sub StopData()
If Not mTicker Is Nothing Then mTicker.StopTicker
If Not mBars Is Nothing Then mBars.RemoveCollectionChangeListener Me
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputBar()
Dim sv As SValue

If mBar Is Nothing Then Exit Sub

gCon.writeString FormatTimestamp(mBar.TimeStamp, TimestampDateAndTimeISO8601)
gCon.writeString ","
gCon.writeString mTicker.FormatPrice(mBar.OpenValue)
gCon.writeString ","
gCon.writeString mTicker.FormatPrice(mBar.HighValue)
gCon.writeString ","
gCon.writeString mTicker.FormatPrice(mBar.LowValue)
gCon.writeString ","
gCon.writeString mTicker.FormatPrice(mBar.CloseValue)
gCon.writeString ","
gCon.writeString mBar.Volume
gCon.writeString ","
gCon.writeString mBar.TickVolume
gCon.writeString ","
gCon.writeLine mBar.OpenInterest
End Sub

