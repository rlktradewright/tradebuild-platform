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

Private Sub CollectionChangeListener_Change(ev As TWUtilities30.CollectionChangeEventData)
Const ProcName As String = "CollectionChangeListener_Change"
On Error GoTo Err

outputBar
Set mBar = ev.affectedItem

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' LogListener Interface Members
'@================================================================================

Private Sub LogListener_Finish()

End Sub

Private Sub LogListener_Notify(ByVal Logrec As TWUtilities30.LogRecord)
Static formatter As LogFormatter
Const ProcName As String = "LogListener_Notify"
On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter
gCon.WriteLineToConsole formatter.FormatRecord(Logrec)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEventData)
Const ProcName As String = "mTB_Error"
On Error GoTo Err

gCon.WriteErrorLine "Error " & ev.errorCode & _
                        ": " & ev.ErrorMessage

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTB_Notification(ev As TWUtilities30.NotificationEventData)
Const ProcName As String = "mTB_Notification"
On Error GoTo Err

If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Dim spe As ServiceProviderError
    Set spe = mTB.GetServiceProviderError
    gCon.WriteLineToConsole "Service provider error " & spe.errorCode & _
                            ": " & spe.Message & _
                            " (" & spe.ServiceProviderName & ")"
Else
    gCon.WriteLineToConsole ev.eventMessage
End If

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
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
Const ProcName As String = "Timeframe"
On Error GoTo Err

Set mBars = value.TradeStudy.Bars
mBars.AddCollectionChangeListener Me

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

outputBar
mBars.RemoveCollectionChangeListener Me
GetLogger("").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub StopData()
Const ProcName As String = "StopData"
On Error GoTo Err

If Not mTicker Is Nothing Then mTicker.StopTicker
If Not mBars Is Nothing Then mBars.RemoveCollectionChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub outputBar()
Dim sv As SValue

Const ProcName As String = "outputBar"
On Error GoTo Err

If mBar Is Nothing Then Exit Sub

gCon.WriteString FormatTimestamp(mBar.TimeStamp, TimestampDateAndTimeISO8601)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.OpenValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.HighValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.LowValue)
gCon.WriteString ","
gCon.WriteString mTicker.FormatPrice(mBar.CloseValue)
gCon.WriteString ","
gCon.WriteString mBar.Volume
gCon.WriteString ","
gCon.WriteString mBar.TickVolume
gCon.WriteString ","
gCon.WriteLine mBar.OpenInterest

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

