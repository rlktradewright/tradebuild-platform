VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataProcessor"
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

Private Const ModuleName                            As String = "TickDataProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mRS                                         As Recordset
Private mBarsBuilder                                As BarsBuilder
Private mSessionBuilder                             As SessionBuilder

Private mBarType                                    As BarTypes

Private mNotifiedFirstVolume                        As Boolean

Private mAccumulatedVolume                          As Long

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pRS As Recordset, _
                ByVal pBarsBuilder As BarsBuilder, _
                ByVal pSessionBuilder As SessionBuilder, _
                ByVal pBarType As BarTypes)
Set mRS = pRS
Set mBarsBuilder = pBarsBuilder
Set mSessionBuilder = pSessionBuilder
mBarType = pBarType

mNotifiedFirstVolume = False
mAccumulatedVolume = 0
End Sub

Friend Function ProcessNextTickData() As Boolean
Const ProcName As String = "ProcessTickData"
On Error GoTo Err

Dim lTdd As ITickDataDecoder
Set lTdd = getNextTickDataDecoder(mRS)

If lTdd Is Nothing Then
    ProcessNextTickData = False
    Exit Function
End If

Dim tick As GenericTick
Do While lTdd.GetNextTick(tick)
    processTick tick, mBarType, mBarsBuilder, mNotifiedFirstVolume
Loop

ProcessNextTickData = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ProcessTicks()
Const ProcName As String = "ProcessTicks"
On Error GoTo Err

Do While ProcessNextTickData
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getNextTickDataDecoder(ByVal pRS As Recordset) As ITickDataDecoder
Const ProcName As String = "getNextTickDataDecoder"
On Error GoTo Err

If mRS.EOF Then Exit Function

Dim data() As Byte
data = pRS("data")

Set getNextTickDataDecoder = CreateTickDecoder(pRS("datetime"), _
                                            pRS("tickSize"), _
                                            pRS("baseprice"), _
                                            data, _
                                            pRS("version"))

mRS.MoveNext

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processTick( _
                ByRef pTick As GenericTick, _
                ByVal pBarType As BarTypes, _
                ByVal pBarsBuilder As BarsBuilder, _
                ByRef pNotifiedFirstVolume As Boolean)
Const ProcName As String = "processTick"
On Error GoTo Err

setSessionTime pTick.timestamp

Dim val As SValue
val.timestamp = pTick.timestamp
Select Case pTick.TickType
Case TickTypeBid
    If pBarType = BarTypeBid Then
        val.value = pTick.price
        pBarsBuilder.NotifyValue val
    End If
Case TickTypeAsk
    If pBarType = BarTypeAsk Then
        val.value = pTick.price
        pBarsBuilder.NotifyValue val
    End If
Case TickTypeTrade
    If pBarType = BarTypeTrade Then
        val.value = pTick.price
        pBarsBuilder.NotifyValue val
    End If
Case TickTypeVolume
    If pBarType = BarTypeTrade Then
        If Not pNotifiedFirstVolume Then
            val.value = 1
            pBarsBuilder.NotifyVolume val
            pNotifiedFirstVolume = True
            val.value = pTick.size - 1
            pBarsBuilder.NotifyVolume val
        Else
            val.value = pTick.size
            pBarsBuilder.NotifyVolume val
        End If
    End If
Case TickTypeOpenInterest
    If pBarType = BarTypeTrade Then
        val.value = pTick.size
        pBarsBuilder.NotifyOpenInterest val
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionTime( _
                ByVal timestamp As Date)
Const ProcName As String = "setSessionTime"
On Error GoTo Err

Dim lSessEv As SessionEventData
lSessEv = mSessionBuilder.SetSessionCurrentTime(timestamp)
If lSessEv.ChangeType = SessionChangeStart Then
    mAccumulatedVolume = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



