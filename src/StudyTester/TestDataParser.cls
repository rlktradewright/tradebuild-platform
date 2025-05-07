VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDataParser"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum TestDataFileColumns
    TimeStamp
    OpenValue
    HighValue
    LowValue
    CloseValue
    Volume
    OpenInterest
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TestDataParser"

'@================================================================================
' Member variables
'@================================================================================

Private mBarsBuilder As BarsBuilder
Private mMinimumPriceTick As Double

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

Public Property Get Bars() As Bars
Set Bars = mBarsBuilder.Bars
End Property

Public Property Get MinimumPriceTick() As Double
MinimumPriceTick = mMinimumPriceTick
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ParseData( _
                ByVal pFilename As String)
Const ProcName As String = "ParseData"
On Error GoTo Err

Dim lSessionBuilder As SessionBuilder
Set lSessionBuilder = CreateSessionBuilder(0#, 0#, Nothing)

Set mBarsBuilder = CreateBarsBuilder(determineTimeframe(getTextStream(pFilename)), lSessionBuilder.Session)

Dim lAnalyzer As DataAnalyzer
Set lAnalyzer = New DataAnalyzer

buildBars getTextStream(pFilename), lAnalyzer

mMinimumPriceTick = lAnalyzer.Analyze

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub buildBar(ByRef pTokens() As String, ByVal pAnalyzer As DataAnalyzer)
Const ProcName As String = "buildBar"
On Error GoTo Err

Dim lTimestamp As Date
Dim lBar As New BarUtils27.Bar

lBar.Initialise GetTimestamp(pTokens), _
                getValue(pTokens, OpenValue, pAnalyzer), _
                getValue(pTokens, HighValue, pAnalyzer), _
                getValue(pTokens, LowValue, pAnalyzer), _
                getValue(pTokens, CloseValue, pAnalyzer), _
                getVolume(pTokens), _
                getOpenInterest(pTokens)

mBarsBuilder.AddThisBar lBar
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub buildBars(ByVal pTs As TextStream, ByVal pAnalyzer As DataAnalyzer)
Dim lTokens() As String

Const ProcName As String = "buildBars"
On Error GoTo Err

Do While readLine(pTs, lTokens)
    buildBar lTokens, pAnalyzer
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function determineTimeframe(ByVal pTs As TextStream) As TimePeriod
Dim lTimestamp1 As Date
Dim lTimestamp2 As Date
Dim lTokens() As String

Const ProcName As String = "determineTimeframe"
On Error GoTo Err

Assert readLine(pTs, lTokens), "Not enough data to determine timeframe"
lTimestamp1 = GetTimestamp(lTokens)

Assert readLine(pTs, lTokens), "Not enough data to determine timeframe"

lTimestamp2 = GetTimestamp(lTokens)
        
If DateDiff("s", lTimestamp1, lTimestamp2) < 60 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("s", lTimestamp1, lTimestamp2), TimePeriodSecond)
ElseIf DateDiff("n", lTimestamp1, lTimestamp2) < 60 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("n", lTimestamp1, lTimestamp2), TimePeriodMinute)
ElseIf DateDiff("h", lTimestamp1, lTimestamp2) < 24 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("h", lTimestamp1, lTimestamp2), TimePeriodHour)
ElseIf DateDiff("d", lTimestamp1, lTimestamp2) < 5 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("d", lTimestamp1, lTimestamp2), TimePeriodDay)
ElseIf DateDiff("ww", lTimestamp1, lTimestamp2) < 5 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("ww", lTimestamp1, lTimestamp2), TimePeriodWeek)
ElseIf DateDiff("m", lTimestamp1, lTimestamp2) < 12 Then
    Set determineTimeframe = GetTimePeriod(DateDiff("m", lTimestamp1, lTimestamp2), TimePeriodMonth)
Else
    Set determineTimeframe = GetTimePeriod(DateDiff("yyyy", lTimestamp1, lTimestamp2), TimePeriodYear)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOpenInterest(ByRef pTokens() As String) As Long
Const ProcName As String = "getOpenInterest"
On Error GoTo Err

Dim oi As String
oi = getToken(pTokens, TestDataFileColumns.OpenInterest)
If oi <> "" Then getOpenInterest = CLng(oi)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTextStream(ByVal pFilename As String) As TextStream
Dim fso As New FileSystemObject
Const ProcName As String = "getTextStream"
On Error GoTo Err

Set getTextStream = fso.OpenTextFile(pFilename, ForReading)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetTimestamp(ByRef pTokens() As String) As Date
Const ProcName As String = "GetTimestamp"
On Error GoTo Err

GetTimestamp = CDate(pTokens(TestDataFileColumns.TimeStamp))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getToken(ByRef pTokens() As String, ByVal pIndex As TestDataFileColumns) As String
On Error Resume Next
getToken = pTokens(pIndex)
End Function

Private Function getValue( _
                ByRef pTokens() As String, _
                ByVal pIndex As TestDataFileColumns, _
                ByVal pAnalyzer As DataAnalyzer) As Double
Const ProcName As String = "getValue"
On Error GoTo Err

getValue = CDbl(getToken(pTokens, pIndex))

pAnalyzer.AddDataValue getValue

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getVolume(ByRef pTokens() As String) As BoxedDecimal
Const ProcName As String = "getVolume"
On Error GoTo Err

Dim vol As String
vol = getToken(pTokens, TestDataFileColumns.Volume)
If vol <> "" Then getVolume = CreateBoxedDecimal(vol)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function readLine( _
                ByVal pTs As TextStream, _
                ByRef pTokens() As String) As Boolean
Dim rec As String
    
Const ProcName As String = "readLine"
On Error GoTo Err

Do While Not pTs.AtEndOfStream
    rec = pTs.readLine
    If rec <> "" And Left$(rec, 2) <> "//" Then
        pTokens = Split(rec, ",")
        readLine = True
        Exit Function
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Function



