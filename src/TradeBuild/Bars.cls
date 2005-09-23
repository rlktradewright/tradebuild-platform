VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

'================================================================================
' Events
'================================================================================

Event BarAdded(ByVal theBar As TradeBuild.Bar)
Event HistoricBarAdded(ByVal theBar As TradeBuild.Bar)

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

Private mBars As Collection
Private mBarType As TradeBuild.BarTypes
Private mBarLength As Long

Private mContract As Contract

Private mCurrentBarNumber As Long
Private mFirstSessionBarNumber As Long

Private mLastBarStartTime As Date

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Public Property Get barTime(Optional ByVal ref As Integer) As Variant
barTime = getTargetBar(ref).datetime
End Property

Friend Property Let barType(ByVal value As TradeBuild.BarTypes)
mBarType = value
End Property

Public Property Get closeValue(Optional ByVal ref As Integer) As Variant
closeValue = getTargetBar(ref).closeValue
End Property

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
mLastBarStartTime = mContract.currentSessionEndTime - mBarLength / 1440
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

Public Property Get currentBarNumber() As Long
currentBarNumber = mCurrentBarNumber
End Property

Public Property Get currentBarNumberInSession() As Long
currentBarNumberInSession = mCurrentBarNumber - mFirstSessionBarNumber + 1
End Property

Public Property Get down(Optional ByVal ref As Integer) As Boolean
down = (closeValue(ref) < closeValue(ref - 1))
End Property

Public Property Get highestCloseValue(ByVal periods As Integer, _
                                    Optional ByVal ref As Integer) As Double
Dim i                   As Integer
Dim high                As Double
Static lastPeriods      As Integer
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Integer

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If closeValue <= lastValue Then
        highestCloseValue = lastValue
        Exit Property
    Else
        highestCloseValue = closeValue
        lastValue = highestCloseValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    high = closeValue(ref - i)
    If high > highestCloseValue Then
        highestCloseValue = high
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = highestCloseValue
End Property

Public Property Get highestHighValue(ByVal periods As Integer, _
                                    Optional ByVal ref As Integer) As Variant
Dim i                   As Integer
Dim high                As Double
Static lastPeriods      As Integer
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Integer

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If highValue <= lastValue Then
        highestHighValue = lastValue
        Exit Property
    Else
        highestHighValue = highValue
        lastValue = highestHighValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    high = highValue(ref - i)
    If high > highestHighValue Then
        highestHighValue = high
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = highestHighValue
End Property

Public Property Get highValue(Optional ByVal ref As Integer) As Double
highValue = getTargetBar(ref).highValue
End Property

Public Property Get inside(Optional ByVal ref As Integer) As Variant
inside = (highValue(ref - 1) <= highValue(ref)) And (lowValue(ref - 1) >= lowValue(ref))
End Property

Public Property Get isFirstBarInSession(Optional ByVal ref As Integer) As Variant
isFirstBarInSession = Abs(getTargetBar(ref).datetime - mContract.currentSessionStartTime) < OneSecond
End Property

Public Property Get isLastBarInSession() As Boolean
isLastBarInSession = (Abs(barTime(0) - mLastBarStartTime) < OneSecond)
End Property

Public Property Get lowestCloseValue(ByVal periods As Integer, _
                                    Optional ByVal ref As Integer) As Double
Dim i As Integer
Dim Low As Currency
Static lastPeriods      As Integer
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Integer

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If closeValue >= lastValue Then
        lowestCloseValue = lastValue
        Exit Property
    Else
        lowestCloseValue = closeValue
        lastValue = lowestCloseValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    Low = closeValue(ref - i)
    If Low < lowestCloseValue Then
        lowestCloseValue = Low
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = lowestCloseValue
End Property

Public Property Get lowestLowValue(ByVal periods As Integer, _
                                    Optional ByVal ref As Integer) As Variant
Dim i As Integer
Dim Low As Currency
Static lastPeriods      As Integer
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Integer

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If lowValue >= lastValue Then
        lowestLowValue = lastValue
        Exit Property
    Else
        lowestLowValue = lowValue
        lastValue = lowestLowValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    Low = lowValue(ref - i)
    If Low < lowestLowValue Then
        lowestLowValue = Low
    End If

Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = lowestLowValue
End Property

Public Property Get lowValue(Optional ByVal ref As Integer) As Double
lowValue = getTargetBar(ref).lowValue
End Property

Public Property Get openValue(Optional ByVal ref As Integer) As Variant
openValue = getTargetBar(ref).openValue
End Property

Public Property Get outside(Optional ByVal ref As Integer) As Boolean
inside = (highValue(ref - 1) > highValue(ref)) And (lowValue(ref - 1) < lowValue(ref))
End Property

Public Property Get TickVolume(Optional ByVal ref As Integer) As Long
TickVolume = getTargetBar(ref).TickVolume
End Property

Public Property Get up(Optional ByVal ref As Integer) As Boolean
up = (closeValue(ref) > closeValue(ref - 1))
End Property

Public Property Get volume(Optional ByVal ref As Integer) As Long
volume = getTargetBar(ref).volume
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add(ByVal timestamp As Date) As Bar
Set add = New Bar
add.barType = mBarType
add.barLength = mBarLength
add.datetime = timestamp
mBars.add add, CalcBarKey(timestamp)
mCurrentBarNumber = mCurrentBarNumber + 1
If isFirstBarInSession Then mFirstSessionBarNumber = mCurrentBarNumber
RaiseEvent BarAdded(add)
End Function

Friend Function addHistoric( _
                            ByVal timestamp As Date, _
                            ByVal barType As TradeBuild.BarTypes, _
                            ByVal barLength As Long, _
                            ByVal openValue As Double, _
                            ByVal highValue As Double, _
                            ByVal lowValue As Double, _
                            ByVal closeValue As Double, _
                            ByVal volume As Long, _
                            ByVal TickVolume As Long, _
                            ByVal openInterest As Long) As Bar
Debug.Assert barType = mBarType
Debug.Assert barLength = mBarLength
Set addHistoric = New Bar
With addHistoric
    .barType = barType
    .barLength = barLength
    .datetime = timestamp
    .openValue = openValue
    .highValue = highValue
    .lowValue = lowValue
    .closeValue = closeValue
    .volume = volume
    .TickVolume = TickVolume
    .openInterest = openInterest
End With
mBars.add addHistoric, CalcBarKey(timestamp)
mCurrentBarNumber = mCurrentBarNumber + 1
If isFirstBarInSession Then mFirstSessionBarNumber = mCurrentBarNumber
RaiseEvent HistoricBarAdded(addHistoric)
End Function

Public Function Item(index As Variant) As Bar
Attribute Item.VB_UserMemId = 0
If VarType(index) = vbDate Then
        Set Item = mBars(CalcBarKey(CDate(index)))
Else
    Set Item = mBars(index)
End If
End Function

Public Sub Remove(index As Variant)
mBars.Remove index
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = mBars.[_NewEnum]
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function CalcBarKey(ByVal barTime As Date) As String
CalcBarKey = Format(barTime, "yyyymmddhhnn")
End Function

Private Function getTargetBar(ByVal ref As Long) As TradeBuild.Bar
If ref > mCurrentBarNumber Or _
    ref < -(mCurrentBarNumber - 1) _
Then
    err.Raise ErrorCodes.IllegalArgumentException, _
            "TradeBuild.Bars::getTargetBar", _
            "<ref> out of bounds"
End If

Set getTargetBar = mBars.Item(refToBarNumber(ref))
End Function

Private Function refToBarNumber(ByVal ref As Long) As Long
If ref > 0 Then
    refToBarNumber = ref
Else
    refToBarNumber = mCurrentBarNumber + ref
End If

End Function


