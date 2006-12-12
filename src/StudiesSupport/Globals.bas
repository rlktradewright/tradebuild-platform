Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const DefaultStudyValueName As String = "$DEFAULT"

Public Const OneMicroSecond As Double = 1.15740740740741E-11

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mStudyServiceProviders      As New studyServiceProviders

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get studyServiceProviders() As studyServiceProviders
Set studyServiceProviders = mStudyServiceProviders
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Public Function gBarStartTime( _
                ByVal timestamp As Date, _
                ByVal barLength As Long, _
                ByVal timeUnits As TradeBuildSP.TimePeriodUnits, _
                Optional ByVal sessionStartTime As Date) As Date
' minutes from midnight to start of sesssion
Dim sessionOffset              As Long
Dim theDate As Long
Dim theTime As Double
Dim theTimeMins As Long
Dim theTimeSecs As Long

sessionOffset = Int(1440 * (sessionStartTime - Int(sessionStartTime)))

theDate = Int(CDbl(timestamp))
' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
theTime = CDbl(timestamp + OneMicroSecond) - theDate

Select Case timeUnits
Case TimePeriodUnits.Second
    theTimeSecs = Fix(theTime * 86400) ' seconds since midnight
    gBarStartTime = theDate + _
                ((barLength) * Int((theTimeSecs - sessionOffset * 60) / barLength) + _
                    sessionOffset * 60) / 86400
Case TimePeriodUnits.Minute
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    gBarStartTime = theDate + _
                (barLength * Int((theTimeMins - sessionOffset) / barLength) + _
                    sessionOffset) / 1440
Case TimePeriodUnits.Hour
    theTimeMins = Fix(theTime * 1440) ' minutes since midnight
    gBarStartTime = theDate + _
                (60 * barLength * Int((theTimeMins - sessionOffset) / (60 * barLength)) + _
                    sessionOffset) / 1440
Case TimePeriodUnits.Day
    If theTime >= sessionStartTime Then
        gBarStartTime = theDate
    Else
        gBarStartTime = theDate - 1
    End If
Case TimePeriodUnits.Week
    gBarStartTime = theDate - DatePart("w", theDate, vbSunday) + 1
Case TimePeriodUnits.Month
    gBarStartTime = DateSerial(Year(theDate), Month(theDate), 1)
Case TimePeriodUnits.LunarMonth

Case TimePeriodUnits.Year
    gBarStartTime = DateSerial(Year(theDate), 1, 1)
End Select

End Function

Public Function gFormatTimestamp(ByVal timestamp As Date, _
                                Optional ByVal formatOption As TimestampFormats = TimestampDateAndTime, _
                                Optional ByVal formatString As String = "yyyymmddhhnnss") As String
Dim timestampDays As Long
Dim timestampSecs As Double
Dim timestampAsDate As Date
Dim milliseconds As Long

timestampDays = Int(timestamp)
timestampSecs = Int((timestamp - Int(timestamp)) * 86400) / 86400#
timestampAsDate = CDate(CDbl(timestampDays) + timestampSecs)
milliseconds = CLng((timestamp - timestampAsDate) * 86400# * 1000#)

If milliseconds >= 1000& Then
    milliseconds = milliseconds - 1000&
    timestampSecs = timestampSecs + (1# / 86400#)
    timestampAsDate = CDate(CDbl(timestampDays) + timestampSecs)
End If

Select Case formatOption
Case TimestampFormats.TimestampTimeOnly
    gFormatTimestamp = Format(timestampAsDate, "hhnnss") & "." & _
                        Format(milliseconds, "000")
Case TimestampFormats.TimestampDateOnly
    gFormatTimestamp = Format(timestampAsDate, "yyyymmdd")
Case TimestampFormats.TimestampDateAndTime
    gFormatTimestamp = Format(timestampAsDate, "yyyymmddhhnnss") & "." & _
                        Format(milliseconds, "000")
Case TimestampFormats.TimestampCustom
    gFormatTimestamp = Format(timestampAsDate, formatString) & "." & _
                        Format(milliseconds, "000")
End Select
End Function

Public Function gGenerateGUID() As TradeBuildSP.GUID
Dim lReturn As Long
Dim lGUID As GUIDStruct

lReturn = CoCreateGuid(lGUID)

gGenerateGUID.data(0) = lGUID.data(0)
gGenerateGUID.data(1) = lGUID.data(1)
gGenerateGUID.data(2) = lGUID.data(2)
gGenerateGUID.data(3) = lGUID.data(3)

If (lReturn <> S_OK) Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
                "TWUtilities.Utilities::gGenerateGUID", _
                "Can't create GUID"
End If

End Function

Public Function gGenerateGUIDString() As String
gGenerateGUIDString = gGUIDToString(gGenerateGUID)
End Function

Public Function gGenerateID() As Long
Randomize
gGenerateID = Fix(Rnd() * 1999999999 + 1)
End Function

Public Function gGenerateIDString() As String
gGenerateIDString = Hex(gGenerateID)
End Function

Public Function gGUIDToString(ByRef pGUID As TradeBuildSP.GUID) As String
Dim lGUID As GUIDStruct
Dim lGUIDString As GUIDString
Dim iChars As Integer

lGUID.data(0) = pGUID.data(0)
lGUID.data(1) = pGUID.data(1)
lGUID.data(2) = pGUID.data(2)
lGUID.data(3) = pGUID.data(3)
' convert binary GUID to string form
iChars = StringFromGUID2(lGUID, lGUIDString, Len(lGUIDString))
' convert string to ANSI
gGUIDToString = StrConv(lGUIDString.GUIDProper, vbFromUnicode)
End Function

Public Sub gSortObjects(data() As SortEntry, _
                            Low As Long, _
                            Hi As Long)
  
  Dim lTmpLow As Long
  Dim lTmpHi As Long
  Dim lTmpMid As Long
  Dim vTempVal As SortEntry
  Dim vTmpHold As SortEntry
  
  lTmpLow = Low
  lTmpHi = Hi
  
' ---------------------------------------------------------
' Leave if there is nothing to sort
' ---------------------------------------------------------
  If Hi <= Low Then Exit Sub

' ---------------------------------------------------------
' Find the middle to start comparing values
' ---------------------------------------------------------
  lTmpMid = (Low + Hi) \ 2
      
' ---------------------------------------------------------
' Move the item in the middle of the array to the
' temporary holding area as a point of reference while
' sorting.  This will change each time we make a recursive
' call to this routine.
' ---------------------------------------------------------
  vTempVal = data(lTmpMid)
      
' ---------------------------------------------------------
' Loop until we eventually meet in the middle
' ---------------------------------------------------------
  Do While (lTmpLow <= lTmpHi)
 
     ' Always process the low end first.  Loop as long as
     ' the array data element is less than the data in
     ' the temporary holding area and the temporary low
     ' value is less than the maximum number of array
     ' elements.
     Do While (data(lTmpLow).Key < vTempVal.Key And lTmpLow < Hi)
           lTmpLow = lTmpLow + 1
     Loop
      
     ' Now, we will process the high end.  Loop as long as
     ' the data in the temporary holding area is less
     ' than the array data element and the temporary high
     ' value is greater than the minimum number of array
     ' elements.
     Do While (vTempVal.Key < data(lTmpHi).Key And lTmpHi > Low)
           lTmpHi = lTmpHi - 1
     Loop
            
     ' if the temp low end is less than or equal
     ' to the temp high end, then swap places
     If (lTmpLow <= lTmpHi) Then
         vTmpHold = data(lTmpLow)          ' Move the Low value to Temp Hold
         data(lTmpLow) = data(lTmpHi)     ' Move the high value to the low
         data(lTmpHi) = vTmpHold           ' move the Temp Hod to the High
         lTmpLow = lTmpLow + 1              ' Increment the temp low counter
         lTmpHi = lTmpHi - 1                ' Dcrement the temp high counter
     End If
     
  Loop
          
' ---------------------------------------------------------
' If the minimum number of elements in the array is
' less than the temp high end, then make a recursive
' call to this routine.  I always sort the low end
' of the array first.
' ---------------------------------------------------------
  If (Low < lTmpHi) Then
      gSortObjects data, Low, lTmpHi
  End If
          
' ---------------------------------------------------------
' If the temp low end is less than the maximum number
' of elements in the array, then make a recursive call
' to this routine.  The high end is always sorted last.
' ---------------------------------------------------------
  If (lTmpLow < Hi) Then
       gSortObjects data, lTmpLow, Hi
  End If
  
End Sub

Public Function gHexStringToBytes(inString As String) As Byte()
Dim i As Long
Dim outAr() As Byte

ReDim outAr(Len(inString) / 2 - 1) As Byte
For i = 0 To (Len(inString) / 2) - 1
    outAr(i) = Val("&h" & Mid$(inString, 2 * i + 1, 2))
Next
gHexStringToBytes = outAr
End Function

Public Function gBytesToHexString(inAr() As Byte) As String
Dim i As Long

gBytesToHexString = String(2 * (UBound(inAr) + 1), " ")

For i = 0 To UBound(inAr)
    If inAr(i) < 16 Then
        Mid(gBytesToHexString, 2 * i + 1, 1) = "0"
        Mid(gBytesToHexString, 2 * i + 2, 1) = Hex$(inAr(i))
    Else
        Mid(gBytesToHexString, 2 * i + 1, 2) = Hex$(inAr(i))
    End If
Next

End Function



