Attribute VB_Name = "Module1"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const S_OK = 0
Public Const NoValidID As Long = -1
Public Const InitialMaxTickers As Long = 100&

Public Const MinDouble As Double = -1.79769313486231E+308
Public Const MaxDouble As Double = 1.79769313486231E+308

Public Const OneSecond As Double = 1.15740740740741E-05
Public Const OneMicroSecond As Double = 1.15740740740741E-11

Public Const MultiTaskingTimeQuantumMillisecs As Long = 20

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Public Type GUIDString
    StartBrace  As String * 2
    GUIDProper  As String * 72
    EndBrace    As String * 2
    ZeroByte    As String * 1
End Type

'================================================================================
' Global object references
'================================================================================

Public gServiceProviders As ServiceProviders
Public gTradeBuildAPI As TradeBuildAPI
Public gTickers As Tickers
Public gListeners As listeners
Public gOrderSimulator As AdvancedOrderSimulator

Public gNextOrderID As Long
Public gAllOrders As Collection


'================================================================================
' External function declarations
'================================================================================

Public Declare Function CoCreateGuid Lib "OLE32.dll" (pGUID As GUID) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            Source As Any, _
                            ByVal length As Long)
                            
Public Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            Source As Any, _
                            ByVal length As Long)

Public Declare Function StringFromGUID2 Lib "OLE32.dll" ( _
                            ByRef rguid As GUID, _
                            ByRef lpsz As GUIDString, _
                            ByVal cchMax As Long) As Integer

'================================================================================
' Variables
'================================================================================


'================================================================================
' Procedures
'================================================================================

Public Function gCurrentTime() As Date
gCurrentTime = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
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

Public Function gGenerateGUID() As GUID
Dim lReturn As Long

lReturn = CoCreateGuid(gGenerateGUID)

If (lReturn <> S_OK) Then
    err.Raise ErrorCodes.CantCreateGUID, _
                "TWUtilities.Utilities::GenerateGUID", _
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

Public Function gGUIDToString(ByRef pGUID As GUID) As String
Dim GUIDString As GUIDString
Dim iChars As Integer

' convert binary GUID to string form
iChars = StringFromGUID2(pGUID, GUIDString, Len(GUIDString))
' convert string to ANSI
gGUIDToString = StrConv(GUIDString.GUIDProper, vbFromUnicode)
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
     Do While (data(lTmpLow).key < vTempVal.key And lTmpLow < Hi)
           lTmpLow = lTmpLow + 1
     Loop
      
     ' Now, we will process the high end.  Loop as long as
     ' the data in the temporary holding area is less
     ' than the array data element and the temporary high
     ' value is greater than the minimum number of array
     ' elements.
     Do While (vTempVal.key < data(lTmpHi).key And lTmpHi > Low)
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

Public Function gToBytes(inString As String) As Byte()
Dim i As Long
Dim outAr() As Byte

ReDim outAr(Len(inString) / 2) As Byte
For i = 0 To (Len(inString) / 2) - 1
    outAr(i) = val("&h" & Mid$(inString, 2 * i + 1, 2))
Next
gToBytes = outAr
End Function

Public Function gToHex(inAr() As Byte) As String
Dim s As String
Dim i As Long

For i = 0 To UBound(inAr)
    If inAr(i) < 16 Then
        s = s & "0" & Hex$(inAr(i))
    Else
        s = s & Hex$(inAr(i))
    End If
Next
gToHex = s
End Function

Public Function LegOpenCloseFromString(ByVal value As String) As TradeBuild.LegOpenClose
Select Case UCase$(value)
Case ""
    LegOpenCloseFromString = LegUnknownPos
Case "SAME"
    LegOpenCloseFromString = LegSamePos
Case "OPEN"
    LegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    LegOpenCloseFromString = LegClosePos
End Select
End Function

Public Function LegOpenCloseToString(ByVal value As TradeBuild.LegOpenClose) As String
Select Case value
Case LegSamePos
    LegOpenCloseToString = "SAME"
Case LegOpenPos
    LegOpenCloseToString = "OPEN"
Case LegClosePos
    LegOpenCloseToString = "CLOSE"
End Select
End Function

Public Function OptRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    OptRightFromString = OptNone
Case "CALL"
    OptRightFromString = OptCall
Case "PUT"
    OptRightFromString = OptPut
End Select
End Function

Public Function OptRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    OptRightToString = ""
Case OptCall
    OptRightToString = "Call"
Case OptPut
    OptRightToString = "Put"
End Select
End Function

Public Function orderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case "BUY"
    orderActionFromString = OrderActions.ActionBuy
Case "SELL"
    orderActionFromString = OrderActions.ActionSell
End Select
End Function

Public Function orderActionToString(ByVal value As OrderActions) As String
Select Case value
Case OrderActions.ActionBuy
    orderActionToString = "BUY"
Case OrderActions.ActionSell
    orderActionToString = "SELL"
End Select
End Function

Public Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STK"
    secTypeFromString = SecTypeStock
Case "FUT"
    secTypeFromString = SecTypeFuture
Case "OPT"
    secTypeFromString = SecTypeOption
Case "FOP"
    secTypeFromString = SecTypeFuturesOption
Case "CASH"
    secTypeFromString = SecTypeCash
Case "BAG"
    secTypeFromString = SecTypeBag
Case "IND"
    secTypeFromString = SecTypeIndex
End Select
End Function

Public Function secTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    secTypeToString = "STK"
Case SecTypeFuture
    secTypeToString = "FUT"
Case SecTypeOption
    secTypeToString = "OPT"
Case SecTypeFuturesOption
    secTypeToString = "FOP"
Case SecTypeCash
    secTypeToString = "CASH"
Case SecTypeBag
    secTypeToString = "BAG"
Case SecTypeIndex
    secTypeToString = "IND"
End Select
End Function

Public Sub showRecord(rs As Recordset)
Dim fld As Field
Dim s As String
For Each fld In rs.fields
    s = s & fld.Name & "=" & fld.value & vbCrLf
Next
MsgBox s
End Sub

Public Function gTickfileSpecifierToString(TickfileSpec As TradeBuild.TickfileSpecifier) As String
If TickfileSpec.filename <> "" Then
    gTickfileSpecifierToString = TickfileSpec.filename
Else
    gTickfileSpecifierToString = "Contract: " & _
                                Replace(TickfileSpec.Contract.specifier.ToString, vbCrLf, "; ") & _
                            ": From: " & FormatDateTime(TickfileSpec.From, vbGeneralDate) & _
                            " To: " & FormatDateTime(TickfileSpec.To, vbGeneralDate)
End If
End Function
