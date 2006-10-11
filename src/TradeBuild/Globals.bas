Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const S_OK = 0
Public Const NoValidID As Long = -1
Public Const InitialMaxTickers As Long = 100&

Public Const DefaultStudyValue As String = "$default"

Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const OneSecond As Double = 1.15740740740741E-05
Public Const OneMicroSecond As Double = 1.15740740740741E-11

Public Const MultiTaskingTimeQuantumMillisecs As Long = 20

Public Const StrOrderTypeNone As String = ""
Public Const StrOrderTypeMarket As String = "Market"
Public Const StrOrderTypeMarketClose As String = "Market on Close"
Public Const StrOrderTypeLimit As String = "Limit"
Public Const StrOrderTypeLimitClose As String = "Limit on Close"
Public Const StrOrderTypePegMarket As String = "Peg to Market"
Public Const StrOrderTypeStop As String = "Stop"
Public Const StrOrderTypeStopLimit As String = "Stop Limit"
Public Const StrOrderTypeTrail As String = "Trailing Stop"
Public Const StrOrderTypeRelative As String = "Relative"
Public Const StrOrderTypeVWAP As String = "VWAP"
Public Const StrOrderTypeMarketToLimit As String = "Market to Limit"
Public Const StrOrderTypeQuote As String = "Quote"
Public Const StrOrderTypeAutoStop As String = "Auto Stop"
Public Const StrOrderTypeAutoLimit As String = "Auto Limit"
Public Const StrOrderTypeAdjust As String = "Adjust"
Public Const StrOrderTypeAlert As String = "Alert"
Public Const StrOrderTypeLimitIfTouched As String = "Limit if Touched"
Public Const StrOrderTypeMarketIfTouched As String = "Market if Touched"
Public Const StrOrderTypeTrailLimit As String = "Trail Limit"
Public Const StrOrderTypeMarketWithProtection As String = "Market with Protection"
Public Const StrOrderTypeMarketOnOpen As String = "Market on Open"
Public Const StrOrderTypeLimitOnOpen As String = "Limit on Open"
Public Const StrOrderTypePeggedToPrimary As String = "Pegged to Primary"

Public Const StrOrderActionBuy As String = "Buy"
Public Const StrOrderActionSell As String = "Sell"

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

''================================================================================
'' Global object references
''================================================================================
'
'Public gServiceProviders As ServiceProviders
'Public gTradeBuildAPI As TradeBuildAPI
'Public gListeners As InfoListeners
'Public gTaskManager As taskManager

'================================================================================
' External function declarations
'================================================================================

Public Declare Function CoCreateGuid Lib "OLE32.dll" (pGUID As GUIDStruct) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)
                            
Public Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)

Public Declare Function StringFromGUID2 Lib "OLE32.dll" ( _
                            ByRef rguid As GUIDStruct, _
                            ByRef lpsz As GUIDString, _
                            ByVal cchMax As Long) As Integer

'================================================================================
' Variables
'================================================================================


'================================================================================
' Procedures
'================================================================================

Public Sub gAddItemToCombo( _
                ByVal combo As ComboBox, _
                ByVal itemText As String, _
                ByVal itemData As Long)
combo.AddItem itemText
combo.itemData(combo.ListCount - 1) = itemData
End Sub

Public Function gCurrentTime() As Date
gCurrentTime = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
End Function

'/**
' Converts a member of the EntryOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pEntryOrderType the EntryOrderTypes value to be converted
' @ see
'
'*/
Public Function gEntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As TradeBuild.EntryOrderTypes) As TradeBuild.OrderTypes
Select Case pEntryOrderType
Case EntryOrderTypeMarket
    gEntryOrderTypeToOrderType = OrderTypeMarket
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToOrderType = OrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToOrderType = OrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToOrderType = OrderTypeMarketIfTouched
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToOrderType = OrderTypeMarketToLimit
Case EntryOrderTypeBid
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeAsk
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLast
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimit
    gEntryOrderTypeToOrderType = OrderTypeLimit
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToOrderType = OrderTypeLimitOnOpen
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToOrderType = OrderTypeLimitOnClose
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToOrderType = OrderTypeLimitIfTouched
Case EntryOrderTypeStop
    gEntryOrderTypeToOrderType = OrderTypeStop
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToOrderType = OrderTypeStopLimit
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild.Module1::gEntryOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gEntryOrderTypeToString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    gEntryOrderTypeToString = "Market"
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToString = "Market on open"
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToString = "Market on close"
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToString = "Market if touched"
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToString = "Market to limit"
Case EntryOrderTypeBid
    gEntryOrderTypeToString = "Bid price"
Case EntryOrderTypeAsk
    gEntryOrderTypeToString = "Ask price"
Case EntryOrderTypeLast
    gEntryOrderTypeToString = "Last trade price"
Case EntryOrderTypeLimit
    gEntryOrderTypeToString = "Limit"
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToString = "Limit on open"
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToString = "Limit on close"
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToString = "Limit if touched"
Case EntryOrderTypeStop
    gEntryOrderTypeToString = "Stop"
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToString = "Stop limit"
End Select
End Function

Public Function gEntryOrderTypeToShortString(ByVal value As EntryOrderTypes) As String
Select Case value
Case EntryOrderTypeMarket
    gEntryOrderTypeToShortString = "MKT"
Case EntryOrderTypeMarketOnOpen
    gEntryOrderTypeToShortString = "MOO"
Case EntryOrderTypeMarketOnClose
    gEntryOrderTypeToShortString = "MOC"
Case EntryOrderTypeMarketIfTouched
    gEntryOrderTypeToShortString = "MIT"
Case EntryOrderTypeMarketToLimit
    gEntryOrderTypeToShortString = "MTL"
Case EntryOrderTypeBid
    gEntryOrderTypeToShortString = "BID"
Case EntryOrderTypeAsk
    gEntryOrderTypeToShortString = "ASK"
Case EntryOrderTypeLast
    gEntryOrderTypeToShortString = "LAST"
Case EntryOrderTypeLimit
    gEntryOrderTypeToShortString = "LMT"
Case EntryOrderTypeLimitOnOpen
    gEntryOrderTypeToShortString = "LOO"
Case EntryOrderTypeLimitOnClose
    gEntryOrderTypeToShortString = "LOC"
Case EntryOrderTypeLimitIfTouched
    gEntryOrderTypeToShortString = "LIT"
Case EntryOrderTypeStop
    gEntryOrderTypeToShortString = "STP"
Case EntryOrderTypeStopLimit
    gEntryOrderTypeToShortString = "STPLMT"
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

Public Function gGenerateGUID() As GUIDStruct
Dim lReturn As Long

lReturn = CoCreateGuid(gGenerateGUID)

If (lReturn <> S_OK) Then
    err.Raise ErrorCodes.ErrRuntimeException, _
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

Public Function gGUIDToString(ByRef pGUID As GUIDStruct) As String
Dim GUIDString As GUIDString
Dim iChars As Integer

' convert binary GUID to string form
iChars = StringFromGUID2(pGUID, GUIDString, Len(GUIDString))
' convert string to ANSI
gGUIDToString = StrConv(GUIDString.GUIDProper, vbFromUnicode)
End Function

Public Function gLegOpenCloseFromString(ByVal value As String) As TradeBuild.LegOpenClose
Select Case UCase$(value)
Case ""
    gLegOpenCloseFromString = LegUnknownPos
Case "SAME"
    gLegOpenCloseFromString = LegSamePos
Case "OPEN"
    gLegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    gLegOpenCloseFromString = LegClosePos
End Select
End Function

Public Function gLegOpenCloseToString(ByVal value As TradeBuild.LegOpenClose) As String
Select Case value
Case LegSamePos
    gLegOpenCloseToString = "SAME"
Case LegOpenPos
    gLegOpenCloseToString = "OPEN"
Case LegClosePos
    gLegOpenCloseToString = "CLOSE"
End Select
End Function

Public Function gNewWeakReference(ByVal target As Object) As WeakReference
Set gNewWeakReference = New WeakReference
gNewWeakReference.initialise target
End Function

Public Function gOptionRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    gOptionRightFromString = OptNone
Case "CALL"
    gOptionRightFromString = OptCall
Case "PUT"
    gOptionRightFromString = OptPut
End Select
End Function

Public Function gOptionRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    gOptionRightToString = ""
Case OptCall
    gOptionRightToString = "Call"
Case OptPut
    gOptionRightToString = "Put"
End Select
End Function

Public Function gOrderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case StrOrderActionBuy
    gOrderActionFromString = ActionBuy
Case StrOrderActionSell
    gOrderActionFromString = ActionSell
End Select
End Function

Public Function gOrderActionToString(ByVal value As OrderActions) As String
Select Case value
Case ActionBuy
    gOrderActionToString = StrOrderActionBuy
Case ActionSell
    gOrderActionToString = StrOrderActionSell
End Select
End Function

Public Function gOrderStatusToString(ByVal value As OrderStatuses) As String
Select Case UCase$(value)
Case OrderStatusCreated
    gOrderStatusToString = "Created"
Case OrderStatusPendingSubmit
    gOrderStatusToString = "Pending Submit"
Case OrderStatusPreSubmitted
    gOrderStatusToString = "Presubmitted"
Case OrderStatusSubmitted
    gOrderStatusToString = "Submitted"
Case OrderStatusCancelling
    gOrderStatusToString = "Cancelling"
Case OrderStatusCancelled
    gOrderStatusToString = "Cancelled"
Case OrderStatusFilled
    gOrderStatusToString = "Filled"
End Select
End Function

Public Function gOrderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
Select Case value
Case StopTriggerMethods.StopTriggerDefault
    gOrderStopTriggerMethodToString = "Default"
Case StopTriggerMethods.StopTriggerDoubleBidAsk
    gOrderStopTriggerMethodToString = "Double bid/ask"
Case StopTriggerMethods.StopTriggerDoubleLast
    gOrderStopTriggerMethodToString = "Double last"
Case StopTriggerMethods.StopTriggerLast
    gOrderStopTriggerMethodToString = "Last"
End Select
End Function

Public Function gOrderTIFToString(ByVal value As OrderTifs) As String
Select Case value
Case TIFDay
    gOrderTIFToString = "DAY"
Case TIFGoodTillCancelled
    gOrderTIFToString = "GTC"
Case TIFImmediateOrCancel
    gOrderTIFToString = "IOC"
End Select
End Function

Public Function gOrderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case OrderTypeNone
    gOrderTypeToString = StrOrderTypeNone
Case OrderTypeMarket
    gOrderTypeToString = StrOrderTypeMarket
Case OrderTypeMarketOnClose
    gOrderTypeToString = StrOrderTypeMarketClose
Case OrderTypeLimit
    gOrderTypeToString = StrOrderTypeLimit
Case OrderTypeLimitOnClose
    gOrderTypeToString = StrOrderTypeLimitClose
Case OrderTypePeggedToMarket
    gOrderTypeToString = StrOrderTypePegMarket
Case OrderTypeStop
    gOrderTypeToString = StrOrderTypeStop
Case OrderTypeStopLimit
    gOrderTypeToString = StrOrderTypeStopLimit
Case OrderTypeTrail
    gOrderTypeToString = StrOrderTypeTrail
Case OrderTypeRelative
    gOrderTypeToString = StrOrderTypeRelative
Case OrderTypeVWAP
    gOrderTypeToString = StrOrderTypeVWAP
Case OrderTypeMarketToLimit
    gOrderTypeToString = StrOrderTypeMarketToLimit
Case OrderTypeQuote
    gOrderTypeToString = StrOrderTypeQuote
Case OrderTypeAdjust
    gOrderTypeToString = StrOrderTypeAdjust
Case OrderTypeAlert
    gOrderTypeToString = StrOrderTypeAlert
Case OrderTypeLimitIfTouched
    gOrderTypeToString = StrOrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
    gOrderTypeToString = StrOrderTypeMarketIfTouched
Case OrderTypeTrailLimit
    gOrderTypeToString = StrOrderTypeTrailLimit
Case OrderTypeMarketWithProtection
    gOrderTypeToString = StrOrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
    gOrderTypeToString = StrOrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
    gOrderTypeToString = StrOrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
    gOrderTypeToString = StrOrderTypePeggedToPrimary
End Select

End Function

Public Function gSecTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STOCK", "STK"
    gSecTypeFromString = SecTypeStock
Case "FUTURE", "FUT"
    gSecTypeFromString = SecTypeFuture
Case "OPTION", "OPT"
    gSecTypeFromString = SecTypeOption
Case "FUTURES OPTION", "FOP"
    gSecTypeFromString = SecTypeFuturesOption
Case "CASH"
    gSecTypeFromString = SecTypeCash
Case "BAG"
    gSecTypeFromString = SecTypeBag
Case "INDEX", "IND"
    gSecTypeFromString = SecTypeIndex
End Select
End Function

Public Function gSecTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToString = "Stock"
Case SecTypeFuture
    gSecTypeToString = "Future"
Case SecTypeOption
    gSecTypeToString = "Option"
Case SecTypeFuturesOption
    gSecTypeToString = "Futures Option"
Case SecTypeCash
    gSecTypeToString = "Cash"
Case SecTypeBag
    gSecTypeToString = "Bag"
Case SecTypeIndex
    gSecTypeToString = "Index"
End Select
End Function

Public Function gSecTypeToShortString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToShortString = "STK"
Case SecTypeFuture
    gSecTypeToShortString = "FUT"
Case SecTypeOption
    gSecTypeToShortString = "OPT"
Case SecTypeFuturesOption
    gSecTypeToShortString = "FOP"
Case SecTypeCash
    gSecTypeToShortString = "CASH"
Case SecTypeBag
    gSecTypeToShortString = "BAG"
Case SecTypeIndex
    gSecTypeToShortString = "IND"
End Select
End Function

Public Sub gSortObjects(data() As SortEntryStruct, _
                            Low As Long, _
                            Hi As Long)
  
  Dim lTmpLow As Long
  Dim lTmpHi As Long
  Dim lTmpMid As Long
  Dim vTempVal As SortEntryStruct
  Dim vTmpHold As SortEntryStruct
  
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

'/**
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
' @ see
'
'*/
Public Function gStopOrderTypeToOrderType( _
                ByVal pStopOrderType As TradeBuild.StopOrderTypes) As TradeBuild.OrderTypes
Select Case pStopOrderType
Case StopOrderTypeNone
    gStopOrderTypeToOrderType = OrderTypeNone
Case StopOrderTypeStop
    gStopOrderTypeToOrderType = OrderTypeStop
Case StopOrderTypeStopLimit
    gStopOrderTypeToOrderType = OrderTypeLimit
Case StopOrderTypeBid
    gStopOrderTypeToOrderType = OrderTypeLimit
Case StopOrderTypeAsk
    gStopOrderTypeToOrderType = OrderTypeLimit
Case StopOrderTypeLast
    gStopOrderTypeToOrderType = OrderTypeLimit
Case StopOrderTypeAuto
    gStopOrderTypeToOrderType = OrderTypeAutoLimit
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild.Module1::gStopOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gStopOrderTypeToString(ByVal value As StopOrderTypes)
Select Case value
Case StopOrderTypeNone
    gStopOrderTypeToString = "None"
Case StopOrderTypeStop
    gStopOrderTypeToString = "Stop"
Case StopOrderTypeStopLimit
    gStopOrderTypeToString = "Stop limit"
Case StopOrderTypeBid
    gStopOrderTypeToString = "Bid price"
Case StopOrderTypeAsk
    gStopOrderTypeToString = "Ask price"
Case StopOrderTypeLast
    gStopOrderTypeToString = "Last trade price"
Case StopOrderTypeAuto
    gStopOrderTypeToString = "Auto"
End Select
End Function

'/**
' Converts a member of the TargetOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pTargetOrderType the TargetOrderTypes value to be converted
' @ see
'
'*/
Public Function gTargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TradeBuild.TargetOrderTypes) As TradeBuild.OrderTypes
Select Case pTargetOrderType
Case TargetOrderTypeNone
    gTargetOrderTypeToOrderType = OrderTypeNone
Case TargetOrderTypeLimit
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLimitIfTouched
    gTargetOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TargetOrderTypeMarketIfTouched
    gTargetOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TargetOrderTypeBid
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAsk
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeLast
    gTargetOrderTypeToOrderType = OrderTypeLimit
Case TargetOrderTypeAuto
    gTargetOrderTypeToOrderType = OrderTypeAutoLimit
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild.Module1::gTargetOrderTypeToOrderType", _
                "Invalid entry type"
End Select
End Function

Public Function gTargetOrderTypeToString(ByVal value As TargetOrderTypes)
Select Case value
Case TargetOrderTypeNone
    gTargetOrderTypeToString = "None"
Case TargetOrderTypeLimit
    gTargetOrderTypeToString = "Limit"
Case TargetOrderTypeMarketIfTouched
    gTargetOrderTypeToString = "Market if touched"
Case TargetOrderTypeBid
    gTargetOrderTypeToString = "Bid price"
Case TargetOrderTypeAsk
    gTargetOrderTypeToString = "Ask price"
Case TargetOrderTypeLast
    gTargetOrderTypeToString = "Last trade price"
Case TargetOrderTypeAuto
    gTargetOrderTypeToString = "Auto"
End Select
End Function

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

