Attribute VB_Name = "gMain"
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

Public Type TickInfo
    Tick                As GenericTick
    ReceivedOffset      As Double
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "gMain"

Public Const TickfilePath                           As String = "D:\projects\tradeBuild-platform\src\TickfileUtils\TickfileUtilsUnitTests\Tickfiles"

Public Const OneMillisec                            As Double = 1# / 86400000
Public Const OneSecond                              As Double = 1# / 86400

' This should really have value 5, but I've set it to 15 because of the
' problem with timer resolution
Private Const TickOffsetTolerance                   As Double = 15#

'@================================================================================
' Member variables
'@================================================================================

Public ReceivedTicks()                              As TickInfo
Public NumberOfReceivedTicks                        As Long

Private mElapsedTimer                               As New ElapsedTimer

Public TicksA(4)                                    As GenericTick
Public TicksB(6)                                    As GenericTick

Private mMergedAandB(11)                            As TickInfo

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

Public Sub gAddReceivedTick(ByRef pTick As GenericTick)
Dim lOffset As Double
NumberOfReceivedTicks = NumberOfReceivedTicks + 1

If NumberOfReceivedTicks = 1 Then
    lOffset = 0#
    mElapsedTimer.StartTiming
Else
    lOffset = mElapsedTimer.ElapsedTimeMicroseconds / 1000#
End If

If NumberOfReceivedTicks - 1 > UBound(ReceivedTicks) Then ReDim Preserve ReceivedTicks(2 * (UBound(ReceivedTicks) + 1) - 1) As TickInfo
ReceivedTicks(NumberOfReceivedTicks - 1).Tick = pTick
ReceivedTicks(NumberOfReceivedTicks - 1).ReceivedOffset = lOffset
End Sub

Public Sub gCheckReceivedTicks( _
                ByRef pSourceTicks() As GenericTick, _
                Optional ByVal pExpectedCount As Long, _
                Optional ByVal pStartPauseAt As Long, _
                Optional ByVal pPauseFor As Long)
Dim lExpectedCount As Long
lExpectedCount = IIf(pExpectedCount <> 0, pExpectedCount, UBound(pSourceTicks) + 1)
Assert.IsTrue NumberOfReceivedTicks = lExpectedCount, "Received " & NumberOfReceivedTicks & " ticks instead of " & CStr(lExpectedCount)

Dim lBaseTime As Date
lBaseTime = pSourceTicks(0).TimeStamp

Dim lPauseTime As Date
lPauseTime = lBaseTime + pStartPauseAt * OneMillisec

Dim lDiscrepancy As Double
Dim lMaxDiscrepancy As Double
Dim lAdjustment As Long
Dim i As Long
For i = 0 To lExpectedCount - 1
    Dim lTicksDiffer As Boolean
    If pSourceTicks(i).MarketMaker <> ReceivedTicks(i).Tick.MarketMaker Then lTicksDiffer = True
    If pSourceTicks(i).Operation <> ReceivedTicks(i).Tick.Operation Then lTicksDiffer = True
    If pSourceTicks(i).Position <> ReceivedTicks(i).Tick.Position Then lTicksDiffer = True
    If pSourceTicks(i).Price <> ReceivedTicks(i).Tick.Price Then lTicksDiffer = True
    If pSourceTicks(i).Side <> ReceivedTicks(i).Tick.Side Then lTicksDiffer = True
    If pSourceTicks(i).Size <> ReceivedTicks(i).Tick.Size Then lTicksDiffer = True
    If pSourceTicks(i).TickType <> ReceivedTicks(i).Tick.TickType Then lTicksDiffer = True
    If pSourceTicks(i).TimeStamp <> ReceivedTicks(i).Tick.TimeStamp Then lTicksDiffer = True
    
    If pSourceTicks(i).TimeStamp >= lPauseTime Then lAdjustment = pPauseFor
    
    Assert.IsFalse lTicksDiffer, "Received tick " & i + 1 & " is not the same as the source tick"
    
    lDiscrepancy = ((pSourceTicks(i).TimeStamp - lBaseTime) * 86400000 + lAdjustment) - ReceivedTicks(i).ReceivedOffset
    LogMessage "Tick " & i + 1 & ": offset discrepancy is " & Format(lDiscrepancy, "0.00")
    If Abs(lDiscrepancy) > lMaxDiscrepancy Then lMaxDiscrepancy = Abs(lDiscrepancy)
'    Assert.AreEqualFloats (pSourceTicks(i).TimeStamp - lBaseTime) * 86400000 + lAdjustment, _
'                        ReceivedTicks(i).ReceivedOffset, _
'                        TickOffsetTolerance, _
'                        "Tick " & i + 1 & " received at offset " & ReceivedTicks(i).ReceivedOffset
Next

Assert.Less lMaxDiscrepancy, TickOffsetTolerance, "A tick's delivery time was outside the tolerance"
End Sub

Public Sub gCheckReceivedTicksConcurrent()
Assert.IsTrue NumberOfReceivedTicks = UBound(mMergedAandB) + 1, "Received " & NumberOfReceivedTicks & " ticks instead of " & CStr(UBound(mMergedAandB) + 1)

Dim lDiscrepancy As Double
Dim lMaxDiscrepancy As Double
Dim i As Long

For i = 0 To NumberOfReceivedTicks - 1
    Dim lTicksDiffer As Boolean
    If mMergedAandB(i).Tick.MarketMaker <> ReceivedTicks(i).Tick.MarketMaker Then lTicksDiffer = True
    If mMergedAandB(i).Tick.Operation <> ReceivedTicks(i).Tick.Operation Then lTicksDiffer = True
    If mMergedAandB(i).Tick.Position <> ReceivedTicks(i).Tick.Position Then lTicksDiffer = True
    If mMergedAandB(i).Tick.Price <> ReceivedTicks(i).Tick.Price Then lTicksDiffer = True
    If mMergedAandB(i).Tick.Side <> ReceivedTicks(i).Tick.Side Then lTicksDiffer = True
    If mMergedAandB(i).Tick.Size <> ReceivedTicks(i).Tick.Size Then lTicksDiffer = True
    If mMergedAandB(i).Tick.TickType <> ReceivedTicks(i).Tick.TickType Then lTicksDiffer = True
    If mMergedAandB(i).Tick.TimeStamp <> ReceivedTicks(i).Tick.TimeStamp Then lTicksDiffer = True
    
    Assert.IsFalse lTicksDiffer, "Received tick " & i + 1 & " is not the same as the source tick"
    
    lDiscrepancy = mMergedAandB(i).ReceivedOffset - ReceivedTicks(i).ReceivedOffset
    LogMessage "Tick " & i + 1 & ": offset discrepancy is " & Format(lDiscrepancy, "0.00")
    If Abs(lDiscrepancy) > lMaxDiscrepancy Then lMaxDiscrepancy = Abs(lDiscrepancy)
'    Assert.AreEqualFloats mMergedAandB(i).ReceivedOffset, _
'                        ReceivedTicks(i).ReceivedOffset, _
'                        TickOffsetTolerance, _
'                        "Tick " & i + 1 & " received at offset " & ReceivedTicks(i).ReceivedOffset
Next

Assert.Less lMaxDiscrepancy, TickOffsetTolerance, "A tick's delivery time was outside the tolerance"
End Sub

Public Sub gCheckReceivedTicksTwoLots( _
                ByRef pSourceA() As GenericTick, _
                ByRef pSourceB() As GenericTick, _
                Optional ByVal pExpectedCountA As Long, _
                Optional ByVal pExpectedCountB As Long)
Dim lExpectedCountA As Long
lExpectedCountA = IIf(pExpectedCountA <> 0, pExpectedCountA, UBound(pSourceA) + 1)
Dim lExpectedCountB As Long
lExpectedCountB = IIf(pExpectedCountB <> 0, pExpectedCountB, UBound(pSourceB) + 1)
Assert.IsTrue NumberOfReceivedTicks = lExpectedCountA + lExpectedCountB, "Received " & NumberOfReceivedTicks & " ticks instead of " & CStr(lExpectedCountA + lExpectedCountB)

Dim lDiscrepancy As Double
Dim lMaxDiscrepancy As Double
Dim lTicksDiffer As Boolean
Dim i As Long
Dim j As Long

For i = 0 To lExpectedCountA - 1
    If pSourceA(i).MarketMaker <> ReceivedTicks(j).Tick.MarketMaker Then lTicksDiffer = True
    If pSourceA(i).Operation <> ReceivedTicks(j).Tick.Operation Then lTicksDiffer = True
    If pSourceA(i).Position <> ReceivedTicks(j).Tick.Position Then lTicksDiffer = True
    If pSourceA(i).Price <> ReceivedTicks(j).Tick.Price Then lTicksDiffer = True
    If pSourceA(i).Side <> ReceivedTicks(j).Tick.Side Then lTicksDiffer = True
    If pSourceA(i).Size <> ReceivedTicks(j).Tick.Size Then lTicksDiffer = True
    If pSourceA(i).TickType <> ReceivedTicks(j).Tick.TickType Then lTicksDiffer = True
    If pSourceA(i).TimeStamp <> ReceivedTicks(j).Tick.TimeStamp Then lTicksDiffer = True
    
    Assert.IsFalse lTicksDiffer, "Received tick " & j + 1 & " is not the same as the source tick"
    
    lDiscrepancy = ((pSourceA(i).TimeStamp - pSourceA(0).TimeStamp) * 86400000) - ReceivedTicks(j).ReceivedOffset
    LogMessage "Tick " & i + 1 & ": offset discrepancy is " & Format(lDiscrepancy, "0.00")
    If Abs(lDiscrepancy) > lMaxDiscrepancy Then lMaxDiscrepancy = Abs(lDiscrepancy)
'    Assert.AreEqualFloats (pSourceA(i).TimeStamp - pSourceA(0).TimeStamp) * 86400000, _
'                        ReceivedTicks(j).ReceivedOffset, _
'                        TickOffsetTolerance, _
'                        "TickA " & j + 1 & " received at offset " & ReceivedTicks(j).ReceivedOffset
    j = j + 1
Next

j = lExpectedCountA
Dim lBaseOffset As Double
lBaseOffset = ReceivedTicks(lExpectedCountA).ReceivedOffset

For i = 0 To lExpectedCountB - 1
    If pSourceB(i).MarketMaker <> ReceivedTicks(j).Tick.MarketMaker Then lTicksDiffer = True
    If pSourceB(i).Operation <> ReceivedTicks(j).Tick.Operation Then lTicksDiffer = True
    If pSourceB(i).Position <> ReceivedTicks(j).Tick.Position Then lTicksDiffer = True
    If pSourceB(i).Price <> ReceivedTicks(j).Tick.Price Then lTicksDiffer = True
    If pSourceB(i).Side <> ReceivedTicks(j).Tick.Side Then lTicksDiffer = True
    If pSourceB(i).Size <> ReceivedTicks(j).Tick.Size Then lTicksDiffer = True
    If pSourceB(i).TickType <> ReceivedTicks(j).Tick.TickType Then lTicksDiffer = True
    If pSourceB(i).TimeStamp <> ReceivedTicks(j).Tick.TimeStamp Then lTicksDiffer = True
    
    Assert.IsFalse lTicksDiffer, "Received tick " & j + 1 & " is not the same as the source tick"
    
    lDiscrepancy = ((pSourceB(i).TimeStamp - pSourceB(0).TimeStamp) * 86400000) - (ReceivedTicks(j).ReceivedOffset - lBaseOffset)
    LogMessage "Tick " & i + 1 & ": offset discrepancy is " & Format(lDiscrepancy, "0.00")
    If Abs(lDiscrepancy) > lMaxDiscrepancy Then lMaxDiscrepancy = Abs(lDiscrepancy)
'    Assert.AreEqualFloats (pSourceB(i).TimeStamp - pSourceB(0).TimeStamp) * 86400000, _
'                        ReceivedTicks(j).ReceivedOffset - lBaseOffset, _
'                        TickOffsetTolerance, _
'                        "TickB " & j + 1 & " received at offset " & (ReceivedTicks(j).ReceivedOffset - lBaseOffset)
    j = j + 1
Next

Assert.Less lMaxDiscrepancy, TickOffsetTolerance, "A tick's delivery time was outside the tolerance"
End Sub

Public Function gCreateContract(ByVal pContractSpecifier As IContractSpecifier) As IContract
Select Case pContractSpecifier.localSymbol
Case "ZH13"
    Set gCreateContract = gCreateContractFromLocalSymbol("ZH13")
Case "ZM13"
    Set gCreateContract = gCreateContractFromLocalSymbol("ZM13")
Case "ESM3"
    Set gCreateContract = gCreateContractFromLocalSymbol("ESM3")
Case Else
    If pContractSpecifier.Symbol = "Z" And pContractSpecifier.SecType = SecTypeFuture And pContractSpecifier.expiry = "200306" Then
        Set gCreateContract = gCreateContractFromLocalSymbol("ZM03")
    Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Localname not known"
    End If
End Select
End Function

Public Function gCreateContractFromLocalSymbol(ByVal pLocalSymbol As String) As IContract
Dim lContract As IContract

Select Case pLocalSymbol
Case "ESM3"
    Set lContract = createESContract("ESM3", "20130621")
Case "ZM03"
    Set lContract = createZContract("ZM03", "20030620")
Case "ZZ2"
    Set lContract = createZContract("ZZ2", "20121221")
Case "ZH13"
    Set lContract = createZContract("ZH13", "20130315")
Case "ZM13"
    Set lContract = createZContract("ZM13", "20130621")
Case "ZU3"
    Set lContract = createZContract("ZU13", "20130920")
Case "ZZ13"
    Set lContract = createZContract("ZZ13", "20131220")
Case "ZU14"
    Set lContract = createZContract("ZU14", "20140919")
Case "IBM"
    Set lContract = createStockContract("IBM")
Case "MSFT"
    Set lContract = createStockContract("MSFT")
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Localname not known"
End Select
    
Set gCreateContractFromLocalSymbol = lContract
End Function

Public Function gCreateTick( _
                ByVal pTimestamp As Date, _
                ByVal pTickType As TickTypes, _
                Optional ByVal pPrice As Double, _
                Optional ByVal pSize As BoxedDecimal, _
                Optional ByVal pSide As DOMSides, _
                Optional ByVal pPosition As Long, _
                Optional ByVal pOperation As DOMOperations, _
                Optional ByVal pMarketMaker As String) As GenericTick
gCreateTick.MarketMaker = pMarketMaker
gCreateTick.Operation = pOperation
gCreateTick.Position = pPosition
gCreateTick.Price = pPrice
gCreateTick.Side = pSide
Set gCreateTick.Size = pSize
gCreateTick.TickType = pTickType
gCreateTick.TimeStamp = pTimestamp
End Function

Private Sub Main()
setupTicksA
setupTicksB
mergeAandB
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createESContract(ByVal localSymbol As String, ByVal expiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, "ES", , "CME", SecTypeFuture, "USD", expiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("16:30")
lContractBuilder.TickSize = 0.25
lContractBuilder.ExpiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
lContractBuilder.TimezoneName = "Central Standard Time"
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createESContract = lContractBuilder.Contract
End Function

Private Function createStockContract(ByVal localSymbol As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, localSymbol, , "SMART", SecTypeStock, "USD", "")

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("09:30")
lContractBuilder.TickSize = 0.01
lContractBuilder.TimezoneName = "Eastern Standard Time"
Set createStockContract = lContractBuilder.Contract
End Function

Private Function createZContract(ByVal localSymbol As String, ByVal expiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, "Z", , "ICEEU", SecTypeFuture, "GBP", expiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("17:30")
lContractBuilder.SessionStartTime = CDate("08:00")
lContractBuilder.TickSize = 0.5
lContractBuilder.TimezoneName = "GMT Standard Time"
lContractBuilder.ExpiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createZContract = lContractBuilder.Contract
End Function

Private Sub mergeAandB()
mMergedAandB(0).Tick = TicksB(0)
mMergedAandB(0).ReceivedOffset = 0#

mMergedAandB(1).Tick = TicksA(0)
mMergedAandB(1).ReceivedOffset = 2000#

mMergedAandB(2).Tick = TicksA(1)
mMergedAandB(2).ReceivedOffset = 3000#

mMergedAandB(3).Tick = TicksB(1)
mMergedAandB(3).ReceivedOffset = 3007#

mMergedAandB(4).Tick = TicksB(2)
mMergedAandB(4).ReceivedOffset = 4000#

mMergedAandB(5).Tick = TicksB(3)
mMergedAandB(5).ReceivedOffset = 4001#

mMergedAandB(6).Tick = TicksA(2)
mMergedAandB(6).ReceivedOffset = 4120#

mMergedAandB(7).Tick = TicksA(3)
mMergedAandB(7).ReceivedOffset = 4255#

mMergedAandB(8).Tick = TicksB(4)
mMergedAandB(8).ReceivedOffset = 5000#

mMergedAandB(9).Tick = TicksB(5)
mMergedAandB(9).ReceivedOffset = 5010#

mMergedAandB(10).Tick = TicksB(6)
mMergedAandB(10).ReceivedOffset = 5011#

mMergedAandB(11).Tick = TicksA(4)
mMergedAandB(11).ReceivedOffset = 6000#

End Sub

Private Sub setupTicksA()
TicksA(0) = gCreateTick(CDate("21/02/2013 08:14:25"), TickTypeAsk, 6720#, CreateBoxedDecimal(3))
TicksA(1) = gCreateTick(CDate("21/02/2013 08:14:26"), TickTypeBid, 6720.5, CreateBoxedDecimal(5))
TicksA(2) = gCreateTick(CDate("21/02/2013 08:14:27") + 120 * OneMillisec, TickTypeTrade, 6720#, CreateBoxedDecimal(1))
TicksA(3) = gCreateTick(CDate("21/02/2013 08:14:27") + 255 * OneMillisec, TickTypeVolume, , CreateBoxedDecimal(7625))
TicksA(4) = gCreateTick(CDate("21/02/2013 08:14:29"), TickTypeClosePrice, 6708#)
End Sub

Private Sub setupTicksB()
TicksB(0) = gCreateTick(CDate("21/02/2013 02:14:23"), TickTypeAsk, 1550.25, CreateBoxedDecimal(5))
TicksB(1) = gCreateTick(CDate("21/02/2013 02:14:26") + 7 * OneMillisec, TickTypeBid, 1550#, CreateBoxedDecimal(12))
TicksB(2) = gCreateTick(CDate("21/02/2013 02:14:27"), TickTypeTrade, 1550#, CreateBoxedDecimal(1))
TicksB(3) = gCreateTick(CDate("21/02/2013 02:14:27") + 1 * OneMillisec, TickTypeVolume, , CreateBoxedDecimal(362455))
TicksB(4) = gCreateTick(CDate("21/02/2013 02:14:28"), TickTypeClosePrice, 1543.75)
TicksB(5) = gCreateTick(CDate("21/02/2013 02:14:28") + 10 * OneMillisec, TickTypeHighPrice, 1560.5)
TicksB(6) = gCreateTick(CDate("21/02/2013 02:14:28") + 11 * OneMillisec, TickTypeLowPrice, 1548#)
End Sub


