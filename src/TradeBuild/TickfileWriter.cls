VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type Tick
    timestamp       As Date
    TickType        As TradeBuild.TickTypes
    TickPrice       As Double
    TickSize        As Long
    Volume          As Long
    MDPosition      As Long
    MDMarketMaker   As String
    MDOperation     As Long
    MDSide          As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mWriter As TradeBuildSP.ITickfileWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Contract(ByVal value As Contract)
mWriter.Contract = value
End Property

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.ITickfileWriter)
Set mWriter = value
End Property

Friend Property Let ServiceProviderWriterReady(ByVal value As Boolean)
Dim item As Variant
Dim pendingWrite As Tick

mServiceProviderWriterReady = value

If mPendingWrites.Count <> 0 Then
    For Each item In mPendingWrites
        pendingWrite = item
        With pendingWrite
            Select Case .TickType
            Case bid
                mWriter.writeBid timestamp, .TickPrice, .TickSize
            Case ask
                mWriter.writeAsk .timestamp, .TickPrice, .TickSize
            Case Last
                mWriter.writeTrade .timestamp, .TickPrice, .TickSize
            Case High
                mWriter.writeHigh .timestamp, .TickPrice
            Case Low
                mWriter.writeLow .timestamp, .TickPrice
            Case PrevClose
                mWriter.writeClose .timestamp, .TickPrice
            Case Volume
                mWriter.writeVolume .timestamp, .TickSize
            Case OpenInterest

            Case marketDepth
                mWriter.writeMarketDepthData .timestamp, _
                            .MDPosition, _
                            .MDMarketMaker, _
                            .MDOperation, _
                            .MDSide, _
                            .TickPrice, _
                            .TickSize
            Case MarketDepthReset
                mWriter.writeMarketDepthReset .timestamp
            End Select
        End With
    Next
End If

Set mPendingWrites = Nothing
Set mPendingWrites = New Collection
End Property

'================================================================================
' Methods
'================================================================================

Public Sub closeOutputFile()
mWriter.closeOutputFile
End Sub

Friend Sub writeAsk(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.ask
        .timestamp = timestamp
        .TickPrice = price
        .TickSize = Size
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeAsk timestamp, price, Size
End If
End Sub


Friend Sub writeBid(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.bid
        .timestamp = timestamp
        .TickPrice = price
        .TickSize = Size
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeBid timestamp, price, Size
End If
End Sub

Friend Sub writeClose(ByVal timestamp As Date, _
                    ByVal price As Double)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.PrevClose
        .timestamp = timestamp
        .TickPrice = price
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeClose timestamp, price
End If
End Sub

Friend Sub writeHigh(ByVal timestamp As Date, _
                    ByVal price As Double)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.High
        .timestamp = timestamp
        .TickPrice = price
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeHigh timestamp, price
End If
End Sub

Friend Sub writeLow(ByVal timestamp As Date, _
                    ByVal price As Double)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.Low
        .timestamp = timestamp
        .TickPrice = price
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeLow timestamp, price
End If
End Sub

Friend Sub writeMarketDepthData( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As Long, _
                        ByVal side As Long, _
                        ByVal price As Currency, _
                        ByVal Size As Long)

Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.marketDepth
        .MDPosition = position
        .MDMarketMaker = marketMaker
        .MDOperation = operation
        .MDSide = side
        .timestamp = timestamp
        .TickPrice = price
        .TickSize = Size
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeMarketDepthData timestamp, _
            position, _
            marketMaker, _
            operation, _
            side, _
            price, _
            Size
End If
End Sub

Friend Sub writeMarketDepthReset( _
                        ByVal timestamp As Date)
                        
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.MarketDepthReset
        .timestamp = timestamp
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeMarketDepthReset timestamp
End If
End Sub
                        
Friend Sub writeTrade(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.Last
        .timestamp = timestamp
        .TickPrice = price
        .TickSize = Size
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeTrade timestamp, price, Size
End If
End Sub

Friend Sub writeVolume(ByVal timestamp As Date, _
                    ByVal Size As Long)
Dim pendingTick As Tick
If Not mServiceProviderWriterReady Then
    With pendingTick
        .TickType = TradeBuild.TickTypes.Volume
        .timestamp = timestamp
        .TickSize = Size
    End With
    mPendingWrites.add pendingTick
Else
    mWriter.writeVolume timestamp, Size
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================


