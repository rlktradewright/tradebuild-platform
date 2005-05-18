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

'================================================================================
' Member variables
'================================================================================

Private mWriter As TradeBuildSP.ITickfileWriter

'================================================================================
' Class Event Handlers
'================================================================================

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

'================================================================================
' Methods
'================================================================================

Public Sub closeOutputFile()
mWriter.closeOutputFile
End Sub

Friend Sub writeAsk(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
mWriter.writeAsk timestamp, price, Size
End Sub


Friend Sub writeBid(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
mWriter.writeBid timestamp, price, Size
End Sub

Friend Sub writeClose(ByVal timestamp As Date, _
                    ByVal price As Double)
mWriter.writeClose timestamp, price
End Sub

Friend Sub writeHigh(ByVal timestamp As Date, _
                    ByVal price As Double)
mWriter.writeHigh timestamp, price
End Sub

Friend Sub writeLow(ByVal timestamp As Date, _
                    ByVal price As Double)
mWriter.writeLow timestamp, price
End Sub

Friend Sub writeMarketDepthData( _
                        ByVal timestamp As Date, _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As Long, _
                        ByVal side As Long, _
                        ByVal price As Currency, _
                        ByVal Size As Long)

mWriter.writeMarketDepthData timestamp, _
            position, _
            marketMaker, _
            operation, _
            side, _
            price, _
            Size
End Sub

Friend Sub writeMarketDepthReset( _
                        ByVal timestamp As Date)
                        
mWriter.writeMarketDepthReset timestamp
End Sub
                        
Friend Sub writeTrade(ByVal timestamp As Date, _
                    ByVal price As Double, _
                    ByVal Size As Long)
mWriter.writeTrade timestamp, price, Size
End Sub

Friend Sub writeVolume(ByVal timestamp As Date, _
                    ByVal Size As Long)
mWriter.writeVolume timestamp, Size
End Sub


'================================================================================
' Helper Functions
'================================================================================


