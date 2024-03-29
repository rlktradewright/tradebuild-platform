VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickDataEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright 2008-2011 Richard L King
'
' This file is part of TradeBuild Tick Utilities Package.
'
' TradeBuild Tick Utilities Package is free software: you can redistribute it
' and/or modify it under the terms of the GNU General Public License as
' published by the Free Software Foundation, either version 3 of the License,
' or (at your option) any later version.
'
' TradeBuild Tick Utilities Package is distributed in the hope that it will
' be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with TradeBuild Tick Utilities Package.  If not, see
' <http://www.gnu.org/licenses/>.
 
Option Explicit

''
' This interface is implemented by classes that provide the ability to encode
' tick data into encoded segments.
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

Private Const ModuleName                    As String = "ITickDataEncoder"

'@================================================================================
' Member variables
'@================================================================================

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

''
' Returns the first price recorded in the data encoded by this encoder.
'
' This value must be stored by the application in addition to the encoded data itself,
' because it has to be passed to the decoder to enable the data to be correctly decoded.
' @return The first price recorded in the data encoded by this encoder. Note that
'   if there are no ticks with prices during the encoding period, the value returned is
'   the maximum positive value that can be held in a 64-bit floating point number, ie
'   (2-2<sup>-52</sup>)&middot;2<sup>1023</sup>.
'@/
Public Property Get BasePrice() As Double
End Property

''
' Returns the encoded data generated by this encoder.
' @return The encoded data generated by this encoder.
'@/
Public Property Get EncodedData() As Byte()
End Property

''
' Returns the length in bytes of the encoded data generated by this encoder.
' @return The length in bytes of the encoded data generated by this encoder.
'@/
Public Property Get EncodedDataLength() As Long
End Property

''
' Returns a URN uniquely identifying the format of the encoded data.
'
' This value must be stored by the application in addition to the encoded data itself, because it has to be
' passed to the decoder to enable the data to be correctly decoded.
'
' Due to the length of this identifier, it may be desirable to design the application such that this value
' is not repeated for every stored segment of encoded data. This design may need to allow for individual segments
' to be updated in the future by a later version of the encoding that uses a different format and hence supplies
' a different identifier.
' @return A URN uniquely identifying the format of the encoded data.
'@/
Public Property Get EncodingFormatIdentifier() As String
End Property

''
' Returns the period start time supplied when instantiating this decoder using {@link com.tradewright.tradebuild.tickutils.TickUtils#getTickEncoder getTickEncoder}.
' @return The period start time supplied when instantiating this decoder.
'@/
Public Property Get PeriodStartTime() As Date
End Property

''
' Returns the minimum tick size supplied when instantiating this decoder using {@link com.tradewright.tradebuild.tickutils.TickUtils#getTickEncoder getTickEncoder}.
' @return The minimum tick size for this encoder.
'@/
Public Property Get TickSize() As Double
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Encodes an ask tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
' @param pSize The tick size.
'@/
Public Sub EncodeAsk( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
End Sub

''
' Encodes a bid tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
' @param pSize The tick size.
'@/
Public Sub EncodeBid( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
End Sub

''
' Encodes a close tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
'@/
Public Sub EncodeClose( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
End Sub

''
' Encodes a high tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
'@/
Public Sub EncodeHigh( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
End Sub

''
' Encodes a low tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
'@/
Public Sub EncodeLow( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
End Sub

''
' Encodes a market depth tick.
' @param pTimestamp The tick's timestamp.
' @param position The position in the Depth-of-Market table.
' @param marketmaker The market maker's identifier.
' @param operation The operation to be performed in the Depth-of-Market table.
' @param side The side of the Depth-of-Market table affected.
' @param pPrice The tick price.
' @param pSize The tick size.
'@/
Public Sub EncodeMarketDepthData( _
                ByVal pTimestamp As Date, _
                ByVal pPosition As Long, _
                ByVal pMarketmaker As String, _
                ByVal pOperation As Long, _
                ByVal pSide As Long, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
End Sub

''
' Encodes a market depth reset event.
' @param pTimestamp The tick's timestamp.
'@/
Public Sub EncodeMarketDepthReset( _
                ByVal pTimestamp As Date)
End Sub

''
' Encodes an open tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
'@/
Public Sub EncodeOpen( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double)
End Sub

''
' Encodes an open interest tick.
' @param pTimestamp The tick's timestamp.
' @param pSize The tick size.
'@/
Public Sub EncodeOpenInterest( _
                ByVal pTimestamp As Date, ByVal pSize As BoxedDecimal)
End Sub

''
' Encodes a trade tick.
' @param pTimestamp The tick's timestamp.
' @param pPrice The tick price.
' @param pSize The tick size.
'@/
Public Sub EncodeTrade( _
                ByVal pTimestamp As Date, _
                ByVal pPrice As Double, _
                ByVal pSize As BoxedDecimal)
End Sub

''
' Encodes a volume tick.
' @param pTimestamp The tick's timestamp.
' @param pSize The tick size.
'@/
Public Sub EncodeVolume( _
                ByVal pTimestamp As Date, _
                ByVal pSize As BoxedDecimal)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

