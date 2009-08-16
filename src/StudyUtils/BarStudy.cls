VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' An interface to be implemented by studies that generate OHLC Bars.
'
' A <code>BarStudy</code> can act as the base study for a timeframe by generating
' Bars based on the timeframe's bar length and bar unit specification.
'
' <p>When a study is added to a study that implements <code>BarStudy</code>, the
' historical values of open, high, low and close are replayed as input values.</p>
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

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
' Returns the <code>Bars</code> object containing the Bars created by this
' <code>BarStudy object.</code>
'
' @return the<code>Bars</code> object.
'
'@/
Public Property Get Bars() As Bars

End Property

''
' The Name of the value that represents the closing value of a bar.
'
' @return The closing value Name.
'
'@/
Public Property Get CloseValueName() As String

End Property

''
' The Name of the value that represents the high value of a bar.
'
' @return The high value Name.
'
'@/
Public Property Get HighValueName() As String

End Property

''
' The Name of the value that represents the low value of a bar.
'
' @return The low value Name.
'
'@/
Public Property Get LowValueName() As String

End Property

''
' The Name of the value that represents the opening value of a bar.
'
' @return The opening value Name.
'
'@/
Public Property Get OpenValueName() As String

End Property

'@================================================================================
' Methods
'@================================================================================

''
' Switches the <code>BarStudy</code> from its current <code>InputStudy</code>
' to the specified <code>InputStudy</code>.
'
' @param newUnderlyingStudy The <code>InputStudy</code> from which this <code>BarStudy</code>
'                           will now receive input.
'
'@/
Public Sub SwitchUnderlyingStudy( _
                ByVal newUnderlyingStudy As InputStudy)
            
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


