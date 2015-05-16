VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBarStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' An interface to be implemented by studies that generate OHLC Bars.
'
' An <code>IBarStudy</code> can act as the base study for a timeframe by generating
' Bars based on the timeframe's bar length and bar unit specification.
'
' <p>When a study is added to a study that implements <code>IBarStudy</code>, the
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

Public Property Get InitialBarFetchListener() As IBarFetchListener
End Property

''
' Returns the <code>Bars</code> object containing the Bars created by this
' <code>BarStudy</code> object.
'
' @return the<code>Bars</code> object.
'
'@/
Public Property Get BarsFuture() As IFuture
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
' Sets the <code>Bars</code> object containing the initial bars for this
' <code>BarStudy</code> object.
'
' @param Value the initial <code>Bars</code> object.
'
' @remarks
'   It is not necessary for all the initial bars to have been loaded into the
'   supplied <code>Bars</code> object at the time this method is called. However, the
'   <code>StudyManager</code>'s <code>StartStudy</code> method should not be called for
'   this <code>BarStudy</code> until loading of initial bars has been completed: if this
'   rule is not observed, any bars created as a result of values notified by the underlying
'   study may be interleaved with the initial bars, and hence be in the wrong sequence.
'
'   An <code>ErrorCodes.ErrIllegalStateException</code> error is raised if this property
'   is set after the <code>Bars</code> property has been accessed.
'
'@/
Public Property Let InitialBarsFuture(ByVal Value As IFuture)
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

'''
'' Notify each new bar to the <code>BarStudy</code> object as it is created.
''
''@/
'Public Sub NotifyBar(ByVal pBar As Bar)
'End Sub

'''
'' Switches the <code>BarStudy</code> from its current <code>InputStudy</code>
'' to the specified <code>InputStudy</code>.
''
'' @param newUnderlyingStudy The <code>InputStudy</code> from which this <code>BarStudy</code>
''                           will now receive input.
''
''@/
'Public Sub SwitchUnderlyingStudy( _
'                ByVal newUnderlyingStudy As StudyInputHandler)
'
'End Sub

'@================================================================================
' Helper Functions
'@================================================================================


