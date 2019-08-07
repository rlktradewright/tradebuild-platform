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

''
' Specifies whether initial bars can be defined for this <code>IBarStudy</code>.
'
' @remarks
'    If <code>True</code>, then the <code>InitialBarFetchListener</code> and
'    <code>InitialBarsFuture</code> properties are disabled.
'
' @param Value
'
'@/
'
Public Property Let AllowInitialBars(ByVal Value As Boolean)
End Property

''
' Returns an <code>IFuture</code> object that will yield the <code>Bars</code>Bars
' created by this <code>IBarStudy</code> object.
'
' @return The<code>Bars</code> object.
'
'@/
Public Property Get BarsFuture() As IFuture
End Property

''
' The Name of the Value that represents the closing Value of a bar.
'
' @return The closing Value Name.
'
'@/
Public Property Get CloseValueName() As String
End Property

''
' The Name of the Value that represents the high Value of a bar.
'
' @return The high Value Name.
'
'@/
Public Property Get HighValueName() As String
End Property

''
' Returns an <code>IBarFetchListener</code> object that can be used to build the
' intial bars dynamically, rather than via the <code>InitialBarsFuture</code>
' parameter.
'
' @remarks
'
' @return
'
' @param name
'
' @see
'
'@/
'
Public Property Get InitialBarFetchListener() As IBarFetchListener
End Property

''
' Sets the <code>Bars</code> object containing the initial bars for this
' <code>BarStudy</code> object.
'
' @param Value
'  An <code>IFuture</code> object that will provide the initial <code>Bars</code> for
'  this <code>IBarStudy</code>.
'
' @remarks
'   It is not necessary for all the initial bars to have been loaded into the
'   supplied <code>Bars</code> object at the time this method is called. However, the
'   <code>StudyManager</code>'s <code>StartStudy</code> method should not be called for
'   this <code>BarStudy</code> until loading of initial bars has been completed: if this
'   rule is not observed, any bars created as a result of values notified by the underlying
'   study may be interleaved with the initial bars, and hence be in the wrong sequence.
'@/
Public Property Let InitialBarsFuture(ByVal Value As IFuture)
End Property

''
' The Name of the Value that represents the low Value of a bar.
'
' @return The low Value Name.
'
'@/
Public Property Get LowValueName() As String
End Property

''
' The Name of the Value that represents the opening Value of a bar.
'
' @return The opening Value Name.
'
'@/
Public Property Get OpenValueName() As String
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


