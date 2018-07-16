VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGraphicObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "IGraphicObject"

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

Public Property Get BoundingRectangle() As TRectangle

End Property

Public Property Get BoundingRectangleInViewport() As TRectangle

End Property

Public Property Get BoundingRectanglesAt( _
                ByVal PeriodNumber As Long) As TRectangle()

End Property

Public Property Get Capabilities() As GraphicObjectCapabilities

End Property

Public Property Get ExtendedObject() As Boolean

End Property

Public Property Get GaugeDependent() As Boolean

End Property

Public Property Let Handle(ByVal Value As Long)

End Property

Public Property Get Handle() As Long

End Property

Public Property Get IncludeInAutoscale() As Boolean

End Property

Public Property Let InScope(ByVal Value As Boolean)

End Property

Public Property Get IsSelectable() As Boolean

End Property

Public Property Get Layer() As LayerNumbers

End Property

Public Property Get NoDraw() As Boolean

End Property

Public Property Get PeriodNumber() As Long

End Property

Public Property Get ScaleDependent() As Boolean

End Property

Public Property Let Selected(ByVal Value As Boolean)

End Property

Public Property Get Selected() As Boolean

End Property

Public Property Get SeriesID() As Long

End Property

Public Property Get Timestamp() As Date

End Property

Public Property Get VerticalRangeInXinterval(ByRef pXInterval As TInterval) As TInterval

End Property

Public Property Let Visible(ByVal Value As Boolean)

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Click()

End Sub

Public Sub DblCLick()

End Sub

Public Sub Draw( _
                ByRef areas() As TRectangle)

End Sub

Public Sub Finish()

End Sub

Public Function HitTest( _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
                
End Function

'@================================================================================
' Helper Functions
'@================================================================================


