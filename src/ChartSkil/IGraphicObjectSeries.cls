VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGraphicObjectSeries"
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


Private Const ModuleName                    As String = "IGraphicObjectSeries"

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

Public Property Get Count() As Long

End Property

Public Property Get IsSelectable() As Boolean

End Property

Public Property Get Name() As String

End Property

Public Property Let Selected(ByVal Value As Boolean)

End Property

Public Property Get Selected() As Boolean

End Property

Public Property Get SeriesID() As Long

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Click()

End Sub

Public Sub DblCLick()

End Sub

Public Sub Finish()

End Sub

Public Sub Initialise( _
                ByVal pName As String, _
                ByVal pViewport As Viewport, _
                ByVal pLayer As LayerNumbers, _
                ByVal pChart As Chart, _
                ByVal pSeriesID As Long, _
                ByVal pGOSP As GraphObjServiceProvider)

End Sub

Public Function Item(ByVal index As Long) As IGraphicObject

End Function

Public Sub Remove(ByVal index As Long)

End Sub

Public Sub RemoveObject(ByVal Value As IGraphicObject)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


