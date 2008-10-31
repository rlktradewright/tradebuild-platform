VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectableObjectHandler"
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

Private Const ProjectName                   As String = "ChartUtils26"
Private Const ModuleName                    As String = "ConstituentObjectHandler"

'@================================================================================
' Member variables
'@================================================================================

Private mObj                                As IGraphicObject

Private mIndex                              As Long

Private mAbstractToolRef                    As WeakReference

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

Private Sub Initialise( _
                ByVal abstractTool As AbstractChartTool, _
                ByVal selHandle As DataPoint, _
                ByVal index As Long)
Set mAbstractToolRef = CreateWeakReference(abstractTool)
Set mSelHandle = selHandle
mIndex = index
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


Private Sub mSelHandle_SelectionStateChanged()
If mSelHandle.IsSelected Then

End Sub


