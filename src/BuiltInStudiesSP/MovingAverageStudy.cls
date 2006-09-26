VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovingAverageStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
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

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get falling( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue

End Property

Public Property Get flat( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue

End Property

Public Property Get ma(Optional ByVal ref As Long) As StudyValue

End Property

Public Property Get rising( _
    Optional ByVal ref As Long, _
    Optional ByVal periods As Long = 2) As StudyValue

End Property

Public Property Get slope(Optional ByVal ref As Long, _
                            Optional ByVal periods As Long = 2) As StudyValue

End Property

Public Property Let slopeThreshold(ByVal value As Long)

End Property

'================================================================================
' Methods
'================================================================================

Public Function addStudy( _
                study As IStudy, _
                ByVal ValueName As String, _
                ByVal useAllUnderlyingValues As Boolean) As TradeBuildSP.ITaskCompletion

End Function

Public Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal numberOfValuesToCache As Long, _
                ByVal params As IParameters)

End Sub

Public Function numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long

End Function

'================================================================================
' Helper Functions
'================================================================================




