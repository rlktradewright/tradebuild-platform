VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddStudyTask"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Task

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

Private mTaskContext As TaskContext

Private mBars As bars

Private mStudy As TradeBuildSP.IStudy
Private mValueName As String

Private mBarsToNotify As Long
Private mBarsNotified As Long

'Private mNotifyNewbar As Boolean
Private mNotifyClose As Boolean
Private mNotifyVolume As Boolean
Private mNotifyTickvolume As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_run()
Dim lBar As Bar
Dim i As Long
Dim ev As TradeBuildSP.StudyValueEvent

For i = (mBars.Count - mBarsToNotify + mBarsNotified + 1) To _
        (mBars.Count - mBarsToNotify + mBarsNotified + 20)
    
    If mBarsNotified = mBars.Count Then Exit For
    
    mBarsNotified = mBarsNotified + 1
    
    Set lBar = mBars(i)
    
    If Not lBar.Blank Then
        Set ev.source = mBars
        ev.barNumber = i
        
'        If mNotifyNewbar Then
'            ' use the bar number as the value
'            ev.value = i
'            ev.valueName = "newbar"
'            mStudy.notify ev
'        End If
        
        If mNotifyClose Then
            ev.timestamp = lBar.datetime
            ev.value = lBar.openValue
            ev.ValueName = "Close"
            mStudy.notify ev
        End If
        
        If lBar.IsDownBar Then
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                ev.ValueName = "Close"
                mStudy.notify ev
            End If
            
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                ev.ValueName = "Close"
                mStudy.notify ev
            End If
        Else
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                ev.ValueName = "Close"
                mStudy.notify ev
            End If
            
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                ev.ValueName = "Close"
                mStudy.notify ev
            End If
        End If
        
        If mNotifyClose Then
            ev.timestamp = lBar.datetime + lBar.barLength - OneSecond
            ev.value = lBar.closeValue
            ev.ValueName = "Close"
            mStudy.notify ev
        End If
        
        If mNotifyVolume Then
            ev.value = lBar.volume
            ev.ValueName = "Volume"
            mStudy.notify ev
        End If
        
        If mNotifyTickvolume Then
            ev.value = lBar.tickVolume
            ev.ValueName = "Tick volume"
            mStudy.notify ev
        End If
    End If
Next

If mBarsNotified = mBars.Count Then
    mBars.studies.add mStudy, mValueName
    mTaskContext.finish
    Set mTaskContext = Nothing
End If

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal bars As bars, _
                ByVal study As TradeBuildSP.IStudy, _
                ByRef ValueName As String, _
                ByVal useAllUnderlyingValues As Boolean)

Set mBars = bars
Set mStudy = study
mValueName = LCase$(ValueName)

mBarsToNotify = study.numberOfBarsRequired
If mBarsToNotify = 0 Or useAllUnderlyingValues Then mBarsToNotify = mBars.Count
If mBarsToNotify > mBars.Count Then mBarsToNotify = mBars.Count

Select Case mValueName
Case "$default"
    mNotifyClose = True
'Case "newbar"
'    mNotifyNewbar = True
Case "close"
    mNotifyClose = True
Case "volume"
    mNotifyVolume = True
Case "tick volume"
    mNotifyTickvolume = True
End Select
End Sub

'================================================================================
' Helper Functions
'================================================================================

