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
Private mStudyToNotify As TradeBuildSP.IStudy
Private mValueName As String

Private mStartBarNumber As Long
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

If mBars.Count = 0 Then
    mBars.studies.add mStudy, mValueName
    mTaskContext.finish
    Set mTaskContext = Nothing
    Exit Sub
End If

For i = (mStartBarNumber + mBarsNotified) To _
        (mStartBarNumber + mBarsNotified + 19)
    
    Set lBar = mBars(i)
    
    If Not lBar.Blank Then
        Set ev.source = mBars
        ev.barNumber = i
        
        If mNotifyClose Then
            ev.timestamp = lBar.datetime
            ev.value = lBar.openValue
            ev.ValueName = GBars.BarsValueClose
            mStudyToNotify.notify ev
        End If
        
        If lBar.IsDownBar Then
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                ev.ValueName = GBars.BarsValueClose
                mStudyToNotify.notify ev
            End If
            
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                ev.ValueName = GBars.BarsValueClose
                mStudyToNotify.notify ev
            End If
        Else
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (lBar.barLength / 3) / 1440
                ev.value = lBar.lowValue
                ev.ValueName = GBars.BarsValueClose
                mStudyToNotify.notify ev
            End If
            
            If mNotifyClose Then
                ev.timestamp = lBar.datetime + (2 * lBar.barLength / 3) / 1440
                ev.value = lBar.highValue
                ev.ValueName = GBars.BarsValueClose
                mStudyToNotify.notify ev
            End If
        End If
        
        If mNotifyClose Then
            ev.timestamp = lBar.datetime + lBar.barLength / 1440 - OneSecond
            ev.value = lBar.closeValue
            ev.ValueName = GBars.BarsValueClose
            mStudyToNotify.notify ev
        End If
        
        If mNotifyVolume Then
            ev.value = lBar.volume
            ev.ValueName = GBars.BarsValueVolume
            mStudyToNotify.notify ev
        End If
        
        If mNotifyTickvolume Then
            ev.value = lBar.tickVolume
            ev.ValueName = GBars.BarsValueTickVolume
            mStudyToNotify.notify ev
        End If
    End If
    
    mBarsNotified = mBarsNotified + 1
    If i = mBars.Count Then
        mBars.studies.add mStudyToNotify, mValueName
        mTaskContext.finish
        Set mTaskContext = Nothing
        Exit For
    End If
Next

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
                ByVal numUnderlyingValuesToUse As Long)
Dim numBarsToNotify As Long

Set mBars = bars
Set mStudy = study
Set mStudyToNotify = mStudy.baseStudy
mValueName = ValueName

If numUnderlyingValuesToUse < 0 Then
    ' just replay the number of values needed by the study
    numBarsToNotify = study.numberOfBarsRequired
Else
    numBarsToNotify = numUnderlyingValuesToUse
End If
If numBarsToNotify > mBars.Count Then numBarsToNotify = mBars.Count

mStartBarNumber = mBars.Count - numBarsToNotify + 1

Select Case UCase$(mValueName)
Case UCase$(DefaultStudyValue)
    mNotifyClose = True
Case UCase$(BarsValueClose)
    mNotifyClose = True
Case UCase$(BarsValueVolume)
    mNotifyVolume = True
Case UCase$(BarsValueTickVolume)
    mNotifyTickvolume = True
End Select
End Sub

'================================================================================
' Helper Functions
'================================================================================

