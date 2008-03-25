VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickProcessingTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements Task

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

Private Const ProjectName                   As String = "TBInfoBase26"
Private Const ModuleName                    As String = "TickProcessingTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTC                                 As TaskContext

Private mTickReader                         As TickDataReader
Attribute mTickReader.VB_VarHelpID = -1
Private mStudyManager                       As StudyManager
Private mPriceHandle                        As Long
Private mVolumeHandle                       As Long
Private mBarType                            As BarTypes

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_run()
Dim data() As Byte
Dim tdd As TickDataDecoder
Dim startTime As Date
Dim formatId As String
Dim basePrice As Double
Dim ticksize As Double
Dim tick As GenericTick

If mTickReader.ReadData(startTime, formatId, basePrice, ticksize, data) Then
    
    Set tdd = CreateTickDecoder(startTime, ticksize, basePrice, data, formatId)
    
    Do While tdd.getNextTick(tick)
        processTick tick
    Loop
Else
    mTC.finish 0, False
End If
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTC = RHS
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTC.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal pTickReader As TickDataReader, _
                ByVal pStudyManager As StudyManager, _
                ByVal pPriceHandle As Long, _
                ByVal pVolumeHandle As Long, _
                ByVal barType As BarTypes)
Set mTickReader = pTickReader
Set mStudyManager = pStudyManager
mPriceHandle = pPriceHandle
mVolumeHandle = pVolumeHandle
mBarType = barType
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processTick( _
                ByRef tick As GenericTick)
Select Case tick.tickType
Case TickTypeBid
    If mBarType = BarTypeBid Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeAsk
    If mBarType = BarTypeAsk Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeTrade
    If mBarType = BarTypeTrade Then mStudyManager.notifyInput mPriceHandle, tick.price, tick.timestamp
Case TickTypeVolume
    If mBarType = BarTypeTrade Then mStudyManager.notifyInput mVolumeHandle, tick.size, tick.timestamp
End Select
End Sub


