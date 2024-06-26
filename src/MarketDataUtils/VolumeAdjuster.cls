VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VolumeAdjuster"
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

Private Const ModuleName                            As String = "VolumeAdjuster"

'@================================================================================
' Member variables
'@================================================================================

Private mGotFirstVolume                             As Boolean
Private mReceivedFirstVolumeInSession               As Boolean
Private mPrevSessionAccumulatedVolume               As BoxedDecimal
Private mVolumeCorrectionIncrement                  As BoxedDecimal
Private mAccumulatedVolume                          As BoxedDecimal

Private mLatestTrade                                As GenericTick

Private mNoVolumeAdjustments                        As Boolean
Private mNoImpliedTrades                            As Boolean

Private WithEvents mSession                         As Session
Attribute mSession.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPrevSessionAccumulatedVolume = DecimalZero
Set mVolumeCorrectionIncrement = DecimalZero
Set mAccumulatedVolume = DecimalZero
Set mLatestTrade.Size = DecimalZero
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted(ev As SessionEventData)
mReceivedFirstVolumeInSession = False
Set mPrevSessionAccumulatedVolume = mAccumulatedVolume
Set mAccumulatedVolume = DecimalZero
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function GetVolumeTickForTrade( _
                            ByRef pTradeTick As GenericTick) As GenericTick
Const ProcName As String = "GetVolumeTickForTrade"
On Error GoTo Err

If mNoVolumeAdjustments Or Not mGotFirstVolume Then Exit Function

mLatestTrade = pTradeTick

Set mAccumulatedVolume = mAccumulatedVolume.Add(pTradeTick.Size)
GetVolumeTickForTrade.TickType = TickTypeVolume
Set GetVolumeTickForTrade.Size = CreateBoxedDecimal(mAccumulatedVolume)
GetVolumeTickForTrade.Timestamp = pTradeTick.Timestamp

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrOverflow Then
    Dim errSource As String
    errSource = Err.Source
    GMktData.Logger.Log "Overflow: mAccumulatedVolume=" & mAccumulatedVolume & "; tick=" & GenericTickToString(pTradeTick), ProcName, ModuleName
    Err.Raise VBErrorCodes.VbErrOverflow, errSource, "Overflow"
End If
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pSession As Session, _
                ByVal pNoVolumeAdjustments As Boolean, _
                ByVal pNoImpliedTrades As Boolean)
Set mSession = pSession
mNoVolumeAdjustments = pNoVolumeAdjustments
mNoImpliedTrades = pNoImpliedTrades
End Sub

Friend Function ProcessVolumeTick( _
                            ByRef pTick As GenericTick, _
                            ByRef pAdjustedTick As GenericTick, _
                            ByRef pImpliedTick As GenericTick) As Boolean
Const ProcName As String = "ProcessVolumeTick"
On Error GoTo Err

mGotFirstVolume = True

If Not mReceivedFirstVolumeInSession Then
    ' this is the start of a new Session, but there may be a Volume
    ' tick including the accumulated Volume for the previous Session
    ' before the first true Volume tick for the new Session
    If mPrevSessionAccumulatedVolume = 0 Or _
        pTick.Size + mVolumeCorrectionIncrement < mPrevSessionAccumulatedVolume _
    Then
        ' this is the first true Volume tick for the new Session
        mReceivedFirstVolumeInSession = True
        Set mVolumeCorrectionIncrement = DecimalZero
    Else
        ' correct this Volume figure
        Set pTick.Size = pTick.Size.Subtr(mPrevSessionAccumulatedVolume.Subtr(mVolumeCorrectionIncrement))
    End If
End If

If isInvalidSize(pTick.Size) Then Exit Function

' we write the corrected Size to the tickfile rather than the raw Size because
' otherwise replaying from the start of the Session gives inflated Volume figures
' if the first tick was subject to correction

' note that the previous Trade will have notified Volume of mAccumulatedVolume

Dim lExpectedSize As BoxedDecimal
Set lExpectedSize = CreateBoxedDecimal(mAccumulatedVolume.Subtr(mVolumeCorrectionIncrement))

If pTick.Size = lExpectedSize Then
    ' nothing to do, because the volume will have been notified when processing
    ' the preceding Trade tick
    ProcessVolumeTick = False
ElseIf pTick.Size < lExpectedSize Then
    ' size is too small

    If mNoVolumeAdjustments Then
        ProcessVolumeTick = True
    Else
        Set mVolumeCorrectionIncrement = mAccumulatedVolume.Subtr(pTick.Size)
        ProcessVolumeTick = False
    End If

Else
    ' size is too big
    If mNoImpliedTrades Then
        pAdjustedTick = getAdjustedVolumeTick(pTick)
    ElseIf sizeImpliesUnnotifiedTrade(pTick.Size) Then
        pImpliedTick = getImpliedTradeTick(pTick, pTick.Size.Subtr(lExpectedSize))
    Else
        pAdjustedTick = getAdjustedVolumeTick(pTick)
    End If
    ProcessVolumeTick = False
End If
    
Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getAdjustedVolumeTick(ByRef pTick As GenericTick) As GenericTick
Set mAccumulatedVolume = pTick.Size.Add(mVolumeCorrectionIncrement)
getAdjustedVolumeTick.TickType = TickTypeVolume
getAdjustedVolumeTick.Timestamp = pTick.Timestamp
Set getAdjustedVolumeTick.Size = mAccumulatedVolume
End Function

Private Function getImpliedTradeTick(ByRef pTick As GenericTick, ByVal pImpliedSize As BoxedDecimal) As GenericTick
getImpliedTradeTick.TickType = TickTypeTrade
getImpliedTradeTick.Timestamp = pTick.Timestamp
getImpliedTradeTick.Price = mLatestTrade.Price
Set getImpliedTradeTick.Size = pImpliedSize
End Function

Private Function isInvalidSize(pSize As BoxedDecimal) As Boolean
Const ProcName As String = "isInvalidSize"
On Error GoTo Err

If pSize.GE(DecimalZero) Then Exit Function
    
isInvalidSize = True
' tickfile encoding can't cope with negative volumes, so just log it
GMktData.Logger.Log "Corrected Volume tick is negative" & vbCrLf & _
                "Corrected Size: " & DecimalToString(pSize) & vbCrLf & _
                "mPrevSessionAccumulatedVolume: " & DecimalToString(mPrevSessionAccumulatedVolume) & vbCrLf & _
                "mVolumeCorrectionIncrement: " & DecimalToString(mVolumeCorrectionIncrement), _
            ProcName, _
            ModuleName, _
            LogLevelWarning

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function sizeImpliesUnnotifiedTrade(pSize As BoxedDecimal) As Boolean
Dim lExpectedSize As BoxedDecimal
Set lExpectedSize = mAccumulatedVolume.Subtr(mVolumeCorrectionIncrement)
sizeImpliesUnnotifiedTrade = pSize.Subtr(lExpectedSize).EQ(mLatestTrade.Size)
End Function



