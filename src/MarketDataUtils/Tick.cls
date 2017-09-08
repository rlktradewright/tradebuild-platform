VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tick"
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

Private Const ModuleName                            As String = "Tick"

'@================================================================================
' Member variables
'@================================================================================

Private mAllowedTickType                            As TickTypes

Private mTick                                       As GenericTick
Private mPreviousTick                               As GenericTick

Private mIsLive                                    As Boolean

Private mPriceChange                                As ValueChanges
Private mSizeChange                                 As ValueChanges

Private mPrevPriceChange                            As ValueChanges
Private mPrevSizeChange                             As ValueChanges

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mAllowedTickType = -1
mPriceChange = ValueChangeNone
mSizeChange = ValueChangeNone
mPrevPriceChange = ValueChangeNone
mPrevSizeChange = ValueChangeNone
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get IsLive() As Boolean
IsLive = mIsLive
End Property

Friend Property Get PriceChange() As ValueChanges
PriceChange = mPriceChange
End Property

Friend Property Get SizeChange() As ValueChanges
SizeChange = mSizeChange
End Property

Friend Property Get Quote() As Quote
Quote.Price = mTick.Price
Quote.PriceChange = mPriceChange
Quote.RecentPriceChange = mPrevPriceChange
Quote.RecentSizeChange = mPrevSizeChange
Quote.Size = mTick.Size
Quote.SizeChange = mSizeChange
Quote.Timestamp = mTick.Timestamp
End Property

Friend Property Let Tick(ByRef pTick As GenericTick)
Assert mAllowedTickType <> -1, "Tick object has not been initialised"
validateTickType pTick.TickType
Assert pTick.TickType = mAllowedTickType, "Supplied GenericTick not of correct tick type"

If mIsLive Then
    mPreviousTick = mTick
    
    mPrevPriceChange = mPriceChange
    mPriceChange = gCalcValueChange(pTick.Price, mPreviousTick.Price)
    
    mPrevSizeChange = mSizeChange
    mSizeChange = gCalcValueChange(pTick.Size, mPreviousTick.Size)
End If

mTick = pTick
mIsLive = True
End Property

Friend Property Get Tick() As GenericTick
Assert mAllowedTickType <> -1, "Tick object has not been initialised"
Assert mIsLive, "No tick available"

Tick = mTick
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pTickType As TickTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

validateTickType pTickType

mAllowedTickType = pTickType

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub validateTickType(ByVal Value As TickTypes)
Select Case Value
Case TickTypeBid
Case TickTypeAsk
Case TickTypeClosePrice
Case TickTypeHighPrice
Case TickTypeLowPrice
Case TickTypeMarketDepth
Case TickTypeMarketDepthReset
Case TickTypeTrade
Case TickTypeVolume
Case TickTypeOpenInterest
Case TickTypeOpenPrice
Case Else
    AssertArgument False, "Invalid tick type"
End Select
End Sub


