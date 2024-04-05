VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsScannerSubscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "TwsScannerSubscription"

'@================================================================================
' Member variables
'@================================================================================

Public NumberOfRows                                 As Long
Public Instrument                                   As String
Public LocationCode                                 As String
Public ScanCode                                     As String
Public AbovePrice                                   As Double
Public BelowPrice                                   As Double
Public AboveVolume                                  As Long
Public AverageOptionVolumeAbove                     As Long
Public MarketCapAbove                               As Double
Public MarketCapBelow                               As Double
Public MoodyRatingAbove                             As String
Public MoodyRatingBelow                             As String
Public SpRatingAbove                                As String
Public SpRatingBelow                                As String
Public MaturityDateAbove                            As String
Public MaturityDateBelow                            As String
Public CouponRateAbove                              As Double
Public CouponRateBelow                              As Double
Public ExcludeConvertible                           As String
Public ScannerSettingPairs                          As String
Public StockTypeFilter                              As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
NumberOfRows = NO_ROW_NUMBER_SPECIFIED
AbovePrice = MaxDouble
BelowPrice = MaxDouble
AboveVolume = MaxLong
AverageOptionVolumeAbove = MaxLong
MarketCapAbove = MaxDouble
MarketCapBelow = MaxDouble
CouponRateAbove = MaxDouble
CouponRateBelow = MaxDouble
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

Public Property Get NO_ROW_NUMBER_SPECIFIED() As Long
NO_ROW_NUMBER_SPECIFIED = -1
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




