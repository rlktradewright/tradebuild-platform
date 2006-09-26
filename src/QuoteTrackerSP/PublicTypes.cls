VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicTypes"
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

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnConnected
End Enum

Public Enum ErrorCodes
    ' generic run-time error codes defined by VB
    ErrInvalidProcedureCall = 5
    ErrOverflow = 6
    ErrSubscriptOutOfRange = 9
    ErrDivisionByZero = 11
    ErrTypeMismatch = 13
    ErrFileNotFound = 53
    ErrFileAlreadyOpen = 55
    ErrFileAlreadyExists = 58
    ErrDiskFull = 61
    ErrPermissionDenied = 70
    ErrPathNotFound = 76
    ErrInvalidObjectReference = 91
    
    ErrInvalidPropertyValue = 380
    ErrInvalidPropertyArrayIndex = 381
    
    ' generic error codes
    ErrArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ErrArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ErrClassCastException              ' attempt to cast an object to class of which it is not an instance
    ErrIllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    ErrIllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    ErrIndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    ErrNullPointerException            ' attempt to use Nothing in a case where an object is required
    ErrNumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    ErrRuntimeException                ' an unspecified runtime error has occurred
    ErrSecurityException               ' a security violation has occurred
    ErrUnsupportedOperationException   ' the requested operation is not supported
End Enum

Public Enum SPErrorCodes
    CantConnectToQuoteTracker = 700
    LogonResponseCannotBeParsed
    FieldDescriptionCannotBeParsed
    ErrorCannotBeParsed
    PasswordInvalid
    UnexpectedError
    ErrorFromQuoteTracker
End Enum

Public Enum TickTypes
    None
    bid
    ask
    last
    volume
    prevClose
    high
    low
    openInterest
End Enum

'================================================================================
' Types
'================================================================================

Public Type HistoricBar
    timestamp       As Date
    openValue       As Double
    highValue       As Double
    lowValue        As Double
    closeValue      As Double
    volume          As Long
End Type

Public Type QTHistoricalDataRequestQueueEntry
    symbol                  As String
    startDateTime           As Date
    endDateTime             As Date
    barLengthMinutes        As Long
    rth                     As Boolean
End Type

Public Type Tick
    TickType    As TickTypes
    timestamp   As Date
    Price       As Double
    size        As Long
End Type

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

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

