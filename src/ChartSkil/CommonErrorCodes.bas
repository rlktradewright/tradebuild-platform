Attribute VB_Name = "mCommonErrorCodes"
Option Explicit

Public Enum CommonErrorCodes
    ' generic run-time error codes defined by VB
    InvalidProcedureCall = 5
    Overflow = 6
    SubscriptOutOfRange = 9
    DivisionByZero = 11
    TypeMismatch = 13
    FileNotFound = 53
    FileAlreadyOpen = 55
    FileAlreadyExists = 58
    DiskFull = 61
    PermissionDenied = 70
    PathNotFound = 76
    InvalidObjectReference = 91
    
    InvalidPropertyValue = 380
    InvalidPropertyArrayIndex = 381
    
    ' generic error codes
    ArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ClassCastException              ' attempt to cast an object to class of which it is not an instance
    IllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    IllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    IndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    NullPointerException            ' attempt to use Nothing in a case where an object is required
    NumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    RuntimeException                ' an unspecified runtime error has occurred
    SecurityException               ' a security violation has occurred
    UnsupportedOperationException   ' the requested operation is not supported



End Enum

