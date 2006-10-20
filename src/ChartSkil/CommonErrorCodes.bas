Attribute VB_Name = "mCommonErrorCodes"
Option Explicit

Public Enum CommonErrorCodes
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

