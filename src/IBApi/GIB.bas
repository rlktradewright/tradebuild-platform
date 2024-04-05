Attribute VB_Name = "GIB"
Option Explicit

'================================================================================
' Constants
'================================================================================

#If SingleDll = 0 Then
Public Const ProjectName                        As String = "IBBaseApi"
#End If
Private Const ModuleName                        As String = "GIB"

Public Const MaxLong                            As Long = &H7FFFFFFF
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#
Public Const OneSecond                          As Double = 1# / 86400#

Public Const Infinity                           As String = "Infinity"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                Destination As Any, _
                source As Any, _
                ByVal length As Long)
                            
Public Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                Destination As Any, _
                source As Any, _
                ByVal length As Long)
                            
'================================================================================
' Private variables
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Logger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradebuild.log.ibapi", ProjectName)
Set Logger = sLogger
End Property

Public Property Get SocketLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tradebuild.log.ibapi.socket", ProjectName)
Set SocketLogger = sLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Sub HandleUnexpectedError( _
                ByVal pErrorHandler As IProgramErrorListener, _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

Dim ev As ErrorEventData

If Not pErrorHandler Is Nothing Then
    ev.ErrorCode = errNum
    ev.ErrorMessage = errDesc
    ' ensure the calling proc's details are included in the error source
    ev.ErrorSource = errSource

    On Error GoTo HandleError
    pErrorHandler.NotifyUnexpectedProgramError ev
    ' should never get here!
End If

TWUtilities40.HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource

' will never get here!
Exit Sub

HandleError:
TWUtilities40.HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, Err.Number, Err.Description, Err.source
End Sub

Public Sub NotifyUnhandledError( _
                ByVal pErrorHandler As IProgramErrorListener, _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

Dim ev As ErrorEventData

If Not pErrorHandler Is Nothing Then
    ev.ErrorCode = errNum
    ev.ErrorMessage = errDesc
    ' ensure the calling proc's details are included in the error source
    ev.ErrorSource = errSource
    
    On Error GoTo HandleError
    ' ensure the calling proc's details are included in the error source
    pErrorHandler.NotifyUnhandledProgramError ev
    ' should never get here!
End If

TWUtilities40.UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource

' will never get here!
Exit Sub

HandleError:

' if we get to here, the error handler hasn't called the unhandled error mechanism, so we will
TWUtilities40.UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub Log(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "Log"
On Error GoTo Err


GIB.Logger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

