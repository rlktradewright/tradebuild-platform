Attribute VB_Name = "Globals"
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

Public Enum OrderRoles
    OrderRoleEntry = BracketOrderRoles.BracketOrderRoleEntry
    OrderRoleStopLoss = BracketOrderRoles.BracketOrderRoleStopLoss
    OrderRoleTarget = BracketOrderRoles.BracketOrderRoleTarget
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type Contexts
    Strategy                    As Object
    StrategyRunner              As StrategyRunner
    InitialisationContext       As InitialisationContext
    TradingContext              As TradingContext
    ResourceContext             As ResourceContext
End Type

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "StrategyUtils27"
Private Const ModuleName                            As String = "Globals"

'@================================================================================
' Member variables
'@================================================================================

'Private mContexts                                   As Contexts
Private mContextsStack()                            As Contexts
Private mContextsIndex                              As Long

Private mInitialised                                As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get gInitialisationContext() As InitialisationContext
Set gInitialisationContext = mContextsStack(mContextsIndex).InitialisationContext
End Property

Public Property Get gResourceContext() As ResourceContext
Set gResourceContext = mContextsStack(mContextsIndex).ResourceContext
End Property

Public Property Get gStrategyLogger() As Logger
Static sLogger As Logger
If sLogger Is Nothing Then Set sLogger = GetLogger("strategy")
Set gStrategyLogger = sLogger
End Property

Public Property Get gStrategyRunner() As StrategyRunner
Set gStrategyRunner = mContextsStack(mContextsIndex).StrategyRunner
End Property

Public Property Get gStrategy() As Object
Set gStrategy = mContextsStack(mContextsIndex).Strategy
End Property

Public Property Get gTradingContext() As TradingContext
Set gTradingContext = mContextsStack(mContextsIndex).TradingContext
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateResourceIdentifier(ByVal pResource As Object) As ResourceIdentifier
Set gCreateResourceIdentifier = New ResourceIdentifier
gCreateResourceIdentifier.Initialise pResource
End Function

Public Function gCreateStrategyRunner( _
                ByVal pStrategyHost As IStrategyHost) As StrategyRunner
Const ProcName As String = "gCreateStrategyRunner"
On Error GoTo Err

AssertArgument Not pStrategyHost Is Nothing, "pStrategyHost is Nothing"

Set gCreateStrategyRunner = New StrategyRunner
gCreateStrategyRunner.Initialise pStrategyHost

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gClearStrategyRunner()
Set mContextsStack(mContextsIndex).InitialisationContext = Nothing
Set mContextsStack(mContextsIndex).ResourceContext = Nothing
Set mContextsStack(mContextsIndex).Strategy = Nothing
Set mContextsStack(mContextsIndex).StrategyRunner = Nothing
Set mContextsStack(mContextsIndex).TradingContext = Nothing
mContextsIndex = mContextsIndex - 1
End Sub

Public Sub gSetStrategyRunner( _
                ByVal pStrategyRunner As StrategyRunner, _
                ByVal pInitialisationContext As InitialisationContext, _
                ByVal pTradingContext As TradingContext, _
                ByVal pResourceContext As ResourceContext, _
                ByVal pStrategy As Object)
If Not mInitialised Then
    ReDim mContextsStack(7) As Contexts
    mContextsIndex = -1
    mInitialised = True
End If

mContextsIndex = mContextsIndex + 1
If mContextsIndex > UBound(mContextsStack) Then ReDim Preserve mContextsStack((2 * (UBound(mContextsStack) + 1)) - 1) As Contexts
Set mContextsStack(mContextsIndex).StrategyRunner = pStrategyRunner
Set mContextsStack(mContextsIndex).InitialisationContext = pInitialisationContext
Set mContextsStack(mContextsIndex).TradingContext = pTradingContext
Set mContextsStack(mContextsIndex).ResourceContext = pResourceContext
Set mContextsStack(mContextsIndex).Strategy = pStrategy
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




