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

Private mContexts                                   As Contexts
Private mPrevContexts                               As Contexts

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
Set gInitialisationContext = mContexts.InitialisationContext
End Property

Public Property Get gResourceContext() As ResourceContext
Set gResourceContext = mContexts.ResourceContext
End Property

Public Property Get gStrategyLogger() As Logger
Static sLogger As Logger
If sLogger Is Nothing Then Set sLogger = GetLogger("strategy")
Set gStrategyLogger = sLogger
End Property

Public Property Get gStrategyRunner() As StrategyRunner
Set gStrategyRunner = mContexts.StrategyRunner
End Property

Public Property Get gStrategy() As Object
Set gStrategy = mContexts.Strategy
End Property

Public Property Get gTradingContext() As TradingContext
Set gTradingContext = mContexts.TradingContext
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateResourceIdentifier(ByVal pResource As Object) As ResourceIdentifier
Set gCreateResourceIdentifier = New ResourceIdentifier
gCreateResourceIdentifier.Initialise pResource
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
mContexts = mPrevContexts
Set mPrevContexts.StrategyRunner = Nothing
Set mPrevContexts.InitialisationContext = Nothing
Set mPrevContexts.TradingContext = Nothing
Set mPrevContexts.ResourceContext = Nothing
Set mPrevContexts.Strategy = Nothing
End Sub

Public Sub gSetStrategyRunner( _
                ByVal pStrategyRunner As StrategyRunner, _
                ByVal pInitialisationContext As InitialisationContext, _
                ByVal pTradingContext As TradingContext, _
                ByVal pResourceContext As ResourceContext, _
                ByVal pStrategy As Object)
Assert mPrevContexts.StrategyRunner Is Nothing, "Strategy switching error"
mPrevContexts = mContexts
Set mContexts.StrategyRunner = pStrategyRunner
Set mContexts.InitialisationContext = pInitialisationContext
Set mContexts.TradingContext = pTradingContext
Set mContexts.ResourceContext = pResourceContext
Set mContexts.Strategy = pStrategy
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




