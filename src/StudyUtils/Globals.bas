Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                        As String = "StudyUtils27"

Public Const AskInputName                           As String = "Ask"
Public Const BidInputName                           As String = "Bid"
Public Const OpenInterestInputName                   As String = "Open interest"
Public Const TickVolumeInputName                     As String = "Tick Volume"
Public Const TradeInputName                          As String = "Trade"
Public Const VolumeInputName                         As String = "Total Volume"

Public Const AttributeNameEnabled               As String = "Enabled"
Public Const AttributeNameStudyLibraryBuiltIn   As String = "BuiltIn"
Public Const AttributeNameStudyLibraryProgId    As String = "ProgId"

Public Const BuiltInStudyLibProgId              As String = "CmnStudiesLib27.StudyLib"
Public Const BuiltInStudyLibName                As String = "BuiltIn"

Public Const ConfigNameStudyLibraries           As String = "StudyLibraries"
Public Const ConfigNameStudyLibrary             As String = "StudyLibrary"

Public Const ConstTimeBarsStudyName                As String = "Constant time bars"
Public Const ConstTimeBarsParamBarLength           As String = "Bar length"
Public Const ConstTimeBarsParamTimeUnits           As String = "Time units"

Public Const ConstVolumeBarsStudyName              As String = "Constant Volume bars"
Public Const ConstTimeBarsParamVolPerBar           As String = "Volume per bar"

Public Const ConstMomentumBarsStudyName            As String = "Constant momentum bars"
Public Const ConstMomentumBarsParamTicksPerBar     As String = "Ticks move per bar"

Public Const DefaultStudyValueNameStr           As String = "$DEFAULT"
Public Const MovingAverageStudyValueNameStr     As String = "MA"

Public Const StudyLibrariesRenderer             As String = "StudiesUI27.StudyLibConfigurer"


'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

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

Public Property Get gErrorLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("error")
Set gErrorLogger = lLogger
End Property

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("studyutils", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

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

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



