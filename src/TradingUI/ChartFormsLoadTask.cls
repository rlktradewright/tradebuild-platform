VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartFormsLoadTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum LoadStates
    LoadStateNone
    LoadStateLoadingCharts
    LoadStateDeletingUnwantedCharts
    LoadStateShowingCharts
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartFormsLoadTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mChartForms                                 As ChartForms
Private mConfig                                     As ConfigurationSection
Private mTickers                                    As Tickers
Private mBarFormatterLibManager                     As BarFormatterLibManager
Private mTimePeriodValidator                        As ITimePeriodValidator
Private mOwner                                      As Variant
Private mTheme                                      As ITheme
Private mStudyLibraryManager                        As StudyLibraryManager
Private mHistDataStore                              As IHistoricalDataStore

Private mIsHistorical                               As Boolean
Private mState                                      As LoadStates

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sRemovableCSs As New EnumerableCollection
Static en As Enumerator

Dim lChartConfig As ConfigurationSection

If mState = LoadStateLoadingCharts Then
    If en Is Nothing Then Set en = mConfig.ChildEnumerator

    If en.MoveNext Then
        Set lChartConfig = en.Current
        If mIsHistorical Then
            If Not mChartForms.CreateHistoricalChartFromConfig(lChartConfig, mStudyLibraryManager, mHistDataStore, mBarFormatterLibManager) Then sRemovableCSs.Add lChartConfig
        Else
            If Not mChartForms.CreateChartFromConfig(lChartConfig, mTickers, mBarFormatterLibManager, mTimePeriodValidator) Then sRemovableCSs.Add lChartConfig
        End If
    Else
        Set en = Nothing
        If mIsHistorical Then
            mState = LoadStateShowingCharts
        Else
            mState = LoadStateDeletingUnwantedCharts
        End If
    End If
ElseIf mState = LoadStateDeletingUnwantedCharts Then
    If en Is Nothing Then Set en = sRemovableCSs.Enumerator
    If en.MoveNext Then
        Set lChartConfig = en.Current
        lChartConfig.Remove
    Else
        Set en = Nothing
        mState = LoadStateShowingCharts
    End If
ElseIf mState = LoadStateShowingCharts Then
    If IsMissing(mOwner) Or IsEmpty(mOwner) Then
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    If en Is Nothing Then Set en = mChartForms.Enumerator
    
    If en.MoveNext Then
        Dim lChart As fChart
        Set lChart = en.Current
        If (mIsHistorical And lChart.IsHistorical) Or _
            (Not mIsHistorical And Not lChart.IsHistorical) _
        Then
            lChart.Show vbModeless, mOwner
            If Not mTheme Is Nothing Then lChart.Theme = mTheme
        End If
    Else
        mTaskContext.Finish Empty, False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChartForms As ChartForms, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme)
Const ProcName As String = "Initialise"
On Error GoTo Err

mIsHistorical = False
Set mChartForms = pChartForms
Set mConfig = pConfig
Set mTickers = pTickers
Set mBarFormatterLibManager = pBarFormatterLibManager
Set mTimePeriodValidator = pTimePeriodValidator
gSetVariant mOwner, pOwner
Set mTheme = pTheme

mState = LoadStateLoadingCharts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseHistorical( _
                ByVal pChartForms As ChartForms, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme)
Const ProcName As String = "InitialiseHistorical"
On Error GoTo Err

mIsHistorical = True
Set mChartForms = pChartForms
Set mConfig = pConfig
Set mStudyLibraryManager = pStudyLibraryManager
Set mHistDataStore = pHistDataStore
Set mBarFormatterLibManager = pBarFormatterLibManager
gSetVariant mOwner, pOwner
Set mTheme = pTheme

mState = LoadStateLoadingCharts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




