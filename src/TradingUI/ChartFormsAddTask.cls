VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartFormsAddTask"
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

Private Enum AddStates
    AddStateNone
    AddStateLoadingCharts
    AddStateShowingCharts
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartFormsAddTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mChartForms                                 As ChartForms
Private mConfig                                     As ConfigurationSection
Private mTickers                                    As SelectedTickers
Private mPeriodLength                               As TimePeriod
Private mSpec                                       As ChartSpecifier
Private mBarFormatterLibManager                     As BarFormatterLibManager
Private mTimePeriodValidator                        As ITimePeriodValidator
Private mStyle                                      As ChartStyle
Private mOwner                                      As Variant
Private mTheme                                      As ITheme

Private mIsHistorical                               As Boolean
Private mState                                      As AddStates

Private mCharts                                     As New EnumerableCollection

Private mContracts                                  As IContracts
Private mStudyLibManager                            As StudyLibraryManager
Private mHistDataStore                              As IHistoricalDataStore

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

Static en As Enumerator

If mState = AddStateLoadingCharts Then
    If mIsHistorical Then
        If en Is Nothing Then Set en = mContracts.Enumerator
    Else
        If en Is Nothing Then Set en = mTickers.Enumerator
    End If

    If en.MoveNext Then
        If mIsHistorical Then
            Dim lContract As IContract: Set lContract = en.Current
            mCharts.Add mChartForms.CreateHistoricalChart( _
                                        mPeriodLength, _
                                        CreateFuture(lContract), _
                                        mStudyLibManager.CreateStudyManager( _
                                            lContract.SessionStartTime, _
                                            lContract.SessionEndTime, _
                                            GetTimeZone(lContract.TimezoneName)), _
                                        mHistDataStore, _
                                        mBarFormatterLibManager, _
                                        mConfig, _
                                        mSpec, _
                                        mStyle, _
                                        mOwner)
        Else
            Dim lTicker As Ticker
            Set lTicker = en.Current
            If lTicker.State = MarketDataSourceStateRunning Then
                mCharts.Add mChartForms.CreateChart( _
                                            lTicker, _
                                            mPeriodLength, _
                                            lTicker.Timeframes, _
                                            mBarFormatterLibManager, _
                                            mTimePeriodValidator, _
                                            IIf(lTicker.IsTickReplay, Nothing, mConfig), _
                                            mSpec, _
                                            mStyle, _
                                            mOwner)
            End If
        End If
    Else
        Set en = Nothing
        mState = AddStateShowingCharts
    End If
ElseIf mState = AddStateShowingCharts Then
    If IsMissing(mOwner) Or IsEmpty(mOwner) Then
        mTaskContext.Finish Empty, False
        Exit Sub
    End If
    
    If en Is Nothing Then Set en = mCharts.Enumerator
    
    If en.MoveNext Then
        Dim lChart As fChart: Set lChart = en.Current
        lChart.Show vbModeless, mOwner
        If Not mTheme Is Nothing Then lChart.Theme = mTheme
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
                ByVal pTickers As SelectedTickers, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme)
Const ProcName As String = "Initialise"
On Error GoTo Err

mIsHistorical = False
Set mChartForms = pChartForms
Set mTickers = pTickers
Set mPeriodLength = pPeriodLength
Set mBarFormatterLibManager = pBarFormatterLibManager
Set mTimePeriodValidator = pTimePeriodValidator
Set mConfig = pConfig
Set mSpec = pSpec
Set mStyle = pStyle
gSetVariant mOwner, pOwner
Set mTheme = pTheme

mState = AddStateLoadingCharts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseHistorical( _
                ByVal pChartForms As ChartForms, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pContracts As IContracts, _
                ByVal pStudyLibManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme)
Const ProcName As String = "InitialiseHistorical"
On Error GoTo Err

mIsHistorical = True
Set mChartForms = pChartForms
Set mPeriodLength = pPeriodLength
Set mContracts = pContracts
Set mStudyLibManager = pStudyLibManager
Set mHistDataStore = pHistDataStore
Set mBarFormatterLibManager = pBarFormatterLibManager
Set mConfig = pConfig
Set mSpec = pSpec
Set mStyle = pStyle
gSetVariant mOwner, pOwner
Set mTheme = pTheme

mState = AddStateLoadingCharts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
'@================================================================================
' Helper Functions
'@================================================================================





