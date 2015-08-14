VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{464F646E-C78A-4AAC-AC11-FBC7E41F58BB}#217.0#0"; "StudiesUI27.ocx"
Object = "{5EF6A0B6-9E1F-426C-B84A-601F4CBF70C4}#249.0#0"; "ChartSkil27.ocx"
Begin VB.Form StudyTestForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TradeBuild Study Test Harness v2.7"
   ClientHeight    =   10365
   ClientLeft      =   5070
   ClientTop       =   3540
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   12840
   Begin VB.CommandButton TestButton 
      Caption         =   "Test"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      TabIndex        =   9
      ToolTipText     =   "Test the study"
      Top             =   120
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   16325
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Test data and results"
      TabPicture(0)   =   "StudyTestForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TestDataGrid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TestDataFilenameText"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FindFileButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MinimumPriceTickText"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Study setup"
      TabPicture(1)   =   "StudyTestForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ErrorText"
      Tab(1).Control(1)=   "StudyConfigurer1"
      Tab(1).Control(2)=   "RemoveLibButton"
      Tab(1).Control(3)=   "StudyLibraryList"
      Tab(1).Control(4)=   "AddLibButton"
      Tab(1).Control(5)=   "LibToAddText"
      Tab(1).Control(6)=   "StudiesCombo"
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(8)=   "Label1"
      Tab(1).Control(9)=   "Label2"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Chart"
      TabPicture(2)   =   "StudyTestForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ChartToolbar1"
      Tab(2).Control(1)=   "Chart1"
      Tab(2).ControlCount=   2
      Begin VB.TextBox ErrorText 
         BackColor       =   &H8000000F&
         Height          =   2895
         Left            =   -67560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   480
         Width           =   4935
      End
      Begin ChartSkil27.ChartToolbar ChartToolbar1 
         Height          =   330
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   6465
         _ExtentX        =   9737
         _ExtentY        =   582
      End
      Begin ChartSkil27.Chart Chart1 
         Height          =   8415
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   14843
      End
      Begin StudiesUI27.StudyConfigurer StudyConfigurer1 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   16
         Top             =   3480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9975
      End
      Begin VB.TextBox MinimumPriceTickText 
         Height          =   285
         Left            =   9480
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton RemoveLibButton 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   7
         ToolTipText     =   "Remove the selected service provider from the list"
         Top             =   1860
         Width           =   855
      End
      Begin VB.ListBox StudyLibraryList 
         Height          =   840
         ItemData        =   "StudyTestForm.frx":0054
         Left            =   -72600
         List            =   "StudyTestForm.frx":0056
         TabIndex        =   6
         ToolTipText     =   "Lists all studies service providers you need (except the built-in studies service provider)"
         Top             =   1860
         Width           =   3975
      End
      Begin VB.CommandButton AddLibButton 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   2
         ToolTipText     =   "Add this service provider to the list"
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox LibToAddText 
         Height          =   285
         Left            =   -72600
         TabIndex        =   1
         ToolTipText     =   "Enter the program id of any other studies service provider your service provider needs"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.ComboBox StudiesCombo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72600
         TabIndex        =   0
         ToolTipText     =   "Select the study to test"
         Top             =   540
         Width           =   3975
      End
      Begin VB.CommandButton FindFileButton 
         Caption         =   "..."
         Default         =   -1  'True
         Height          =   285
         Left            =   6720
         TabIndex        =   4
         ToolTipText     =   "Click to browse for the test data file"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox TestDataFilenameText 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The file that contains the test data"
         Top             =   840
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid TestDataGrid 
         Height          =   7935
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1260
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   -2147483636
         Appearance      =   0
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum price tick"
         Height          =   255
         Left            =   7920
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Configure the study - selected output values will appear both on the chart and in the grid"
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   3120
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "Other study libraries to include"
         Height          =   615
         Left            =   -74760
         TabIndex        =   14
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Test data file"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Study to test"
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "StudyTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ILogListener
                                    
'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "StudyTestForm"

Private Const BuiltInStudyLib           As String = "CmnStudiesLib27.StudyLib"

Private Const PriceRegionName           As String = "$price"
Private Const VolumeRegionName          As String = "$volume"

Private Const InputValuePrice           As String = "Price"
Private Const InputValueVolume          As String = "Total volume"

'================================================================================
' Enums
'================================================================================

Private Enum SSTabs
    SSTabData
    SSTabStudies
    SSTabChart
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mUnhandledErrorHandler As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1

Private mIsStudySet                     As Boolean

Private mChartManager                   As ChartManager
Private mStudyLibraryManager            As New StudyLibraryManager
Private mBarFormatterLibManager         As New BarFormatterLibManager
Private mStudyManager                   As StudyManager
Private mSourceStudy                    As IStudy
Private mBarsStudy                      As IStudy
Private mInitialStudyConfigs            As StudyConfigurations

Private mPriceInputHandle               As Long
Private mVolumeInputHandle              As Long

Private mIsInDev                        As Boolean

Private mBars                           As Bars

Private mGrid                           As GridManager

Private mAccumulatedVolume              As Long

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Debug.Print "Running in development environment: " & CStr(inDev)

InitialiseCommonControls  ' enables WinXP look and feel
InitialiseTWUtilities
Set mUnhandledErrorHandler = UnhandledErrorHandler

ApplicationGroupName = "TradeWright"
ApplicationName = App.Title & "-V" & App.Major & "-" & App.Minor
SetupDefaultLogging Command

Set mStudyManager = mStudyLibraryManager.CreateStudyManager
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

Set mGrid = New GridManager
mGrid.Initialise TestDataGrid
mGrid.SetupDataColumns

addStudyLibraries

' need to do this in case the user sets up his Study Library and study
' before loading the test data
setupInitialStudies

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub Form_Terminate()
Const ProcName As String = "Form_Terminate"
On Error GoTo Err

TerminateTWUtilities

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' ILogListener Interface Members
'================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
ErrorText = CStr(Logrec.Data)
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddLibButton_Click()
Const ProcName As String = "AddLibButton_Click"
On Error GoTo Err

If UCase$(LibToAddText) = UCase$(BuiltInStudyLib) Then
    MsgBox "This study library is already available"
ElseIf addStudyLibToList(LibToAddText) Then
    LibToAddText = ""
Else
    MsgBox "'" & LibToAddText & "' is not a valid Study Library"
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FindFileButton_Click()
Const ProcName As String = "FindFileButton_Click"
On Error GoTo Err

'CommonDialog1.CancelError = True
On Error GoTo Err

CommonDialog1.MaxFileSize = 32767
CommonDialog1.DialogTitle = "Open test data file"
CommonDialog1.Filter = "TradeBuild bar data files (*.tbd)|*.tbd"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNFileMustExist + _
                    cdlOFNLongNames + _
                    cdlOFNPathMustExist + _
                    cdlOFNExplorer + _
                    cdlOFNReadOnly
CommonDialog1.ShowOpen

TestDataFilenameText = CommonDialog1.FileName

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MinimumPriceTickText_KeyPress(KeyAscii As Integer)
Const ProcName As String = "MinimumPriceTickText_KeyPress"
On Error GoTo Err

If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then Exit Sub
If Not IsNumeric(MinimumPriceTickText & Chr(KeyAscii)) Then KeyAscii = 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RemoveLibButton_Click()
Const ProcName As String = "RemoveLibButton_Click"
On Error GoTo Err

StudyLibraryList.RemoveItem StudyLibraryList.ListIndex
RemoveLibButton.Enabled = False

mStudyLibraryManager.RemoveAllStudyLibraries
addStudyLibraries

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = SSTabs.SSTabStudies Then StudiesCombo.SetFocus
End Sub

Private Sub StudyLibraryList_Click()
Const ProcName As String = "StudyLibraryList_Click"
On Error GoTo Err

If StudyLibraryList.ListIndex = -1 Then
    RemoveLibButton.Enabled = False
Else
    RemoveLibButton.Enabled = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LibToAddText_Change()
Const ProcName As String = "LibToAddText_Change"
On Error GoTo Err

If LibToAddText = "" Then
    AddLibButton.Enabled = False
Else
    AddLibButton.Enabled = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub StudiesCombo_Click()
Dim regionNames(1) As String

Const ProcName As String = "StudiesCombo_Click"
On Error GoTo Err

setupInitialStudies
                    
regionNames(0) = PriceRegionName
regionNames(1) = VolumeRegionName

StudyConfigurer1.Initialise mStudyLibraryManager.GetStudyDefinition(StudiesCombo), _
                            "", _
                            regionNames, _
                            mChartManager.BaseStudyConfiguration, _
                            Nothing, _
                            mStudyLibraryManager.GetStudyDefaultParameters(StudiesCombo), _
                            False
mIsStudySet = True
If Not mBars Is Nothing Then TestButton.Enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TestButton_Click()
Const ProcName As String = "TestButton_Click"
Dim failPoint As String
On Error GoTo Err

Dim lBar As BarUtils27.Bar
Dim studyToTest As IStudy
Dim testStudyConfig As StudyConfiguration
Dim addTestStudyToSource As Boolean
Dim regionNames(1) As String

Screen.MousePointer = MousePointerConstants.vbArrowHourglass

ErrorText = ""

mGrid.Cols = TestDataGridColumns.StudyValue1

failPoint = "adding study libraries"
addStudyLibraries

Set testStudyConfig = StudyConfigurer1.StudyConfiguration
If testStudyConfig.UnderlyingStudy Is mSourceStudy Then
    addTestStudyToSource = True
End If

failPoint = "creating the study to be tested"

setupInitialStudies

If addTestStudyToSource Then
    testStudyConfig.UnderlyingStudy = mSourceStudy
Else
    testStudyConfig.UnderlyingStudy = mBarsStudy
End If

Set studyToTest = mChartManager.AddStudyConfiguration(testStudyConfig)
mChartManager.StartStudy studyToTest

' now re-setup the study configurer so that only current
' objects are referenced
regionNames(0) = PriceRegionName
regionNames(1) = VolumeRegionName
StudyConfigurer1.Initialise mStudyLibraryManager.GetStudyDefinition(StudiesCombo), _
                            "", _
                            regionNames, _
                            mChartManager.BaseStudyConfiguration, _
                            testStudyConfig, _
                            mStudyLibraryManager.GetStudyDefaultParameters(StudiesCombo), _
                            False

failPoint = "setting up the Study Value grid"
mGrid.SetupStudyValueColumns testStudyConfig

Chart1.Regions.Item(PriceRegionName).YScaleQuantum = CDbl(MinimumPriceTickText)

failPoint = "testing the study"
testStudy testStudyConfig

Chart1.FirstVisiblePeriod = 1

Screen.MousePointer = MousePointerConstants.vbDefault
Exit Sub

Err:

Do Until Chart1.IsDrawingEnabled
    Chart1.EnableDrawing
Loop

Screen.MousePointer = MousePointerConstants.vbDefault

gHandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failPoint
End Sub

Private Sub TestDataFilenameText_Change()
Const ProcName As String = "TestDataFilenameText_Change"
On Error GoTo Err

LoadData

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TestDataGrid_DblClick()
mChartManager.ScrollToTime mBars.Bar(mGrid.Row).TimeStamp
SSTab1.Tab = SSTabChart
End Sub

'================================================================================
' mUnhandledErrorHandler Event Handlers
'================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)
On Error Resume Next    ' ignore any further errors that might arise


MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly. Note that normally one would use the
' End statement to terminate a VB6 program abruptly. However the TWUtilities component interferes
' with the End statement's processing and prevents proper shutdown, so we use the
' TWUtilities component's EndProcess method instead. (However if we are running in the
' development environment, then we call End because EndProcess method kills the
' development environment as well which can have undesirable side effects if other
' components are also loaded.)

'If mIsInDev Then
'    mUnhandledErrorHandler.Handled = True
'    End
'Else
'    EndProcess
'End If

End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Function addStudyLibToList(ByVal pProgId As String) As Boolean
Const ProcName As String = "addStudyLibToList"
On Error GoTo Err

If Not isValidStudyLibrary(pProgId) Then Exit Function
    
StudyLibraryList.AddItem pProgId
StudyLibraryList.ListIndex = StudyLibraryList.ListCount - 1

addStudyLibraries

addStudyLibToList = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
Private Sub addStudyLibraries()
Const ProcName As String = "addStudyLibraries"
On Error GoTo Err

mStudyLibraryManager.RemoveAllStudyLibraries

mStudyLibraryManager.AddStudyLibrary BuiltInStudyLib, True, "Built-in"

Dim i As Long
For i = 0 To StudyLibraryList.ListCount - 1
    mStudyLibraryManager.AddStudyLibrary StudyLibraryList.List(i), True
Next

Dim lAvailableStudies() As StudyListEntry
lAvailableStudies = mStudyLibraryManager.GetAvailableStudies
For i = 0 To UBound(lAvailableStudies)
    StudiesCombo.AddItem lAvailableStudies(i).Name
Next

StudiesCombo.Enabled = True
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createBarNumberSeries() As TextSeries
Dim lPriceRegion As ChartRegion
Set lPriceRegion = Chart1.Regions.Item(PriceRegionName)
Set createBarNumberSeries = lPriceRegion.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText + 1)
createBarNumberSeries.align = AlignBottomCentre
createBarNumberSeries.Box = True
createBarNumberSeries.BoxFillColor = vbWhite
createBarNumberSeries.BoxFillStyle = FillSolid
createBarNumberSeries.BoxThickness = 1
createBarNumberSeries.FixedY = True
End Function

Private Function createBarsStudyConfig() As StudyConfiguration
Dim studyDef As StudyDefinition
Const ProcName As String = "createBarsStudyConfig"
On Error GoTo Err

ReDim InputValueNames(1) As String
Dim params As New Parameters
Dim studyValueConfig As StudyValueConfiguration
Dim barsStyle As BarStyle
Dim volumeStyle As DataPointStyle

Set studyDef = mStudyLibraryManager.GetStudyDefinition("Constant time bars")

Set createBarsStudyConfig = New StudyConfiguration
createBarsStudyConfig.ChartRegionName = PriceRegionName
InputValueNames(0) = InputValuePrice
InputValueNames(1) = InputValueVolume
createBarsStudyConfig.InputValueNames = InputValueNames
createBarsStudyConfig.Name = studyDef.Name
params.SetParameterValue "Bar length", 1
params.SetParameterValue "Time units", "Minutes"
createBarsStudyConfig.Parameters = params

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.Add("Bar")
studyValueConfig.ChartRegionName = PriceRegionName
studyValueConfig.IncludeInChart = True
studyValueConfig.Layer = 200
Set barsStyle = New BarStyle
barsStyle.OutlineThickness = 1
barsStyle.Thickness = 2
barsStyle.Width = 0.6
barsStyle.DisplayMode = BarDisplayModeCandlestick
barsStyle.DownColor = &H43FC2
barsStyle.SolidUpBody = True
barsStyle.TailThickness = 1
barsStyle.UpColor = &H1D9311
studyValueConfig.BarStyle = barsStyle

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.Add("Volume")
studyValueConfig.ChartRegionName = VolumeRegionName
studyValueConfig.IncludeInChart = True
Set volumeStyle = New DataPointStyle
volumeStyle.DownColor = vbRed
volumeStyle.UpColor = vbGreen
volumeStyle.DisplayMode = DataPointDisplayModeHistogram
volumeStyle.HistogramBarWidth = 0.7
volumeStyle.IncludeInAutoscale = True
volumeStyle.LineThickness = 1
studyValueConfig.DataPointStyle = volumeStyle

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub getBars(ByVal pFilename As String)
Dim lParser As New TestDataParser
lParser.ParseData pFilename
Set mBars = lParser.Bars
MinimumPriceTickText = lParser.MinimumPriceTick
mGrid.PriceFormatString = generatePriceFormatString(lParser.MinimumPriceTick)
End Sub


Private Function generatePriceFormatString(ByVal pMinPriceTick As Double) As String
Dim minTickString As String
Dim numberOfDecimals As Long

minTickString = Format(pMinPriceTick, "0.##############")

numberOfDecimals = Len(minTickString) - 2

If numberOfDecimals = 0 Then
    generatePriceFormatString = "0"
Else
    generatePriceFormatString = "0." & String(numberOfDecimals, "0")
End If

End Function

Private Sub getStudyValueAndValueMode( _
                ByVal pStudyConfig As StudyConfiguration, _
                ByVal pStudyValueName As String, _
                ByRef pStudyValue As SValue, _
                ByRef pValueMode As StudyValueModes)
Dim svd As StudyValueDefinition
Const ProcName As String = "getStudyValueAndValueMode"
On Error GoTo Err

Set svd = pStudyConfig.Study.StudyDefinition.StudyValueDefinitions.Item(pStudyValueName)
pStudyValue = pStudyConfig.Study.GetStudyValue(pStudyValueName, 0)

pValueMode = svd.ValueMode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getVolume(ByVal pBar As BarUtils27.Bar, ByVal pQuarter As Long) As Long
Const ProcName As String = "getVolume"
On Error GoTo Err

getVolume = Int(pQuarter * pBar.Volume / 4)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function

Private Sub initialiseChart()
Const ProcName As String = "initialiseChart"
On Error GoTo Err

Chart1.DisableDrawing

Chart1.ClearChart
Chart1.ChartBackColor = vbWhite
Chart1.PointerStyle = PointerCrosshairs
Chart1.HorizontalScrollBarVisible = True
Chart1.XAxisVisible = True

If Not mBars Is Nothing Then Chart1.TimePeriod = mBars.BarTimePeriod

Static defaultRegionStyle As ChartRegionStyle
If defaultRegionStyle Is Nothing Then Set defaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

Static volumeRegionStyle As ChartRegionStyle
If volumeRegionStyle Is Nothing Then
    Set volumeRegionStyle = defaultRegionStyle.Clone
    volumeRegionStyle.YGridlineSpacing = 0.8
    volumeRegionStyle.MinimumHeight = 10
    volumeRegionStyle.IntegerYScale = True
    'volumeRegionStyle.YScaleQuantum = 1#
End If

Dim priceRegion As ChartRegion
Set priceRegion = Chart1.Regions.Add(100, 25, defaultRegionStyle, , PriceRegionName)
priceRegion.Title.Text = TestDataFilenameText
priceRegion.Title.Color = vbBlue

Dim volumeRegion As ChartRegion
Set volumeRegion = Chart1.Regions.Add(20, , volumeRegionStyle, , VolumeRegionName)
volumeRegion.Title.Text = "Volume"
volumeRegion.Title.Color = vbBlue

Chart1.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isValidStudyLibrary(ByVal pProgId As String) As Boolean
Dim lStudyLibrary As IStudyLibrary
On Error Resume Next
Set lStudyLibrary = CreateObject(pProgId)
isValidStudyLibrary = Not (lStudyLibrary Is Nothing)
End Function

Private Sub loadBarToChart(ByVal pBar As BarUtils27.Bar)
Const ProcName As String = "loadBarToChart"
Dim failPoint As String
On Error GoTo Err

failPoint = "notifying open value for bar " & pBar.barNumber
notifyPrice pBar.OpenValue, pBar.TimeStamp

failPoint = "notifying volume at open for bar " & pBar.barNumber
NotifyVolume mAccumulatedVolume + Int(pBar.Volume / 4), pBar.TimeStamp
        
failPoint = "notifying low value for bar " & pBar.barNumber
notifyPrice pBar.LowValue, pBar.TimeStamp

failPoint = "notifying volume at high for bar " & pBar.barNumber
NotifyVolume mAccumulatedVolume + Int(2 * pBar.Volume / 4), pBar.TimeStamp
        
failPoint = "notifying high value for bar " & pBar.barNumber
notifyPrice pBar.HighValue, pBar.TimeStamp

failPoint = "notifying volume at low for bar " & pBar.barNumber
NotifyVolume mAccumulatedVolume + Int(3 * pBar.Volume / 4), pBar.TimeStamp
            
failPoint = "notifying close value for bar " & pBar.barNumber
notifyPrice pBar.CloseValue, pBar.TimeStamp

failPoint = "notifying volume at close for bar " & pBar.barNumber
NotifyVolume mAccumulatedVolume + pBar.Volume, pBar.TimeStamp

mAccumulatedVolume = mAccumulatedVolume + pBar.Volume

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadBarToGrid(ByVal pBar As BarUtils27.Bar)
Const ProcName As String = "loadBarToGrid"
On Error GoTo Err

mGrid.Row = pBar.barNumber

mGrid.SetCellLong TestDataGridColumns.barNumber, CStr(pBar.barNumber)

mGrid.SetCellDate TestDataGridColumns.TimeStamp, CStr(pBar.TimeStamp)

mGrid.SetCellPrice TestDataGridColumns.OpenValue, CStr(pBar.OpenValue)
mGrid.SetCellPrice TestDataGridColumns.HighValue, CStr(pBar.HighValue)
mGrid.SetCellPrice TestDataGridColumns.LowValue, CStr(pBar.LowValue)
mGrid.SetCellPrice TestDataGridColumns.CloseValue, CStr(pBar.CloseValue)

mGrid.SetCellLong TestDataGridColumns.Volume, CStr(pBar.Volume)
mGrid.SetCellLong TestDataGridColumns.OpenInterest, CStr(pBar.OpenInterest)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadBarsToGrid()
Const ProcName As String = "loadBarsToGrid"
On Error GoTo Err

Dim lBar As BarUtils27.Bar

mGrid.Redraw = False
mGrid.SetupDataColumns

mAccumulatedVolume = 0
For Each lBar In mBars
    loadBarToGrid lBar
    
    loadBarToChart lBar
    
    showBarNumber lBar.barNumber
    
Next

mGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LoadData()
Const ProcName As String = "LoadData"
On Error GoTo Err

TestButton.Enabled = False

Screen.MousePointer = MousePointerConstants.vbArrowHourglass

addStudyLibraries

getBars TestDataFilenameText

setupInitialStudies

Chart1.DisableDrawing

loadBarsToGrid

Chart1.FirstVisiblePeriod = 1
Chart1.EnableDrawing

Chart1.Regions.Item(PriceRegionName).YScaleQuantum = CDbl(MinimumPriceTickText)

Screen.MousePointer = MousePointerConstants.vbDefault

If mIsStudySet Then TestButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logErrorsToForm()
GetLogger("error").AddLogListener Me
End Sub

Private Sub notifyPrice(pPrice As Double, pTimestamp As Date)
Const ProcName As String = "notifyPrice"
On Error GoTo Err

mStudyManager.NotifyInput _
                mPriceInputHandle, _
                pPrice, _
                pTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyVolume(ByVal pVolume As Long, pTimestamp As Date)
Const ProcName As String = "NotifyVolume"
On Error GoTo Err

mStudyManager.NotifyInput _
                mVolumeInputHandle, _
                pVolume, _
                pTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function studyValueToString( _
                ByRef pStudyValue As SValue, _
                ByVal pValueMode As StudyValueModes) As String
                
Const ProcName As String = "studyValueToString"
Dim failPoint As String
On Error GoTo Err

Dim lObj As IStringable

Select Case pValueMode
Case ValueModeNone
    studyValueToString = CStr(pStudyValue.value)
Case Else
    Set lObj = pStudyValue.value
    If Not lObj Is Nothing Then studyValueToString = lObj.ToString
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processStudyValues( _
                ByVal pStudyConfig As StudyConfiguration, _
                ByVal pBarNumber As Long)
Const ProcName As String = "processStudyValues"
Dim failPoint As String
On Error GoTo Err

Dim i As Long
Dim j As Long
Dim svc As StudyValueConfiguration
Dim lStudyValue As SValue
Dim lValueMode As StudyValueModes

For i = 1 To pStudyConfig.StudyValueConfigurations.Count
    Set svc = pStudyConfig.StudyValueConfigurations.Item(i)
    If svc.IncludeInChart Then
        failPoint = "getting value for '" & svc.ValueName & "' for bar " & pBarNumber
        mGrid.Row = pBarNumber
        getStudyValueAndValueMode pStudyConfig, svc.ValueName, lStudyValue, lValueMode
        mGrid.SetCell TestDataGridColumns.StudyValue1 + j, studyValueToString(lStudyValue, lValueMode)
        j = j + 1
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pFailpoint:=failPoint
End Sub

Private Sub setupInitialStudies()
Const ProcName As String = "setupInitialStudies"
On Error GoTo Err

initialiseChart

Set mChartManager = CreateChartManager(Chart1.Controller, mStudyManager, mBarFormatterLibManager, True)

Set mSourceStudy = mStudyManager.CreateStudyInputHandler(IIf(TestDataFilenameText = "", _
                                                "Test data", _
                                                TestDataFilenameText))

mPriceInputHandle = mStudyManager.AddInput(mSourceStudy, _
                        InputValuePrice, _
                        ChartUtils27.ChartRegionNamePrice, _
                        InputTypeReal, _
                        True, _
                        MinimumPriceTickText)
mChartManager.SetInputRegion mPriceInputHandle, PriceRegionName

mVolumeInputHandle = mStudyManager.AddInput(mSourceStudy, _
                        InputValueVolume, _
                        ChartUtils27.ChartRegionNameVolume, _
                        InputTypeInteger, _
                        False, _
                        1)
mChartManager.SetInputRegion mVolumeInputHandle, VolumeRegionName

Dim studyConfig As StudyConfiguration
Set studyConfig = createBarsStudyConfig
studyConfig.UnderlyingStudy = mSourceStudy
Set mBarsStudy = mStudyManager.AddStudy(studyConfig.Name, mSourceStudy, studyConfig.InputValueNames, True, studyConfig.Parameters, studyConfig.StudyLibraryName)
studyConfig.Study = mBarsStudy
mChartManager.StartStudy mBarsStudy

mChartManager.BaseStudyConfiguration = studyConfig

Set mInitialStudyConfigs = New StudyConfigurations
mInitialStudyConfigs.Add mChartManager.BaseStudyConfiguration

ChartToolbar1.Initialise Chart1.Controller, Chart1.Regions(PriceRegionName), mChartManager.BaseStudyConfiguration.ValueSeries("Bar")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showBarNumber(ByVal pRow As Long)
Const ProcName As String = "showBarNumber"
On Error GoTo Err

Static lBarNumbers As TextSeries
Dim lBarNumberText As Text

If pRow = 1 Then Set lBarNumbers = createBarNumberSeries

If pRow Mod 10 = 0 Then
    Set lBarNumberText = lBarNumbers.Add
    lBarNumberText.Text = CStr(pRow)
    lBarNumberText.Position = NewPoint(pRow, 0.2, CoordsLogical, CoordsDistance)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showError(ByVal pProcName As String, ByVal pFailpoint As String)
MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
        "At:" & vbCrLf & _
        IIf(Err.Source <> "", Err.Source, "") & vbCrLf & _
        ProjectName & "." & ModuleName & ":" & pProcName & " At: " & pFailpoint
End Sub

Private Sub stopLoggingErrorsToForm()
GetLogger("error").AddLogListener Me
End Sub

Private Sub testStudy(ByVal pStudyConfig As StudyConfiguration)
Const ProcName As String = "testStudy"
Dim failPoint As String
On Error GoTo Err

Dim lBar As BarUtils27.Bar

logErrorsToForm
Chart1.DisableDrawing

mAccumulatedVolume = 0
For Each lBar In mBars
    failPoint = "processing bar " & lBar.barNumber
    loadBarToChart lBar
    
    failPoint = "adding bar number"
    showBarNumber lBar.barNumber
    
    failPoint = "getting study values for bar " & lBar.barNumber
    processStudyValues pStudyConfig, lBar.barNumber
Next

Chart1.EnableDrawing
stopLoggingErrorsToForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
