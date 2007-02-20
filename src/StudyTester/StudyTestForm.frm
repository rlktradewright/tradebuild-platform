VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CA028305-9CAA-44AE-816E-330E4FEBE823}#2.0#0"; "StudiesUI2-5.ocx"
Object = "{015212C3-04F2-4693-B20B-0BEB304EFC1B}#2.0#0"; "ChartSkil2-5.ocx"
Begin VB.Form StudyTestForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TradeBuild Study Test Harness"
   ClientHeight    =   10365
   ClientLeft      =   5070
   ClientTop       =   3540
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   12510
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
      Left            =   11280
      TabIndex        =   11
      ToolTipText     =   "Test the study"
      Top             =   120
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   16325
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   2
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
      TabCaption(1)   =   "Study setup"
      TabPicture(1)   =   "StudyTestForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(4)=   "StudiesCombo"
      Tab(1).Control(5)=   "StudyLibraryClassNameText"
      Tab(1).Control(6)=   "LibToAddText"
      Tab(1).Control(7)=   "AddLibButton"
      Tab(1).Control(8)=   "SpList"
      Tab(1).Control(9)=   "RemoveLibButton"
      Tab(1).Control(10)=   "SetStudyLibraryButton"
      Tab(1).Control(11)=   "StudyConfigurer1"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&Chart"
      TabPicture(2)   =   "StudyTestForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Chart1"
      Tab(2).ControlCount=   1
      Begin ChartSkil25.Chart Chart1 
         Height          =   8775
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15478
         autoscale       =   0   'False
      End
      Begin StudiesUI25.StudyConfigurer StudyConfigurer1 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   19
         Top             =   3480
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   9975
      End
      Begin VB.TextBox MinimumPriceTickText 
         Height          =   285
         Left            =   9480
         TabIndex        =   2
         Text            =   "0.0"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton SetStudyLibraryButton 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   4
         ToolTipText     =   "Click to load your service provider"
         Top             =   540
         Width           =   855
      End
      Begin VB.CommandButton RemoveLibButton 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   9
         ToolTipText     =   "Remove the selected service provider from the list"
         Top             =   2340
         Width           =   855
      End
      Begin VB.ListBox SpList 
         Height          =   840
         ItemData        =   "StudyTestForm.frx":0054
         Left            =   -72600
         List            =   "StudyTestForm.frx":0056
         TabIndex        =   8
         ToolTipText     =   "Lists all studies service providers you need (except the built-in studies service provider)"
         Top             =   2340
         Width           =   3975
      End
      Begin VB.CommandButton AddLibButton 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68640
         TabIndex        =   7
         ToolTipText     =   "Add this service provider to the list"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox LibToAddText 
         Height          =   285
         Left            =   -72600
         TabIndex        =   6
         ToolTipText     =   "Enter the program id of any other studies service provider your service provider needs"
         Top             =   1860
         Width           =   3975
      End
      Begin VB.TextBox StudyLibraryClassNameText 
         Height          =   285
         Left            =   -72600
         TabIndex        =   3
         ToolTipText     =   "Enter your service provider's program id in the form project.class"
         Top             =   540
         Width           =   3975
      End
      Begin VB.ComboBox StudiesCombo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72600
         TabIndex        =   5
         ToolTipText     =   "Select the study to test"
         Top             =   1020
         Width           =   3975
      End
      Begin VB.CommandButton FindFileButton 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         TabIndex        =   1
         ToolTipText     =   "Click to browse for the test data file"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox TestDataFilenameText 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "The file that contains the test data"
         Top             =   840
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid TestDataGrid 
         Height          =   7935
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1260
         Width           =   12015
         _ExtentX        =   21193
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
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Configure the study - selected output values will appear both on the chart and in the grid"
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   3240
         Width           =   11655
      End
      Begin VB.Label Label1 
         Caption         =   "Other study libraries to include"
         Height          =   615
         Left            =   -74760
         TabIndex        =   17
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Test data file"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Program id for Study Library under test"
         Height          =   375
         Left            =   -74760
         TabIndex        =   15
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Study to test"
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   1080
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

                                    
'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const CellBackColorOdd As Long = &HF8F8F8
Private Const CellBackColorEven As Long = &HEEEEEE

Private Const PriceRegionName As String = "$price"
Private Const VolumeRegionName As String = "$volume"

Private Const TestDataGridRowsInitial As Long = 50
Private Const TestDataGridRowsIncrement As Long = 25

Private Const InputValuePrice As String = "Price"
Private Const InputValueVolume As String = "Volume"

'================================================================================
' Enums
'================================================================================

Private Enum TestDataFileColumns
    timestamp
    openValue
    highValue
    lowValue
    closeValue
    Volume
End Enum

Private Enum TestDataGridColumns
    timestamp
    openValue
    highValue
    lowValue
    closeValue
    Volume
    StudyValue1
End Enum

' Character widths of the TestDataGrid columns
Private Enum TestDataGridColumnWidths
    TimeStampWidth = 19
    openValueWidth = 9
    highValueWidth = 9
    lowValueWidth = 9
    closeValueWidth = 9
    volumeWidth = 9
    StudyValue1Width = 20
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

Private mLetterWidth As Single
Private mDigitWidth As Single

Private mStudyParams As Parameters

Private mIsDataLoaded As Boolean
Private mIsStudySet As Boolean

Private mName As String

Private mStudyLibrary As StudyLibrary

Private mStudyDefinition As StudyDefinition

Private mChartManager As ChartManager
Private mStudyManager As StudyManager
Private mSourceStudy As study
Private mBarsStudy As study
Private mInitialStudyConfigs As StudyConfigurations

Private mPriceInputHandle As Long
Private mVolumeInputHandle As Long

Private mPeriodLength As Long
Private mPeriodUnits As TimePeriodUnits

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls  ' enables WinXP look and feel
End Sub

Private Sub Form_Load()
Dim widthString As String

mName = "TradeBuild Study Test Harness"

widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = Me.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = Me.TextWidth(widthString) / Len(widthString)

setupTestDataGrid

AddStudyLibrary New CmnStudiesLib25.StudyLib

' need to do this in case the user sets up his Study Library and study
' before loading the test data
setupInitialStudies

End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddLibButton_Click()
SpList.AddItem LibToAddText
SpList.ListIndex = SpList.ListCount - 1
LibToAddText = ""
End Sub

Private Sub FindFileButton_Click()

CommonDialog1.CancelError = True
On Error GoTo err

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

LoadData

err:
End Sub

Private Sub MinimumPriceTickText_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then Exit Sub
If Not IsNumeric(MinimumPriceTickText & Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub RemoveLibButton_Click()
SpList.RemoveItem SpList.ListIndex
RemoveLibButton.Enabled = False
End Sub

Private Sub StudyLibraryClassNameText_Change()
If StudyLibraryClassNameText = "" Then
    SetStudyLibraryButton.Enabled = False
Else
    SetStudyLibraryButton.Enabled = True
End If
End Sub

Private Sub SetStudyLibraryButton_Click()
Dim availableStudies() As String
Dim i As Long

StudiesCombo.Clear
StudyConfigurer1.Clear
Set mStudyDefinition = Nothing
mIsStudySet = False
TestButton.Enabled = False
TestDataGrid.Cols = TestDataGridColumns.StudyValue1

If StudyLibraryClassNameText = "" Then
    StudiesCombo.Enabled = False
    Exit Sub
End If

On Error Resume Next
Set mStudyLibrary = CreateObject(StudyLibraryClassNameText)
On Error GoTo 0
If mStudyLibrary Is Nothing Then
    StudiesCombo.Enabled = False
    MsgBox StudyLibraryClassNameText & " is not a valid Study Library"
    Exit Sub
End If

RemoveAllStudyLibraries
AddStudyLibrary mStudyLibrary

StudiesCombo.Enabled = True
availableStudies = mStudyLibrary.getImplementedStudyNames
For i = 0 To UBound(availableStudies)
    StudiesCombo.AddItem availableStudies(i)
Next

End Sub

Private Sub SpList_Click()
If SpList.ListIndex = -1 Then
    RemoveLibButton.Enabled = False
Else
    RemoveLibButton.Enabled = True
End If
End Sub

Private Sub LibToAddText_Change()
If LibToAddText = "" Then
    AddLibButton.Enabled = False
Else
    AddLibButton.Enabled = True
End If
End Sub

Private Sub StudiesCombo_Click()
Dim regionNames(1) As String

setupInitialStudies
                    
Set mStudyParams = getStudyDefaultParameters(StudiesCombo)

regionNames(0) = PriceRegionName
regionNames(1) = VolumeRegionName

Set mStudyDefinition = getStudyDefinition(StudiesCombo)

StudyConfigurer1.initialise mStudyDefinition, _
                            "", _
                            regionNames, _
                            mInitialStudyConfigs, _
                            Nothing, _
                            mStudyParams
mIsStudySet = True
If mIsDataLoaded Then TestButton.Enabled = True
End Sub

Private Sub TestButton_Click()
Dim i As Long
Dim when As String
Dim volumeThisBar As Long
Dim timestamp As Date
Dim testStudy As study
Dim testStudyConfig As StudyConfiguration
Dim accumVolume As Long
Dim addTestStudyToSource As Boolean
Dim regionNames(1) As String

On Error GoTo err
Screen.MousePointer = MousePointerConstants.vbArrowHourglass

when = "adding study libraries"
addStudyLibraries

Set testStudyConfig = StudyConfigurer1.StudyConfiguration
If testStudyConfig.underlyingStudy Is mSourceStudy Then
    addTestStudyToSource = True
End If

when = "creating the study to be tested"

setupInitialStudies

If addTestStudyToSource Then
    testStudyConfig.underlyingStudy = mSourceStudy
Else
    testStudyConfig.underlyingStudy = mBarsStudy
End If

Set testStudy = mChartManager.addStudy(testStudyConfig)
mChartManager.startStudy testStudy

' now re-setup the study configurer so that only current
' objects are referenced
regionNames(0) = PriceRegionName
regionNames(1) = VolumeRegionName
StudyConfigurer1.initialise mStudyDefinition, _
                            "", _
                            regionNames, _
                            mInitialStudyConfigs, _
                            testStudyConfig, _
                            mStudyParams

when = "setting up the Study Value grid"
setupStudyValueGridColumns testStudyConfig

mChartManager.suppressDrawing = True
For i = 1 To TestDataGrid.Rows
    TestDataGrid.row = i
    TestDataGrid.Col = TestDataGridColumns.timestamp
    If TestDataGrid.Text = "" Then Exit For
    timestamp = CDate(TestDataGrid.Text)
    
    If TestDataGrid.TextMatrix(i, TestDataGridColumns.Volume) <> "" Then
        volumeThisBar = CLng(TestDataGrid.TextMatrix(i, TestDataGridColumns.Volume))
    Else
        volumeThisBar = 0
    End If
    
    when = "notifying open value for bar " & i
    mStudyManager.notifyInput _
                    mPriceInputHandle, _
                    CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.openValue)), _
                    timestamp

    If volumeThisBar <> 0 Then
        when = "notifying volume at open for bar " & i
        accumVolume = accumVolume + Int(volumeThisBar / 4)
        mStudyManager.notifyInput _
                        mVolumeInputHandle, _
                        accumVolume, _
                        timestamp
    End If
            
    when = "notifying high value for bar " & i
    mStudyManager.notifyInput _
                    mPriceInputHandle, _
                    CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.highValue)), _
                    timestamp

    If volumeThisBar <> 0 Then
        when = "notifying volume at high for bar " & i
        accumVolume = accumVolume + Int(volumeThisBar / 4)
        mStudyManager.notifyInput _
                        mVolumeInputHandle, _
                        accumVolume, _
                        timestamp
    End If
            
    when = "notifying low value for bar " & i
    mStudyManager.notifyInput _
                    mPriceInputHandle, _
                    CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.lowValue)), _
                    timestamp

    If volumeThisBar <> 0 Then
        when = "notifying volume at low for bar " & i
        accumVolume = accumVolume + Int(volumeThisBar / 4)
        mStudyManager.notifyInput _
                        mVolumeInputHandle, _
                        accumVolume, _
                        timestamp
    End If
            
    when = "notifying close value for bar " & i
    mStudyManager.notifyInput _
                    mPriceInputHandle, _
                    CDbl(TestDataGrid.TextMatrix(i, TestDataGridColumns.closeValue)), _
                    timestamp

    If volumeThisBar <> 0 Then
        when = "notifying volume at low for bar " & i
        accumVolume = accumVolume + volumeThisBar - 3 * Int(volumeThisBar / 4)
        mStudyManager.notifyInput _
                        mVolumeInputHandle, _
                        accumVolume, _
                        timestamp
    End If
    
    processStudyValues testStudy, testStudyConfig, i, when
Next

mChartManager.suppressDrawing = False

setTestDataGridRowBackColors 1

Screen.MousePointer = MousePointerConstants.vbDefault
Exit Sub

err:
setTestDataGridRowBackColors 1

Do Until Not Chart1.suppressDrawing
    Chart1.suppressDrawing = False
Loop

MsgBox "Error " & err.Number & _
        " when " & when & _
        ": " & err.Description & _
        IIf(err.Source <> "", ": " & err.Source, "")
Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

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

Private Sub addStudyLibraries()
Dim i As Long
Dim standardLib As Object

RemoveAllStudyLibraries

Set standardLib = New CmnStudiesLib25.StudyLib
AddStudyLibrary standardLib

For i = 0 To SpList.ListCount - 1
    AddStudyLibrary CreateObject(SpList.List(i))
Next

If Not mStudyLibrary Is Nothing Then
    If Not TypeOf mStudyLibrary Is CmnStudiesLib25.StudyLib Then
        AddStudyLibrary mStudyLibrary
    End If
End If
End Sub

Private Function createBarsStudyConfig() As StudyConfiguration
Dim studyDef As StudyDefinition
ReDim inputValueNames(1) As String
Dim params As New Parameters2.Parameters
Dim studyValueConfig As StudyValueConfiguration

Set studyDef = getStudyDefinition("Constant time bars")

Set createBarsStudyConfig = New StudyConfiguration
createBarsStudyConfig.chartRegionName = PriceRegionName
inputValueNames(0) = InputValuePrice
inputValueNames(1) = InputValueVolume
createBarsStudyConfig.inputValueNames = inputValueNames
createBarsStudyConfig.Name = studyDef.Name
params.setParameterValue "Bar length", 1
params.setParameterValue "Time units", "Minutes"
createBarsStudyConfig.Parameters = params
'createBarsStudyConfig.StudyDefinition = studyDef

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.Add("Bar")
studyValueConfig.outlineThickness = 1
studyValueConfig.barThickness = 2
studyValueConfig.barWidth = 0.75
studyValueConfig.chartRegionName = PriceRegionName
studyValueConfig.barDisplayMode = BarDisplayModeCandlestick
studyValueConfig.downColor = &H43FC2
studyValueConfig.includeInAutoscale = True
studyValueConfig.includeInChart = True
studyValueConfig.layer = 200
studyValueConfig.solidUpBody = True
studyValueConfig.tailThickness = 1
studyValueConfig.upColor = &H1D9311

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.Add("Volume")
studyValueConfig.chartRegionName = VolumeRegionName
studyValueConfig.Color = vbBlack
studyValueConfig.dataPointDisplayMode = DataPointDisplayModeHistogram
studyValueConfig.histogramBarWidth = 0.7
studyValueConfig.includeInAutoscale = True
studyValueConfig.includeInChart = True
studyValueConfig.lineThickness = 1
End Function

Private Sub determinePeriodParameters()
Dim fso As FileSystemObject
Dim ts As TextStream
Dim row As Long
Dim timestamp1 As Date
Dim timestamp2 As Date
Dim rec As String
Dim tokens() As String

Set fso = New FileSystemObject
Set ts = fso.OpenTextFile(TestDataFilenameText, ForReading)

Do While Not ts.AtEndOfStream
    rec = ts.ReadLine
    If rec <> "" And Left$(rec, 2) <> "//" Then
        row = row + 1
        tokens = Split(rec, ",")
        
        If row = 1 Then
            timestamp1 = CDate(tokens(TestDataFileColumns.timestamp))
        Else
            timestamp2 = CDate(tokens(TestDataFileColumns.timestamp))
            
            mPeriodUnits = TimePeriodSecond
            mPeriodLength = DateDiff("s", timestamp1, timestamp2)
            If mPeriodLength < 60 Then Exit Sub
            
            mPeriodUnits = TimePeriodMinute
            mPeriodLength = DateDiff("n", timestamp1, timestamp2)
            If mPeriodLength < 60 Then Exit Sub
            
            mPeriodUnits = TimePeriodHour
            mPeriodLength = DateDiff("h", timestamp1, timestamp2)
            If mPeriodLength < 24 Then Exit Sub
            
            mPeriodUnits = TimePeriodDay
            mPeriodLength = DateDiff("d", timestamp1, timestamp2)
            If mPeriodLength < 5 Then Exit Sub
            
            mPeriodUnits = TimePeriodWeek
            mPeriodLength = DateDiff("ww", timestamp1, timestamp2)
            If mPeriodLength < 5 Then Exit Sub
            
            mPeriodUnits = TimePeriodMonth
            mPeriodLength = DateDiff("m", timestamp1, timestamp2)
            If mPeriodLength < 12 Then Exit Sub
            
            mPeriodUnits = TimePeriodYear
            mPeriodLength = DateDiff("yyyy", timestamp1, timestamp2)
            Exit Sub
            
        End If
    End If
Loop
End Sub

Private Sub initialiseChart( _
                ByVal pChartManager As ChartManager)
Dim priceRegion As ChartRegion
Dim volumeRegion As ChartRegion


pChartManager.suppressDrawing = True

pChartManager.clearChart
pChartManager.chartController.chartBackColor = vbWhite
pChartManager.chartController.autoscale = True
pChartManager.chartController.pointerStyle = PointerCrosshairs
pChartManager.chartController.twipsPerBar = 100
pChartManager.chartController.showHorizontalScrollBar = True
pChartManager.chartController.setPeriodParameters mPeriodLength, mPeriodUnits

Set priceRegion = pChartManager.addChartRegion(PriceRegionName, 100, 25)
priceRegion.gridlineSpacingY = 2
priceRegion.showGrid = True
priceRegion.setTitle TestDataFilenameText, vbBlue, Nothing

Set volumeRegion = pChartManager.addChartRegion(VolumeRegionName, 20)
volumeRegion.gridlineSpacingY = 0.8
volumeRegion.minimumHeight = 10
volumeRegion.integerYScale = True
volumeRegion.showGrid = True
volumeRegion.setTitle "Volume", vbBlue, Nothing

pChartManager.suppressDrawing = False

End Sub

Private Sub LoadData()
Dim fso As FileSystemObject
Dim ts As TextStream
Dim rec As String
Dim tokens() As String
Dim row As Long
Dim accumVolume As Long
Dim timestamp As Date

On Error GoTo err

mIsDataLoaded = False
TestButton.Enabled = False

Screen.MousePointer = MousePointerConstants.vbArrowHourglass

TestDataGrid.Clear
setupTestDataGrid
TestDataGrid.Refresh
TestDataGrid.Redraw = False

addStudyLibraries

determinePeriodParameters

setupInitialStudies

Set fso = New FileSystemObject
Set ts = fso.OpenTextFile(TestDataFilenameText, ForReading)

Chart1.suppressDrawing = True

Do While Not ts.AtEndOfStream
    rec = ts.ReadLine
    If rec <> "" And Left$(rec, 2) <> "//" Then
        row = row + 1
        tokens = Split(rec, ",")
        
        'update the chart
        
        timestamp = CDate(tokens(TestDataFileColumns.timestamp))
        
        mStudyManager.notifyInput mPriceInputHandle, _
                        CDbl(tokens(TestDataFileColumns.openValue)), _
                        timestamp
        
        mStudyManager.notifyInput mPriceInputHandle, _
                        CDbl(tokens(TestDataFileColumns.highValue)), _
                        timestamp
        
        mStudyManager.notifyInput mPriceInputHandle, _
                        CDbl(tokens(TestDataFileColumns.lowValue)), _
                        timestamp
        
        mStudyManager.notifyInput mPriceInputHandle, _
                        CDbl(tokens(TestDataFileColumns.closeValue)), _
                        timestamp
        
        If tokens(TestDataFileColumns.Volume) <> "" Then
            accumVolume = accumVolume + CLng(tokens(TestDataFileColumns.Volume))
            mChartManager.notifyInput mVolumeInputHandle, _
                        accumVolume, _
                        timestamp
        End If
        
        'update the grid
        If row > TestDataGrid.Rows - 1 Then TestDataGrid.Rows = TestDataGrid.Rows + TestDataGridRowsIncrement
        TestDataGrid.row = row
        TestDataGrid.Col = TestDataGridColumns.timestamp
        TestDataGrid.Text = CDate(tokens(TestDataFileColumns.timestamp))
        TestDataGrid.Col = TestDataGridColumns.openValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.openValue))
        TestDataGrid.Col = TestDataGridColumns.highValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.highValue))
        TestDataGrid.Col = TestDataGridColumns.lowValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.lowValue))
        TestDataGrid.Col = TestDataGridColumns.closeValue
        TestDataGrid.Text = CDbl(tokens(TestDataFileColumns.closeValue))
        If tokens(TestDataFileColumns.Volume) <> "" Then
            TestDataGrid.Col = TestDataGridColumns.Volume
            TestDataGrid.Text = CLng(tokens(TestDataFileColumns.Volume))
        End If
    End If
    
Loop

TestDataGrid.Redraw = True
setTestDataGridRowBackColors 1
Chart1.suppressDrawing = False

Screen.MousePointer = MousePointerConstants.vbDefault

mIsDataLoaded = True
If mIsStudySet Then TestButton.Enabled = True

Exit Sub

err:
TestDataGrid.Redraw = True
setTestDataGridRowBackColors 1
Chart1.suppressDrawing = False

Screen.MousePointer = MousePointerConstants.vbDefault

MsgBox "Can't load data file: " & TestDataFilenameText & vbCrLf & _
        "Error " & err.Number & ": " & err.Description
End Sub

Private Sub processStudyValues( _
                ByVal study As study, _
                ByVal studyConfig As StudyConfiguration, _
                ByVal row As Long, _
                ByRef when As String)
Dim svd As StudyValueDefinition
Dim svc As StudyValueConfiguration
Dim lStudyValue As StudyValue
Dim i As Long
Dim j As Long
Dim lLine As StudyLine
Dim lBar As Bar
Dim lText As StudyText

For i = 1 To studyConfig.StudyValueConfigurations.Count
    Set svc = studyConfig.StudyValueConfigurations.Item(i)
    If svc.includeInChart Then
        Set svd = studyConfig.study.StudyDefinition.StudyValueDefinitions.Item(svc.valueName)
        when = "getting value for " & svc.valueName & " for bar " & row
        lStudyValue = study.getStudyValue(svc.valueName, 0)
        
        Select Case svd.valueMode
        Case ValueModeNone
            TestDataGrid.TextMatrix(row, TestDataGridColumns.StudyValue1 + j) = lStudyValue.Value
        Case ValueModeLine
            Set lLine = lStudyValue.Value
            If Not lLine Is Nothing Then
                TestDataGrid.TextMatrix(row, TestDataGridColumns.StudyValue1 + j) = _
                        "(" & lLine.point1.x & "," & lLine.point1.y & ")-" & _
                        "(" & lLine.point2.x & "," & lLine.point2.y & ")"
            End If
        Case ValueModeBar
            Set lBar = lStudyValue.Value
            If Not lBar Is Nothing Then
                TestDataGrid.TextMatrix(row, TestDataGridColumns.StudyValue1 + j) = _
                        lBar.openValue & "," & _
                        lBar.highValue & "," & _
                        lBar.lowValue & "," & _
                        lBar.closeValue
            End If
        Case ValueModeText
            Set lText = lStudyValue.Value
            If Not lText Is Nothing Then
                TestDataGrid.TextMatrix(row, TestDataGridColumns.StudyValue1 + j) = _
                        "(" & lText.position.x & "," & lText.position.y & ")," & _
                        """" & lText.Text & """"
            End If
        End Select
        
        j = j + 1
    End If
Next

End Sub

Private Sub setTestDataGridRowBackColors( _
                ByVal startingIndex As Long)
Dim i As Long

TestDataGrid.Redraw = False

For i = startingIndex To TestDataGrid.Rows - 1
    TestDataGrid.row = i
    TestDataGrid.Col = 0
    TestDataGrid.RowSel = i
    TestDataGrid.ColSel = TestDataGrid.Cols - 1
    TestDataGrid.CellBackColor = IIf(i Mod 2 = 0, CellBackColorEven, CellBackColorOdd)
    
Next

TestDataGrid.Redraw = True
End Sub

Private Sub setupInitialStudies()
Dim studyConfig As StudyConfiguration

Set mStudyManager = New StudyManager
Set mChartManager = createChartManager(mStudyManager, Chart1.controller)

initialiseChart mChartManager

Set mSourceStudy = mStudyManager.addSource(IIf(TestDataFilenameText = "", _
                                                "Test data", _
                                                TestDataFilenameText))

mPriceInputHandle = mStudyManager.addInput(mSourceStudy, _
                        InputValuePrice, _
                        "Price", _
                        InputTypeReal, _
                        True, _
                        MinimumPriceTickText)
mChartManager.setInputRegion mPriceInputHandle, PriceRegionName

mVolumeInputHandle = mStudyManager.addInput(mSourceStudy, _
                        InputValueVolume, _
                        "Volume", _
                        InputTypeInteger, _
                        False, _
                        1)
mChartManager.setInputRegion mVolumeInputHandle, VolumeRegionName

Set studyConfig = createBarsStudyConfig
studyConfig.underlyingStudy = mSourceStudy
Set mBarsStudy = mChartManager.addStudy(studyConfig)
mChartManager.startStudy mBarsStudy


Set mInitialStudyConfigs = New StudyConfigurations
For Each studyConfig In mChartManager.StudyConfigurations
    mInitialStudyConfigs.Add studyConfig
Next
End Sub

Private Sub setupStudyValueGridColumns( _
                ByVal studyConfig As StudyConfiguration)
Dim svd As StudyValueDefinition
Dim svc As StudyValueConfiguration
Dim i As Long
Dim j As Long

' remove any existing study value columns
TestDataGrid.Cols = TestDataGridColumns.StudyValue1

For i = 1 To studyConfig.StudyValueConfigurations.Count
    Set svc = studyConfig.StudyValueConfigurations.Item(i)
    If svc.includeInChart Then
        Set svd = studyConfig.study.StudyDefinition.StudyValueDefinitions.Item(svc.valueName)
        setupTestDataGridColumn TestDataGridColumns.StudyValue1 + j, _
                                TestDataGridColumnWidths.StudyValue1Width, _
                                svd.Name, _
                                IIf(svd.valueType = ValueTypeString, True, False), _
                                IIf(svd.valueType = ValueTypeString, AlignmentSettings.flexAlignLeftCenter, AlignmentSettings.flexAlignRightCenter)
        j = j + 1
    End If
Next

End Sub

Private Sub setupTestDataGrid()

With TestDataGrid
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .HighLight = flexHighlightNever
    
    .Cols = TestDataGridColumns.StudyValue1
    .Rows = TestDataGridRowsInitial
    .FixedRows = 1
    .FixedCols = 0
End With
    
setupTestDataGridColumn TestDataGridColumns.timestamp, TestDataGridColumnWidths.TimeStampWidth, "Timestamp", False, AlignmentSettings.flexAlignLeftCenter
setupTestDataGridColumn TestDataGridColumns.openValue, TestDataGridColumnWidths.openValueWidth, "Open", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.highValue, TestDataGridColumnWidths.highValueWidth, "High", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.lowValue, TestDataGridColumnWidths.lowValueWidth, "Low", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.closeValue, TestDataGridColumnWidths.closeValueWidth, "Close", False, AlignmentSettings.flexAlignRightCenter
setupTestDataGridColumn TestDataGridColumns.Volume, TestDataGridColumnWidths.volumeWidth, "Volume", False, AlignmentSettings.flexAlignRightCenter

setTestDataGridRowBackColors 1
End Sub

Private Sub setupTestDataGridColumn( _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As AlignmentSettings)
    
Dim lColumnWidth As Long

With TestDataGrid
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .ColWidth(columnNumber) = 0
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    .ColWidth(columnNumber) = lColumnWidth
    
    .ColAlignment(columnNumber) = align
    .TextMatrix(0, columnNumber) = columnHeader
End With
End Sub
                

