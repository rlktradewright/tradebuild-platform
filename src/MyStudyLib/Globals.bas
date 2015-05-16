Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                    As String = "MyStudyLib26"

Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023

Public Const DummyHigh As Double = MinDouble
Public Const DummyLow As Double = MaxDouble

Public Const DefaultStudyValueName As String = "$default"

' -------------------------------------------------------------------------
' study name constants
'
'   TODO: add constants defining the short and long names for your studies
'
Public Const RsiName As String = "Relative Strength Index"
Public Const RsiShortName As String = "RSI"
'
' -------------------------------------------------------------------------


' -------------------------------------------------------------------------
' generic study parameter names - these are parameter names that are common
' to many studies
Public Const ParamMovingAverageType As String = "Mov avg type"
Public Const ParamPeriods As String = "Periods"
'
' -------------------------------------------------------------------------

' -------------------------------------------------------------------------
' names of the standard moving average types
Public Const MATypeSimpleMovingAverage As String = "SMA"
Public Const MATypeExponentialMovingAverage As String = "EMA"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

Public gLibraryManager As StudyLibraryManager

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

'================================================================================
' Methods
'================================================================================

' get TradeBuild to create a moving average study object of the
' required type, to save us having to implement our own
Public Function gCreateMA( _
                ByVal maType As String, _
                ByVal periods As Long, _
                ByVal numberOfValuesToCache As Long) As study
Dim lparams As Parameters
Dim lStudy As study
Dim valueNames(0) As String

valueNames(0) = "in"

Select Case UCase$(maType)
Case UCase$(MATypeExponentialMovingAverage), UCase$(MATypeSimpleMovingAverage)
Case Else
    ' if an invalid type is supplied use an SMA
    maType = MATypeSimpleMovingAverage
End Select

Set lStudy = gLibraryManager.createStudy(maType, "")
If lStudy Is Nothing Then Exit Function

Set lparams = New Parameters
lparams.setParameterValue ParamPeriods, periods
lStudy.initialise GenerateGUIDString, _
                lparams, _
                numberOfValuesToCache, _
                valueNames, _
                Nothing, _
                Nothing
Set gCreateMA = lStudy

End Function

'================================================================================
' Helper Function
'================================================================================







