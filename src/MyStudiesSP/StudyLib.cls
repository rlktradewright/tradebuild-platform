VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
' A Study Service Provider which serves up IStudy objects that are integrated
' into TradeBuild's study handling and charting mechanisms.

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.ICommonServiceProvider
Implements TradeBuildSP.IStudyServiceProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mLogLevel As TradeBuildSP.LogLevels

Private mStudyLongNames() As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
'
' TODO: set the name of your service provider to something unique
'
mName = "Richard's Sample Service Provider"

ReDim mStudyLongNames(0) As String

'
' TODO: add a similar line for each of your studies
'
addStudyName RsiName
'
'----------------------------------------------------

End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
With details
    .Comments = App.Comments
    .EXEName = App.EXEName
    .FileDescription = App.FileDescription
    .LegalCopyright = App.LegalCopyright
    .LegalTrademarks = App.LegalTrademarks
    .Path = App.Path
    .ProductName = App.ProductName
    .Vendor = App.CompanyName
    .VersionMajor = App.Major
    .VersionMinor = App.Minor
    .VersionRevision = App.Revision
End With
ICommonServiceProvider_Details = details
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer

'
' TODO: add a similar line for each of your studies
'
GRsi.commonServiceConsumer = commonServiceConsumer
'
'----------------------------------------------------

mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.Studies

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IStudyServiceProvider Interface Members
'================================================================================

Private Function IStudyServiceProvider_createStudy( _
                ByVal name As String) As TradeBuildSP.IStudy
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_createStudy = New rsi
End Select
'
'----------------------------------------------------

End Function

Private Function IStudyServiceProvider_getStudyDefaultParameters( _
                ByVal name As String) As TradeBuildSP.IParameters
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_getStudyDefaultParameters = GRsi.defaultParameters
End Select
'
'----------------------------------------------------

End Function

Private Property Get IStudyServiceProvider_implementedStudyNames() As String()
IStudyServiceProvider_implementedStudyNames = mStudyLongNames
End Property

Private Function IStudyServiceProvider_setStudyDefaultParameters( _
                ByVal name As String, _
                ByVal parameters As TradeBuildSP.IParameters) As Boolean

IStudyServiceProvider_setStudyDefaultParameters = True

'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    GRsi.defaultParameters = parameters
Case Else
    IStudyServiceProvider_setStudyDefaultParameters = False
End Select

End Function

Private Function IStudyServiceProvider_studyDefinition( _
                ByVal name As String) As TradeBuildSP.IStudyDefinition
'
' TODO: add a similar Case for each of your studies
'
Select Case UCase$(name)
Case UCase$(RsiName), UCase$(RsiShortName)
    Set IStudyServiceProvider_studyDefinition = GRsi.studyDefinition
End Select

End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub addStudyName(ByVal name As String)
If mStudyLongNames(0) <> "" Then
    ReDim Preserve mStudyLongNames(UBound(mStudyLongNames) + 1) As String
End If
mStudyLongNames(UBound(mStudyLongNames)) = name
End Sub



