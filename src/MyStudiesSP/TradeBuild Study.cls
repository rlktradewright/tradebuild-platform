VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'
' TODO: you must do the following:
'
'  1.   Change the name of this class to your study's name.
'       
'  2.   Define constants for your study's long and short names in the
'       Globals module
'
'  3.   Create a module whose name is the name of this study class preceded 
'       by a G. In this module, define constants for the input values, output 
'       values and parameters for your study.
'       
'       This module must also contain properties to get and set the default 
'       parameters, and to get the study definition.
'       
'       The easiest way to create this module is to copy from an existing one
'       and amend as necessary.
'       
'  4.   Make additions to the StudyServiceProvider class, as indicated by the
'       TODO comments therein.
'	
'  5.   Check out the other TODO comments in this class for work you must do 
'       here.




'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IStudy

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

'-------this variable is required in all studies---------
' It provides a common implementation of the mechanisms for integrating the
' study with charting
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------

'-------these variables may be required in some studies--
Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer
'--------------------------------------------------------

'
' TODO: declare your variables here
'

' The Study Foundation object automatically creates one study value cache object 
' for each output value defined in the study definition. These cache objects store
' the study values for each bar, and are used by the charting infrastructure and
' other components.
' 
' You must store each value in the appropriate cache using its add or update method.
' To access them, you need to set a reference to each one.
'
' For example, if your study has a value whose name is Foo, then declare the 
' following variable - you will then initialise it in the customInitialise
' procedure, which is called after the caches have been created:
'
Private mFooCache As IStudyValueCache


'================================================================================
' Class Event Handlers
'================================================================================

' The following may be useful in some circumstances while debugging

' TODO: change MyStudy to the name of your study
Private Sub Class_Initialize()
Debug.Print "MyStudy study created"
End Sub

' TODO: change MyStudy to the name of your study
Private Sub Class_Terminate()
Debug.Print "MyStudy study terminated"
End Sub

'================================================================================
' Custom procedures that must be implemented
'================================================================================

' Called when the class has been initialised by the infrastructure
Private Sub customInitialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
'
' TODO: remove these lines if you don't need these variables
'
Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
'--------------------------------------------------------

customProcessParameters parameters

'
' TODO: set references to your study value caches here
'
' For example, if your study value name is defined in a constant 
' called MyStudyValueFoo:
'
Set mFooCache = mStudyFoundation.FindValueCache(MyStudyValueFoo)


End Sub

' Called for each input value that needs to be notified to the study. Any
' historic values are notified first (from oldest to most recent), followed by 
' realtime values as they occur.
Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
'
' TODO: write code to handle each input value notified to the study.
'
'
' Notes: 
'
'   ev.ValueName contains the input value name as specifed in the study
'   definition (studies that have multiple inputs can use this to determine
'   which input is being notified
'
'   ev.Value contains the value itself
'
'
' The following methods of the StudyFoundation object are useful:
'
'   mStudyFoundation.barNumber  the number of the current bar within the
'                               timeframe
'
'   mStudyFoundation.isNewBar   indicates whether this input value is the
'                               first in a new bar
'
'   mStudyFoundation.localBarNumber     the first bar processed by this study
'                               has a localNumber of 1. For each subsequent
'                               bar, this is incremented by 1. NB: this is
'                               not necessarily the same as BarNumber
'
'   mStudyFoundation.NotifyValue    use this to notify any interested parties
'                               (eg chart or studies using this study)of the
'                               current output values


' The following code snippets may be of interest:

Static sPrevVal     As Variant  ' the previously notified value in this bar
Dim foo             As Variant  ' the study value to be calculated

If Not mStudyFoundation.isNewBar Then
    ' if this input value is the same as the previous one in this bar, no
    ' need to do anything as it won't change the study value
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

'
' your code to calculate the current study values goes here
'


' After calculating the new value(s), either add them to the relevant cache
' (if it's a new bar) or update the cache
If mStudyFoundation.isNewBar Then
    mFooCache.addValue foo, ev.barnumber, ev.timestamp
Else
    mFooCache.updateValue foo, ev.barnumber, ev.timestamp
End If

' finally notify all interested parties of each new study value
notifyValue foo, MyStudyValueFoo, ev.barnumber, ev.timestamp

End Sub

' Called by the infrastructure so that it can determine how many historical
' input values are need by this study.
Private Function customNumberOfBarsRequired() As Long
'
' TODO: write code to return, as the value of this function, the
' number of bars of historical data needed to ensure this study is
' ready for use.
'
' For example, a moving average of n periods cannot return a meaningful
' value until n bars have been processed.

End Function

' Called from the customInitialise method (this is really a helper method
' but is kept here to emphasise its commonality between studies)
Private Sub customProcessParameters(ByVal params As IParameters)
'
' TODO: remove the code from this method if your study has no parameters
'
Dim param As TradeBuildSP.IParameter

' NB: the For Each construct cannot be used with objects of the
' IParameters class

Set param = params.getFirstParameter

Do While Not param Is Nothing
    '
    ' TODO: add a case for each parameter defined for your study:
    ' replace ParamNameConst by your parameter name constant
    Select Case UCase$(param.name)
    Case UCase$(ParamNameConst)
        ' process the parameter
    End Select
    Set param = params.getNextParameter
Loop

End Sub

' called by the infrastructure to obtain the study definition
Private Function customStudyDefinition() As IStudyDefinition
'
' TODO: add code to return your study's StudyDefinition object. Do
' this by getting the studyDefinition property from your study's custom module,
' for example:
' Set customStudyDefinition = GMyStudy.studyDefinition
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'
' TODO: add any properties you need. Note that such properties are only accessible
' when your study is accessed via its default interface, not via the IStudy
' interface. Only do this if you really know what you're doing!
'

'================================================================================
' Methods
'================================================================================

'
' TODO: add any methods you need. Note that such methods are only accessible
' when your study is accessed via its default interface, not via the IStudy
' interface. Only do this if you really know what you're doing!
'

'================================================================================
' Helper Functions
'================================================================================

'
' TODO: add any helper functions needed by other code you've written
'

' This helper function is useful for notifying output values to
' higher-level studies and listeners.
Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date)
Dim ev As StudyValueEvent
If Not IsEmpty(value) Then
    ev.value = value
    ev.valueName = valueName
    Set ev.Source = Me
    ev.barnumber = barnumber
    ev.timestamp = timestamp
    mStudyFoundation.notifyValue ev
End If
End Sub


'================================================================================
' From this point on the code is identical for all studies - do not change unless
' you are sure you know what you're doing!!!!
'================================================================================


'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Set IStudy_addStudy = mStudyFoundation.addStudy( _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudy_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
IStudy_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Property Get IStudy_id() As String
IStudy_id = mStudyFoundation.id
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy, _
                            commonServiceConsumer, _
                            studyServiceConsumer
                            

customInitialise commonServiceConsumer, _
                studyServiceConsumer, _
                id, _
                parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mStudyFoundation.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mStudyFoundation.parameters
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = customStudyDefinition
End Property










