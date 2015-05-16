VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionParams"
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

Implements Stringable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ConnectionParams"

'@================================================================================
' Member variables
'@================================================================================

Private mDbType                                     As DatabaseTypes
Private mServer                                     As String
Private mDatabaseName                               As String
Private mUsername                                   As String
Private mPassword                                   As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ConnectionString() As String
ConnectionString = generateConnectionString(False)
End Property

Friend Property Get DatabaseType() As DatabaseTypes
DatabaseType = mDbType
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pDbType As DatabaseTypes, _
                ByVal pServer As String, _
                ByVal pDatabaseName As String, _
                ByVal pUsername As String, _
                ByVal pPassword As String)
mDbType = pDbType
mServer = pServer
mDatabaseName = pDatabaseName
mUsername = pUsername
mPassword = pPassword
End Sub

Public Function ToString() As String
ToString = generateConnectionString(True)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


Private Function generateConnectionString( _
                ByVal pHidePassword As Boolean) As String
Const ProcName As String = "generateConnectionString"
On Error GoTo Err

Dim lPassword As String

If pHidePassword Then
    lPassword = String(Len(mPassword), "*")
Else
    lPassword = mPassword
End If

Select Case mDbType
Case DbSQLServer7
    generateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & mServer & _
                        ";Initial Catalog=" & mDatabaseName & _
                        IIf(mUsername = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & mUsername & ";Password=" & lPassword)
Case DbSQLServer2000
    generateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & mServer & _
                        ";Initial Catalog=" & mDatabaseName & _
                        IIf(mUsername = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & mUsername & ";Password=" & lPassword)
Case DbSQLServer2005, DbSQLServer, DbSQLServer7, DbSQLServer2000
'   NB: we don't use the SQL native Client (SQLNCLI) because it
'   seems to cause problems on Windows 2000.
'    GenerateConnectionString = "Provider=SQLNCLI" & _
'                        ";DataTypeCompatibility=80" & _
'                        ";Data Source=" & server & _
'                        ";Initial Catalog=" & databaseName & _
'                        ";Network Library=DBMSSOCN" & _
'                        IIf(username = "", _
'                            ";Integrated Security=SSPI", _
'                            ";Uid=" & username & ";Pwd=" & lpassword)
    generateConnectionString = "Provider=SQLOLEDB" & _
                        ";Data Source=" & mServer & _
                        ";Initial Catalog=" & mDatabaseName & _
                        IIf(mUsername = "", _
                            ";Integrated Security=SSPI", _
                            ";User Id=" & mUsername & ";Password=" & lPassword)
Case DbMySQL5
    generateConnectionString = "Driver={MySQL ODBC 5.1 Driver}" & _
                        ";Server=" & mServer & _
                        ";Port=3306" & _
                        ";Database=" & mDatabaseName & _
                        ";User=" & mUsername & _
                        ";Password=" & lPassword & _
                        ";Option=3"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



