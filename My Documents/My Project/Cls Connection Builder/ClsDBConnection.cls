VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mVarActiveConnection As ADODB.Connection
Private mVarConnectMode As Long

Public ConnectionString As String

Property Get ConnectMode() As Long
    ConnectMode = mVarConnectMode
End Property

Property Get ActiveConnection() As ADODB.Connection
    Set ActiveConnection = mVarActiveConnection
End Property

Function CreateConnection() As Long
    On Error Resume Next
    Set mVarActiveConnection = New ADODB.Connection
    mVarActiveConnection.Open Me.ConnectionString
    CreateConnection = Err.Number
    mVarConnectMode = Err.Number
    On Error GoTo 0
End Function

Sub DestroyConnection()
    mVarConnectMode = -1
    mVarActiveConnection.Close
    Set mVarActiveConnection = Nothing
End Sub

Function ConnectionWizard() As String
    Dim connObj As MSDASC.DataLinks
    Dim StrConnection As String
    
    Err.Clear
    On Error Resume Next
    Set connObj = New MSDASC.DataLinks
    StrConnection = connObj.PromptNew
    If Err.Number = 0 Then
        ConnectionWizard = StrConnection & ";"
    Else
        ConnectionWizard = ""
    End If
    Set connObj = Nothing
    On Error GoTo 0
End Function

