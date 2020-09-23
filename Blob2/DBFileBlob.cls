VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBFileBlob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38AABC080324"
Option Explicit
Public Conn As Connection
Public DDataBase As String
Sub Connect()
Set Conn = New Connection
If DDataBase <> "" Then
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DDataBase & ";" & _
"Mode=ReadWrite|Share Deny None;Persist Security Info=False;Jet OLEDB"
Conn.Open DDataBase
End If
err:
' Debug.Print "Class_Initialize: " & err.Number & " " & err.Description: Resume Next
End Sub

Private Sub Class_Terminate()
On Error GoTo err
    Set Conn = Nothing
err:
' Debug.Print "Class_Terminate: " & err.Number & " " & err.Description: Resume Next
End Sub


