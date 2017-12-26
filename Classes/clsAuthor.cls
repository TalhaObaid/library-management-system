VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rs As Recordset
Dim db As Connection

Private Sub Class_Initialize()
  
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Persist Security Info=False"


  Set rs = New Recordset
  
End Sub

Public Sub AddNew(AuthorId As String, _
                   Name As String, _
                   Description As String)
On Error GoTo AddErr

Dim query As String
query = "Insert into tblAuthor(AuthorId,Name,Description ) values(" & _
                                AuthorId & ",'" & _
                                Name & "','" & _
                                Description & "')"
                                

db.Execute query
Exit Sub

AddErr:
  MsgBox Err.Description
  
End Sub

Public Sub Delete(strAuthorId As String)
On Error GoTo DelErr
Dim query As String
query = "Delete from tblAuthor where AuthorId = " & strAuthorId
                                

db.Execute query
 Exit Sub
DelErr:
  MsgBox Err.Description
End Sub


Public Sub Update(AuthorId As String, _
                   Name As String, _
                   Description As String)
On Error GoTo UpdErr

Dim query As String
query = "Update tblAuthor set Name = '" & Name & "'," & _
                                "Description = '" & Description & "'" & _
                                " where AuthorId = " & AuthorId


db.Execute query
 Exit Sub
UpdErr:
  MsgBox Err.Description
End Sub


Private Sub Class_Terminate()
    Set db = Nothing
    Set rs = Nothing
    
End Sub


