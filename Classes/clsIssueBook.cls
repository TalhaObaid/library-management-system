VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIssueBook"
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

Public Sub AddNew(BookId As Variant, _
                   ReaderId As Variant, _
                   FromDate As String, _
                   ToDate As String, _
                   Returned As Integer)
On Error GoTo AddErr

BookId = getBookId(BookId)
ReaderId = getReaderId(ReaderId)

Dim query As String
query = "Insert into tblIssueBook(BookId,ReaderId," & _
                   "FromDate,ToDate,Returned) values(" & _
                                BookId & "," & _
                                ReaderId & ",'" & _
                                FromDate & "','" & _
                                ToDate & "','" & _
                                Returned & "')"

db.Execute query
Exit Sub

AddErr:
  MsgBox Err.Description
  
End Sub

Public Sub Delete(strBookId As String)
On Error GoTo DelErr
Dim query As String
query = "Delete from tblIssueBook where BookId = " & strBookId
                                

db.Execute query
 Exit Sub
DelErr:
  MsgBox Err.Description
End Sub


Public Sub Update(BookId As Variant, _
                   ReaderId As Variant, _
                   FromDate As String, _
                   ToDate As String, _
                   Returned As Integer)
On Error GoTo UpdErr
BookId = getBookId(BookId)
ReaderId = getReaderId(ReaderId)

Dim query As String
query = "Update tblIssueBook set BookId=" & BookId & "," & _
                                "ReaderId=" & ReaderId & "," & _
                                "FromDate = '" & FromDate & "'," & _
                                "ToDate = '" & ToDate & "'," & _
                                "Returned = '" & Returned & "'" & _
                                " where BookId = " & BookId & "'" & _
                                " and ReaderId = " & ReaderId

db.Execute query
 Exit Sub
UpdErr:
  MsgBox Err.Description
End Sub

Public Sub UpdateCombo(ByRef cmbBook As ComboBox, _
                        ByRef cmbReader As ComboBox)

If rs.State = adStateOpen Then rs.Close

rs.Open "Select Name from tblBook order by Name", db

While rs.EOF = False
    cmbBook.AddItem rs(0)
    rs.MoveNext
    
Wend

If rs.State = adStateOpen Then rs.Close

rs.Open "Select Name from tblReader order by Name", db

While rs.EOF = False
    cmbReader.AddItem rs(0)
    rs.MoveNext
    
Wend

End Sub

Private Function getBookId(Name As Variant) As Integer

If rs.State = adStateOpen Then rs.Close

rs.Open "Select BookId from tblBook where Name = " & "'" & Name & "'", db

If rs.EOF = False Then getBookId = rs(0)
    

End Function

Private Function getReaderId(Name As Variant) As Integer

If rs.State = adStateOpen Then rs.Close

rs.Open "Select ReaderId from tblReader where Name = " & "'" & Name & "'", db

If rs.EOF = False Then getReaderId = rs(0)
    

End Function

Private Sub Class_Terminate()
    Set db = Nothing
    Set rs = Nothing
    
End Sub
