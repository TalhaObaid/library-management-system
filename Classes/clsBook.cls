VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rs As Recordset
Attribute rs.VB_VarHelpID = -1
Dim db As Connection

Private Sub Class_Initialize()
  
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Persist Security Info=False"


  Set rs = New Recordset
  
End Sub

Public Sub AddNew(BookId As String, _
                   Name As String, _
                   AuthorId As String, _
                   CategoryId As String, _
                   Title As String, _
                   ISBN As String, _
                   Edition As String, _
                   Description As String)
On Error GoTo AddErr

AuthorId = getAuthorID(AuthorId)
CategoryId = getCategoryID(CategoryId)

Dim query As String
query = "Insert into tblBook(BookId,Name,AuthorId,CategoryId," & _
                   "Title,ISBN,Edition,Description ) values(" & _
                                BookId & ",'" & _
                                Name & "'," & _
                                AuthorId & "," & _
                                CategoryId & ",'" & _
                                Title & "','" & _
                                ISBN & "','" & _
                                Edition & "','" & _
                                Description & "')"
                                

db.Execute query
Exit Sub

AddErr:
  MsgBox Err.Description
  
End Sub

Public Sub Delete(strBookId As String)
On Error GoTo DelErr
Dim query As String
query = "Delete from tblBook where BookId = " & strBookId
                                

db.Execute query
 Exit Sub
DelErr:
  MsgBox Err.Description
End Sub


Public Sub Update(BookId As String, _
                   Name As String, _
                   AuthorId As Variant, _
                   CategoryId As Variant, _
                   Title As String, _
                   ISBN As String, _
                   Edition As Integer, _
                   Description As String)
On Error GoTo UpdErr
AuthorId = getAuthorID(AuthorId)
CategoryId = getCategoryID(CategoryId)

Dim query As String
query = "Update tblBook set Name = '" & Name & "'," & _
                               "AuthorId=" & AuthorId & "," & _
                                "CategoryId=" & CategoryId & "," & _
                                "Title = '" & Title & "'," & _
                                "ISBN = '" & ISBN & "'," & _
                                "Edition = '" & Edition & "'," & _
                                "Description = '" & Description & "'" & _
                                " where BookId = " & BookId


db.Execute query
 Exit Sub
UpdErr:
  MsgBox Err.Description
End Sub

Public Sub UpdateCombo(ByRef cmbAuthor As ComboBox, _
                        ByRef cmbCategory As ComboBox)

If rs.State = adStateOpen Then rs.Close

rs.Open "Select Name from tblAuthor order by Name", db

While rs.EOF = False
    cmbAuthor.AddItem rs(0)
    rs.MoveNext
    
Wend

If rs.State = adStateOpen Then rs.Close

rs.Open "Select Name from tblCategory order by Name", db

While rs.EOF = False
    cmbCategory.AddItem rs(0)
    rs.MoveNext
    
Wend

End Sub

Private Function getAuthorID(Name As Variant) As Integer

If rs.State = adStateOpen Then rs.Close

rs.Open "Select authorID from tblAuthor where Name = " & "'" & Name & "'", db

If rs.EOF = False Then getAuthorID = rs(0)
    

End Function

Private Function getCategoryID(Name As Variant) As Integer

If rs.State = adStateOpen Then rs.Close

rs.Open "Select CategoryID from tblCategory where Name = " & "'" & Name & "'", db

If rs.EOF = False Then getCategoryID = rs(0)
    

End Function

Private Sub Class_Terminate()
    Set db = Nothing
    Set rs = Nothing
    
End Sub
