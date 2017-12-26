VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCategory"
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

Public Sub AddNew(CategoryId As String, _
                   Name As String, _
                   ParentCategory As String)
On Error GoTo AddErr
ParentCategory = getCategoryID(ParentCategory)
Dim query As String
query = "Insert into tblCategory(CategoryId,Name,ParentCategoryID ) values(" & _
                                CategoryId & ",'" & _
                                Name & "'," & _
                                ParentCategory & ")"
                                

db.Execute query
Exit Sub

AddErr:
  MsgBox Err.Description
    
End Sub

Public Sub Delete(strCategoryId As String)
On Error GoTo DelErr

Dim query As String
query = "Delete from tblCategory where CategoryId = " & strCategoryId
                                

db.Execute query
 Exit Sub
DelErr:
  MsgBox Err.Description
  
End Sub


Public Sub Update(CategoryId As String, _
                   Name As String, _
                   ParentCategory As String)
On Error GoTo UpdErr
ParentCategory = getCategoryID(ParentCategory)
Dim query As String
query = "Update tblCategory set Name = '" & Name & "'," & _
                                "ParentCategoryID = " & ParentCategory & "" & _
                                " where CategoryId = " & CategoryId


db.Execute query
 Exit Sub
UpdErr:
  MsgBox Err.Description
  
End Sub


Private Sub Class_Terminate()
    Set db = Nothing
    Set rs = Nothing
    
End Sub
Public Sub UpdateCombo(ByRef cmbParentCategory As ComboBox)

If rs.State = adStateOpen Then rs.Close

rs.Open "Select Name from tblCategory order by Name", db

While rs.EOF = False
    cmbParentCategory.AddItem rs(0)
    rs.MoveNext
    
Wend

End Sub

Private Function getCategoryID(Name As Variant) As Integer

If rs.State = adStateOpen Then rs.Close

rs.Open "Select CategoryID from tblCategory where Name = " & "'" & Name & "'", db

If rs.EOF = False Then getCategoryID = rs(0)
    

End Function

