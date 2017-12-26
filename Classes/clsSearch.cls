VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSearch"
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
Private Sub Class_Terminate()
    Set db = Nothing
    Set rs = Nothing
    
End Sub

Public Sub UpdateCombo(value As String, ByRef cmb As ComboBox)
Dim query As String

If rs.State = adStateOpen Then rs.Close

Select Case LCase(Trim(value))

    Case "bookid"
        query = "Select bookid from tblBook order by bookid"

    Case "name"
        query = "Select Name from tblBook order by Name"

    Case "author"
        query = "Select Name from tblAuthor order by Name"

    Case "category"
        query = "Select Name from tblcategory order by Name"

    Case "title"
        query = "Select title from tblBook order by title"

    Case "description"
        query = "Select description from tblBook order by description"

End Select

rs.Open query, db

While rs.EOF = False
    cmb.AddItem Trim(rs(0))
    rs.MoveNext

Wend

End Sub


Public Sub search(key As String, value As String, ByRef grd As MSFlexGrid)
Dim query As String

If rs.State = adStateOpen Then rs.Close

Select Case LCase(Trim(key))
    
    Case "bookid"
        query = "Select * from tblBook where " & key & " = " & value & " order by name"
    
    Case "name"
        query = "Select * from tblBook where " & key & " = '" & value & "' order by name"
    
    Case "author"
        query = "Select distinct book.* from tblBook book,tblAuthor author where author.authorId = book.AuthorID and Author.name = '" & value & "' order by book.Name"
        
    Case "category"
        query = "Select distinct book.* from tblbook book,tblcategory category where  category.categoryId = book.categoryID and category.name = '" & value & "' order by book.Name"
        
    Case "title"
        query = "Select * from tblBook where " & key & " = '" & value & "' order by name"
        
    Case "description"
        query = "Select * from tblBook where " & key & " = '" & value & "' order by name"

End Select

rs.Open query, db


grd.Cols = 9

Dim row As Integer
row = 0
grd.row = row: grd.Col = 1: grd.Text = "Book ID"

grd.row = row: grd.Col = 2: grd.Text = "Name"

grd.row = row: grd.Col = 3: grd.Text = "Author ID"

grd.row = row: grd.Col = 4: grd.Text = "Category ID"

grd.row = row: grd.Col = 5: grd.Text = "Title"

grd.row = row: grd.Col = 6: grd.Text = "ISBN"

grd.row = row: grd.Col = 7: grd.Text = "Edition"

grd.row = row: grd.Col = 8: grd.Text = "Description"

grd.Rows = rs.RecordCount + 1

While rs.EOF = False
    row = row + 1
    
    grd.row = row:    grd.Col = 1:    grd.Text = Trim(rs(0))
    
    grd.row = row:    grd.Col = 2:    grd.Text = Trim(rs(1))
    
    grd.row = row:    grd.Col = 3:    grd.Text = Trim(rs(2))
    
    grd.row = row:    grd.Col = 4:    grd.Text = Trim(rs(3))
    
    grd.row = row:    grd.Col = 5:    grd.Text = Trim(rs(4))
    
    grd.row = row:    grd.Col = 6:    grd.Text = Trim(rs(5))
    
    grd.row = row:    grd.Col = 7:    grd.Text = Trim(rs(6))
    
    grd.row = row:    grd.Col = 8:    grd.Text = Trim(rs(7))
    
    rs.MoveNext

Wend


End Sub

