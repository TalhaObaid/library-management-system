VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book"
   ClientHeight    =   5736
   ClientLeft      =   3660
   ClientTop       =   2016
   ClientWidth     =   5880
   Icon            =   "frmBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   5880
   Begin MSAdodcLib.Adodc AdodcBook 
      Height          =   330
      Left            =   4320
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\Visual Basic\Library\Database\Library.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\Visual Basic\Library\Database\Library.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmBook.frx":0442
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grdBook 
      Bindings        =   "frmBook.frx":058C
      Height          =   2055
      Left            =   240
      TabIndex        =   23
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9335
      _ExtentY        =   3620
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1215
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtBook 
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2430
      TabIndex        =   19
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   3585
      TabIndex        =   18
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4740
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1200
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbEdition 
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   2070
      Width           =   2055
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   1095
      Width           =   2055
   End
   Begin VB.ComboBox cmbAuthor 
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   750
      Width           =   2055
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   2415
      Width           =   2775
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   1755
      Width           =   2415
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   435
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Description:"
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   10
      Top             =   2415
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Edition:"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ISBN:"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Title:"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category ID:"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Author ID:"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Book ID:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private BookCLS As ClsBook
Attribute BookCLS.VB_VarHelpID = -1
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean

Private Sub cmdRefresh_Click()
  
AdodcBook.Refresh

AdodcBook.Recordset.Requery
grdBook.ReBind
grdBook.Refresh

Set grdBook.DataSource = AdodcBook

End Sub

Private Sub Form_Load()
  Set BookCLS = New ClsBook
  Call BookCLS.UpdateCombo(Me.cmbAuthor, Me.cmbCategory)
  SetButtons (True)
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
 
  mbAddNewFlag = True
  SetButtons False
  mbEditFlag = False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  BookCLS.Delete (txtBook.Text)
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub


Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  mbEditFlag = True
  mbAddNewFlag = False
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  
  mbEditFlag = False
  mbAddNewFlag = False
    
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
    
If mbAddNewFlag = True And mbEditFlag = False Then

  Call BookCLS.AddNew(Me.txtBook, Me.txtName, _
                        Me.cmbAuthor, Me.cmbCategory, _
                        Me.txtTitle, Me.txtISBN, _
                        Me.cmbEdition, Me.txtDesc)
  
  
ElseIf mbAddNewFlag = False And mbEditFlag = True Then
    Call BookCLS.Update(Me.txtBook, Me.txtName, _
                        Me.cmbAuthor, Me.cmbCategory, _
                        Me.txtTitle, Me.txtISBN, _
                        Me.cmbEdition, Me.txtDesc)
End If
  SetButtons True
  
  mbAddNewFlag = False
  mbEditFlag = False


  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  
  Unload Me
End Sub

Private Sub SetButtons(bVal As Boolean)
    cmdAdd.Visible = bVal
    cmdUpdate.Visible = Not bVal
    cmdCancel.Visible = Not bVal
    cmdEdit.Visible = bVal
    cmdDelete.Visible = bVal
    cmdClose.Visible = bVal
    cmdRefresh.Visible = bVal
    
    txtBook.Enabled = Not bVal And Not mbEditFlag
    txtName.Enabled = Not bVal
    cmbAuthor.Enabled = Not bVal
    cmbCategory.Enabled = Not bVal
    txtTitle.Enabled = Not bVal
    txtISBN.Enabled = Not bVal
    cmbEdition.Enabled = Not bVal
    txtDesc.Enabled = Not bVal
  
End Sub
Private Sub updateCombos()
    Call BookCLS.UpdateCombo(Me.cmbAuthor, _
                        Me.cmbCategory)

End Sub


Private Sub grdBook_Click()
    Me.txtBook = grdBook.Columns(0).Text
    Me.txtName = grdBook.Columns(1).Text
    Me.cmbAuthor = grdBook.Columns(2).Text
    Me.cmbCategory = grdBook.Columns(3).Text
    Me.txtTitle = grdBook.Columns(4).Text
    Me.txtISBN = grdBook.Columns(5).Text
    Me.cmbEdition = grdBook.Columns(6).Text
    Me.txtDesc = grdBook.Columns(7).Text
    
    
End Sub


Private Sub Form_Terminate()
    Set BookCLS = Nothing
    Set formBook = Nothing
    
End Sub

