VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category"
   ClientHeight    =   3180
   ClientLeft      =   4872
   ClientTop       =   4128
   ClientWidth     =   5724
   Icon            =   "frmCategory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5724
   Begin VB.ComboBox cmbParentCategory 
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1095
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtCategory 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2310
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   3465
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4620
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   307
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc AdodcCategory 
      Height          =   330
      Left            =   4440
      Top             =   2040
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
      RecordSource    =   $"frmCategory.frx":0442
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
   Begin MSDataGridLib.DataGrid grdCategory 
      Bindings        =   "frmCategory.frx":0518
      Height          =   1455
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8700
      _ExtentY        =   2561
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
   Begin VB.Label lblLabels 
      Caption         =   "Parent Category"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   12
      Top             =   615
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   375
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category ID:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CategoryCLS As ClsCategory
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean

Private Sub cmdRefresh_Click()
  
AdodcCategory.Refresh

AdodcCategory.Recordset.Requery
grdCategory.ReBind
grdCategory.Refresh


Set grdCategory.DataSource = AdodcCategory

End Sub

Private Sub Form_Load()
  Set CategoryCLS = New ClsCategory
  Call CategoryCLS.UpdateCombo(Me.cmbParentCategory)
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
  CategoryCLS.Delete (txtCategory.Text)
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

  Call CategoryCLS.AddNew(Me.txtCategory, Me.txtName, _
                         Me.cmbParentCategory)
  
  
ElseIf mbAddNewFlag = False And mbEditFlag = True Then
    Call CategoryCLS.Update(Me.txtCategory, Me.txtName, _
                          Me.cmbParentCategory)
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
    
    txtCategory.Enabled = Not bVal And Not mbEditFlag
    txtName.Enabled = Not bVal
    
    cmbParentCategory.Enabled = Not bVal
  
End Sub


Private Sub grdCategory_Click()
    Me.txtCategory = grdCategory.Columns(0).Text
    Me.txtName = grdCategory.Columns(1).Text
    Me.cmbParentCategory = grdCategory.Columns(2).Text
End Sub


Private Sub updateCombos()
    Call CategoryCLS.UpdateCombo(Me.cmbParentCategory)

End Sub


Private Sub Form_Terminate()
    Set CategoryCLS = Nothing
    Set formCategory = Nothing
    
End Sub

