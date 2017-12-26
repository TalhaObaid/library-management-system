VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author"
   ClientHeight    =   3120
   ClientLeft      =   3792
   ClientTop       =   3300
   ClientWidth     =   5748
   Icon            =   "frmAuthor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5748
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   427
      Width           =   1815
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   735
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4620
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   3465
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2310
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1095
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdodcAuthor 
      Height          =   330
      Left            =   4440
      Top             =   2160
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
      RecordSource    =   "Select * from tblAuthor order by AuthorID"
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
   Begin MSDataGridLib.DataGrid grdAuthor 
      Bindings        =   "frmAuthor.frx":0442
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   1200
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
      Caption         =   "Author ID:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Description:"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   11
      Top             =   735
      Width           =   1815
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private AuthorCLS As ClsAuthor
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean

Private Sub cmdRefresh_Click()
  
AdodcAuthor.Refresh

AdodcAuthor.Recordset.Requery
grdAuthor.ReBind
grdAuthor.Refresh


Set grdAuthor.DataSource = AdodcAuthor

End Sub

Private Sub Form_Load()
  Set AuthorCLS = New ClsAuthor
  SetButtons (True)
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    
      
  End Select
End Sub

Private Sub Form_Terminate()
    Set AuthorCLS = Nothing
    Set formAuthor = Nothing
    
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
  AuthorCLS.Delete (txtAuthor.Text)
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

  Call AuthorCLS.AddNew(Me.txtAuthor, Me.txtName, _
                         Me.txtDesc)
  
  
ElseIf mbAddNewFlag = False And mbEditFlag = True Then
    Call AuthorCLS.Update(Me.txtAuthor, Me.txtName, _
                          Me.txtDesc)
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
    
    txtAuthor.Enabled = Not bVal And Not mbEditFlag
    txtName.Enabled = Not bVal
    
    txtDesc.Enabled = Not bVal
  
End Sub


Private Sub grdAuthor_Click()
    Me.txtAuthor = grdAuthor.Columns(0).Text
    Me.txtName = grdAuthor.Columns(1).Text
    Me.txtDesc = grdAuthor.Columns(2).Text
End Sub
