VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReader 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reader"
   ClientHeight    =   4164
   ClientLeft      =   3792
   ClientTop       =   3048
   ClientWidth     =   5760
   Icon            =   "frmReader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4164
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   330
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   660
      Width           =   2415
   End
   Begin VB.TextBox txtContactNo 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   1005
      Width           =   2415
   End
   Begin VB.TextBox txtInstitution 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1335
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4620
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   3465
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   2310
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtReader 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   1095
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdodcReader 
      Height          =   330
      Left            =   4440
      Top             =   3120
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
      RecordSource    =   "select * from tblReader order by ReaderID"
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
   Begin MSDataGridLib.DataGrid grdReader 
      Bindings        =   "frmReader.frx":0442
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5295
      _ExtentX        =   9335
      _ExtentY        =   2985
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
      Caption         =   "Reader ID:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   375
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   15
      Top             =   690
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Contact No:"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Institution:"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   13
      Top             =   1335
      Width           =   1815
   End
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ReaderCLS As ClsReader
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean

Private Sub cmdRefresh_Click()
  
AdodcReader.Refresh

AdodcReader.Recordset.Requery
grdReader.ReBind
grdReader.Refresh

Set grdReader.DataSource = AdodcReader

End Sub

Private Sub Form_Load()
  Set ReaderCLS = New ClsReader
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
  ReaderCLS.Delete (txtReader.Text)
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
  MsgBox Err.Institutionription
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

  Call ReaderCLS.AddNew(Me.txtReader, Me.txtName, _
                        Me.txtAddress, Me.txtContactNo, _
                        Me.txtInstitution)
  
  
ElseIf mbAddNewFlag = False And mbEditFlag = True Then
    Call ReaderCLS.Update(Me.txtReader, Me.txtName, _
                        Me.txtAddress, Me.txtContactNo, _
                        Me.txtInstitution)
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
    
    txtReader.Enabled = Not bVal And Not mbEditFlag
    txtName.Enabled = Not bVal
    txtAddress.Enabled = Not bVal
    txtContactNo.Enabled = Not bVal
    txtInstitution.Enabled = Not bVal
  
End Sub

Private Sub grdReader_Click()
    Me.txtReader = grdReader.Columns(0).Text
    Me.txtName = grdReader.Columns(1).Text
    Me.txtAddress = grdReader.Columns(2).Text
    Me.txtContactNo = grdReader.Columns(3).Text
    Me.txtInstitution = grdReader.Columns(4).Text
    
    
End Sub




Private Sub Form_Terminate()
    Set ReaderCLS = Nothing
    Set formReader = Nothing
    
End Sub

