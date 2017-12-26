VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Book"
   ClientHeight    =   3045
   ClientLeft      =   4170
   ClientTop       =   3300
   ClientWidth     =   4575
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdSearch 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   0
      FormatString    =   ""
   End
   Begin VB.ComboBox cmbValue 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   735
      Left            =   3720
      Picture         =   "frmSearch.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbField 
      Height          =   315
      ItemData        =   "frmSearch.frx":0884
      Left            =   1560
      List            =   "frmSearch.frx":089A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Containig Word :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Search a Book By:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SearchCls As ClsSearch

Private Sub cmbField_Click()
    cmbValue.Clear
    
    Call SearchCls.UpdateCombo(cmbField.Text, cmbValue)
End Sub

Private Sub cmdSearch_Click()
    grdSearch.Clear
    Call SearchCls.search(cmbField.Text, cmbValue.Text, grdSearch)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      Unload Me
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Terminate()
    Set SearchCls = Nothing
    Set formSearch = Nothing
    
End Sub

Private Sub Form_Load()
    Set SearchCls = New ClsSearch
    
End Sub

