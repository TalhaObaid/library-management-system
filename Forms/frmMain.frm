VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Library System"
   ClientHeight    =   3192
   ClientLeft      =   4068
   ClientTop       =   3096
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar mainToolbar 
      Align           =   1  'Align Top
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   550
      ButtonWidth     =   487
      ButtonHeight    =   466
      ImageList       =   "imlToolbarIcons"
      DisabledImageList=   "imlToolbarIcons"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Book"
            Object.ToolTipText     =   "Book (Ctrl + B)"
            Object.Tag             =   "Book"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Author"
            Object.ToolTipText     =   "Author (Ctrl + A)"
            Object.Tag             =   "Author"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Category"
            Object.ToolTipText     =   "Category (Ctrl + C)"
            Object.Tag             =   "Category"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reader"
            Object.ToolTipText     =   "Reader (Ctrl + R)"
            Object.Tag             =   "Reader"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IssueBook"
            Object.ToolTipText     =   "Issue Book (Ctrl + I)"
            Object.Tag             =   "IssueBook"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search (Ctrl + S)"
            Object.Tag             =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar mainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   2916
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   466
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2625
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "9/2/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "6:31 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Book"
            Object.Tag             =   "Book"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0896
            Key             =   "Author"
            Object.Tag             =   "Author"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CEA
            Key             =   "Category"
            Object.Tag             =   "Category"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113E
            Key             =   "Reader"
            Object.Tag             =   "Reader"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1592
            Key             =   "IssueBook"
            Object.Tag             =   "IssueBook"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E6
            Key             =   "Search"
            Object.Tag             =   "Search"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Begin VB.Menu mnuSystemBook 
         Caption         =   "&Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSystemAuthor 
         Caption         =   "&Author"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSystemReader 
         Caption         =   "&Reader"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSystemCategory 
         Caption         =   "&Category"
         Shortcut        =   ^C
      End
      Begin VB.Menu SeparatorSystem1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu SeparatorView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      WindowList      =   -1  'True
      Begin VB.Menu mnuToolsIssueBook 
         Caption         =   "&Issue Book"
         Shortcut        =   ^I
      End
      Begin VB.Menu SeparatorTools1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSearch 
         Caption         =   "&Search"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub mainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

 On Error Resume Next
    Select Case Button.key
        Case "Book"
           Call mnuSystemBook_Click
        Case "Author"
            Call mnuSystemAuthor_Click
        Case "Reader"
            Call mnuSystemReader_Click
        Case "Category"
            Call mnuSystemCategory_Click
        Case "IssueBook"
            Call mnuToolsIssueBook_Click
        Case "Search"
            Call mnuToolsSearch_Click
            
    End Select
End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub



Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuSystemCategory_Click()
    Dim formCategory As frmCategory
    Set formCategory = New frmCategory
    Load formCategory
    formCategory.Show vbModal
End Sub

Private Sub mnuToolsIssueBook_Click()
    Dim formIssueBook As frmIssueBook
    Set formIssueBook = New frmIssueBook
    Load formIssueBook
    formIssueBook.Show vbModal
End Sub



Private Sub mnuToolsSearch_Click()
    Dim formSearch As frmSearch
    Set formSearch = New frmSearch
    Load formSearch
    formSearch.Show vbModal

End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    mainStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    mainToolbar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuSystemAuthor_Click()
    Dim formAuthor As frmAuthor
    Set formAuthor = New frmAuthor
    Load formAuthor
    formAuthor.Show vbModal
    
End Sub

Private Sub mnuSystemBook_Click()
    
    Dim formBook As frmBook
    Set formBook = New frmBook
    Load formBook
    formBook.Show vbModal
    
End Sub

Private Sub mnuSystemReader_Click()
    
    Dim formReader As frmReader
    Set formReader = New frmReader
    Load formReader
    formReader.Show vbModal
End Sub
Private Sub mnuSystemExit_Click()
    fMainForm.Visible = False
    Set fMainForm = Nothing
    
    End
    
End Sub

