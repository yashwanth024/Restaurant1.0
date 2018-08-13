VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Restaurant Management System"
   ClientHeight    =   10560
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14430
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   19680
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   19560
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   21600
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   25800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10395
      Left            =   0
      Picture         =   "mdiMain.frx":C1E92
      ScaleHeight     =   10335
      ScaleWidth      =   14370
      TabIndex        =   0
      Top             =   0
      Width           =   14430
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   0
         Picture         =   "mdiMain.frx":1867C7
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Hide SideBar"
         Top             =   120
         Width           =   615
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Caption         =   "    Today"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   0
         TabIndex        =   8
         Top             =   7440
         Width           =   3375
         Begin VB.Image Image4 
            Height          =   240
            Left            =   120
            Picture         =   "mdiMain.frx":18BFA9
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Caption         =   "   User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   3375
         Begin VB.Image Image3 
            Height          =   240
            Left            =   0
            Picture         =   "mdiMain.frx":18C333
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label18"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   480
            TabIndex        =   7
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "SHREE  KRISHANA  RESTAURANT "
         BeginProperty Font 
            Name            =   "Jokerman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete User"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   6
         Left            =   1320
         TabIndex        =   16
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   240
         Picture         =   "mdiMain.frx":18C8BD
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   3
         Left            =   1320
         TabIndex        =   15
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   240
         Picture         =   "mdiMain.frx":19209F
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   2
         Left            =   1320
         TabIndex        =   14
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   240
         Picture         =   "mdiMain.frx":197881
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New User"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   240
         Picture         =   "mdiMain.frx":19D063
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   7095
         Left            =   0
         Picture         =   "mdiMain.frx":1A2845
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   690
         Left            =   240
         Picture         =   "mdiMain.frx":1A7294
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   3360
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off / Exit"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   5
         Left            =   1320
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblUserAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "User Account Panel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   4
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Image imgUserAccount 
         Height          =   3465
         Left            =   0
         Picture         =   "mdiMain.frx":1A82A7
         Top             =   0
         Width           =   3405
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&User Panel"
      Begin VB.Menu mnuSideBar 
         Caption         =   "SideBar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuMasterCategory 
         Caption         =   "&Category"
      End
      Begin VB.Menu mnuMasterItemTable 
         Caption         =   "&ItemTable"
      End
      Begin VB.Menu mnuMasterTables 
         Caption         =   "&Tables"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuMasterOrder 
         Caption         =   "&Order"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuBill 
         Caption         =   "&Bill"
         Begin VB.Menu mnuConsumption 
            Caption         =   "&Consumption"
         End
      End
      Begin VB.Menu mnuReportMenu 
         Caption         =   "&Menu"
      End
      Begin VB.Menu mnuCollection 
         Caption         =   "&Todays Collection"
      End
      Begin VB.Menu mnuMonthlyCollection 
         Caption         =   "&Monthly Collection"
      End
      Begin VB.Menu mnuAllCustomerOrder 
         Caption         =   "All Customer Order"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuViewHelp 
         Caption         =   "View Help"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Picture1.Visible = True Then
Picture1.Visible = False
mnuSideBar.Checked = False
End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
End Sub
Private Sub imgUserAccount_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False

End Sub

Private Sub lblShortcut_Click(Index As Integer)
    Select Case (Index)
        Case 0:
        mdiMain.Hide
            frmLogin.Show
            
        Case 1:
              frmAddNewUser.Show , Me
              frmDeleteUser.Hide
              frmChangePassword.Hide
        Case 2:
               Shell "calc.exe", vbNormalFocus
               
        Case 3:
               cdlg.ShowHelp
            
        Case 4:
            frmChangePassword.Show , Me
            frmAddNewUser.Hide
            frmDeleteUser.Hide
        Case 5:
           iLogOutReply = MsgBox(UserName & ", Are You Sure You Wish To Log Out Of Your Account?", vbYesNo + vbQuestion, "Log Out?")
            If iLogOutReply = vbYes Then
            End
            End If
        Case 6:
              frmDeleteUser.Show , Me
              frmAddNewUser.Hide
              frmChangePassword.Hide
            End Select
End Sub

Private Sub lblShortcut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case (Index)
        Case 0:
            lblShortcut(0).ForeColor = &H800000
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = True
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
        Case 1:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H800000
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = True
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
        Case 2:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H800000
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = True
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
        Case 3:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H800000
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = True
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
        Case 4:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H800000
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = True
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
        Case 5:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H800000
            lblShortcut(6).ForeColor = &H80000006
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = True
            lblShortcut(6).FontUnderline = False
           
        Case 6:
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H800000
            
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = True
End Select
End Sub

Private Sub MDIForm_Load()
Label2.Caption = Format(Date, "long date")
frmLogin.Show vbModel
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim re As Variant
re = MsgBox("Do You Want Exit", vbYesNo)
If re = vbYes Then
End
Else
Cancel = 1
End If
End Sub
Private Sub mnuAdvanceBookings_Click()
    frmRBooking.Show
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show , Me
End Sub

Private Sub mnuAllCustomerOrder_Click()
rptAllCustomerOrder.Show
End Sub

Private Sub mnuCollection_Click()
    If DataEnvironment1.rscmdTCollection_Grouping.State Then
        DataEnvironment1.rscmdTCollection_Grouping.Close
    End If
    DataEnvironment1.cmdTCollection_Grouping Format(Date, "DD-MM-YY")
    DRTCollection.Show
End Sub
Private Sub mnuConsumption_Click()
    frmCBill.Show
End Sub
Private Sub mnuExit_Click()
If MsgBox("Are You Sure Exit System ?", vbYesNo + vbInformation, "Warning") = vbYes Then
    End
    Unload frmSYSTRAYICON
End If
End Sub
Private Sub mnuHomeDelivery_Click()
    frmHDBill.Show
End Sub
Private Sub mnuMasterAdvanceBooking_Click()
    frmAdvancedBooking.Show
End Sub

Private Sub mnuMasterCategory_Click()
    frmCategory.Show
End Sub

Private Sub mnuMasterItemTable_Click()
    frmItemTable.Show
End Sub
Private Sub mnuMasterOrder_Click()
    frmOrder.Show
End Sub
Private Sub mnuMasterTables_Click()
    frmTables.Show
End Sub
Private Sub MDIfrom_unload(Cancel As Integer)
    CloseConnection
End Sub
Private Sub mnuMonthlyCollection_Click()
    frmMCollection.Show
End Sub
Private Sub mnuReportMenu_Click()
    DRPMenu.Show
End Sub

Private Sub mnuSideBar_Click()
If Picture1.Visible = True Then
Picture1.Visible = False
mnuSideBar.Checked = False
ElseIf Picture1.Visible = False Then
Picture1.Visible = True
mnuSideBar.Checked = True
End If
End Sub

Private Sub mnuViewHelp_Click()
cdlg.ShowHelp
End Sub

Private Sub Timer1_Timer()
Label1.Caption = DateTime.Time
End Sub

Private Sub Timer2_Timer()
Label.FontItalic = True
Label.FontUnderline = True
Label.ForeColor = &HFF&
Label.Left = Label.Left + 10
If Label.Left >= 3500 Then
Label.Left = Label.Left + 10
If Label.Left = 16400 Then
Timer2.Enabled = False
Timer3.Enabled = True
End If
End If
End Sub

Private Sub Timer3_Timer()
Label.FontBold = True
Label.FontUnderline = False
Label.ForeColor = &HFF0000
Label.Left = Label.Left - 10
If Label.Left <= 16800 Then
Label.Left = Label.Left - 10
If Label.Left = 3500 Then
Timer3.Enabled = False
Timer2.Enabled = True
End If
End If
End Sub
