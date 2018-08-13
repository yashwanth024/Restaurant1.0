VERSION 5.00
Begin VB.Form frmDeleteUser 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete User"
   ClientHeight    =   2580
   ClientLeft      =   8070
   ClientTop       =   4650
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4530
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "Delete"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbUsername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   2220
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1560
      Picture         =   "frmDeleteUser.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   120
      Width           =   1425
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Delete User"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.Label lblUsername4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Username:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   2115
      Left            =   120
      Top             =   240
      Width           =   4155
   End
End
Attribute VB_Name = "frmDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDeleteUser_Click()
cn.Execute ("delete from Login where User_name='" + cmbUsername.Text + "'")
MsgBox "User deleted sucessfully!!", vbInformation
cmbUsername.Text = ""
Unload frmDeleteUser
mdiMain.Show
End Sub

Private Sub cmdExit_Click()
Unload frmDeleteUser
End Sub

Private Sub Form_Load()
OpenConnection
Set rs = cn.Execute("select * from Login")
While (Not rs.EOF)
    cmbUsername.AddItem rs(0)
    rs.MoveNext
Wend
End Sub


