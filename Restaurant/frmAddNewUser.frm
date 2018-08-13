VERSION 5.00
Begin VB.Form frmAddNewUser 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   3105
   ClientLeft      =   7860
   ClientTop       =   4230
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdClose 
      Height          =   705
      Left            =   3000
      Picture         =   "frmAddNewUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   750
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   705
      Left            =   1080
      Picture         =   "frmAddNewUser.frx":2674
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   750
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   360
      Picture         =   "frmAddNewUser.frx":4CE8
      ScaleHeight     =   225
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   120
      Width           =   1425
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Add New User"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.Label lblPassword4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblConfirmPassword4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm Password :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblUsername4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Username:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   1875
      Left            =   120
      Top             =   240
      Width           =   4635
   End
End
Attribute VB_Name = "frmAddNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload frmAddNewUser
End Sub

Private Sub cmdSave_Click()
    If txtUsername.Text = "" Then
        MsgBox "Enter UserName and Password ...", vbExclamation
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text = "" Then
        MsgBox "Enter Password ...", vbExclamation
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If txtConfirmPassword.Text = "" Then
    MsgBox "Enter confirpassword ...", vbExclamation
    txtConfirmPassword.SetFocus
    Exit Sub
    End If
    
    If txtPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Confirm password dosenot match with new password ...", vbExclamation
        txtConfirmPassword.Text = ""
        txtPassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    End If

Set rs = cn.Execute("select * from Login where User_name='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Sorry!! User already exists. Try another username", vbCritical
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
Else
    cn.Execute ("insert into Login values('" + txtUsername.Text + "','" + txtPassword.Text + "')")
    MsgBox "User added sucessfully", vbInformation
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
   frmAddNewUser.Hide
   mdiMain.Show
End If
End Sub

Private Sub Form_Load()
OpenConnection
End Sub
