VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3120
   ClientLeft      =   7860
   ClientTop       =   4440
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5040
   Begin VB.TextBox txtConfirmPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1650
      Width           =   2310
   End
   Begin VB.TextBox txtNewPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   2295
   End
   Begin VB.TextBox txtCurrentPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2310
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   360
      Picture         =   "frmChangePassword.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   1785
      TabIndex        =   5
      Top             =   120
      Width           =   1785
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdClose 
      Height          =   705
      Left            =   2760
      Picture         =   "frmChangePassword.frx":09D2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   750
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   705
      Left            =   840
      Picture         =   "frmChangePassword.frx":3046
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   240
      TabIndex        =   8
      Top             =   1140
      Width           =   1890
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Old Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1770
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   1875
      Left            =   120
      Top             =   240
      Width           =   4755
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload frmChangePassword
End Sub
Private Sub cmdSave_Click()
    If txtCurrentPassword.Text = "" Then
        MsgBox "Please Enter Old Password ...", vbExclamation
        txtCurrentPassword.SetFocus
        Exit Sub
    End If
    
    If txtNewPassword.Text = "" Then
        MsgBox "Enter New Password ...", vbExclamation
        txtNewPassword.SetFocus
        Exit Sub
    End If
    
    If txtConfirmPassword.Text = "" Then
    MsgBox "Enter confirm password ...", vbExclamation
    txtConfirmPassword.SetFocus
    Exit Sub
    End If
    
    If txtNewPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Confirm password dosenot match with new password ...", vbExclamation
        txtConfirmPassword.Text = ""
        txtNewPassword.Text = ""
        txtNewPassword.SetFocus
        Exit Sub
    End If
Set rs = New ADODB.Recordset
rs.Open "select * from Login where User_name", cn, adOpenForwardOnly, adLockPessimistic

    If Not rs.EOF Then
      If rs.Fields("Password").Value = txtCurrentPassword.Text Then
            If txtNewPassword.Text = txtConfirmPassword.Text Then
            
                rs.Fields("Password").Value = txtNewPassword.Text
                rs.Update
                rs.Requery
                MsgBox "Password succesfully Changed...", vbExclamation
                txtCurrentPassword.Text = ""
                txtNewPassword.Text = ""
                txtConfirmPassword.Text = ""
                frmChangePassword.Hide
                mdiMain.Show
                End If
                Else
                MsgBox "Incorrect old password", vbExclamation
                txtCurrentPassword.Text = ""
                txtNewPassword.Text = ""
                txtConfirmPassword.Text = ""
                txtCurrentPassword.SetFocus
                End If
                End If
                End Sub
    
   
