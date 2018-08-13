VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Form"
   ClientHeight    =   4320
   ClientLeft      =   6810
   ClientTop       =   4020
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   6405
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   105
      Left            =   2400
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1200
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   5025
      Begin VB.Frame fraLogin 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   3285
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   435
            Left            =   1800
            TabIndex        =   3
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton cmdLogin 
            Caption         =   "&Login"
            Default         =   -1  'True
            Height          =   435
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   930
         Width           =   1785
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "User Name"
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   3900
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   5445
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdLogin_Click()
    If txtInput(0).Text = "" And txtInput(1).Text = "" Then
       MsgBox "Please Enter User Name and Password"
       Exit Sub
    End If
    If rec.RecordCount > 0 Then
       rec.MoveFirst
       rec.Find "User_name='" & txtInput(0).Text & "'"
       If rec.EOF Then
             MsgBox "Incorrect username", vbOKOnly
             txtInput(0).Text = ""
             txtInput(1).Text = ""
             txtInput(0).SetFocus
             Exit Sub
       Else
          If rec.Fields("password").Value <> txtInput(1).Text Then
             MsgBox "Incorrect password", vbOKOnly
             txtInput(0).Text = ""
             txtInput(1).Text = ""
             txtInput(0).SetFocus
            Exit Sub
          End If
       End If
    End If
    Timer1.Enabled = True
    ProgressBar1.Visible = True
    Shape1.Visible = True
    Shape2.Visible = True
    Shape3.Visible = True
    Label2.Visible = True
    Label3.Visible = True
End Sub

Private Sub Form_Activate()
    txtInput(0).SetFocus
End Sub

Private Sub Form_Load()
    If Not OpenConnection Then
        MsgBox "Error while connecting with this database"
        End
    End If
    Set rec = New ADODB.Recordset
    rec.Open "Select * from Login", cn, adOpenKeyset, adLockOptimistic
    
End Sub

Private Sub Timer1_Timer()
      If Shape1.Visible Then
         Shape2.Visible = True
         Shape1.Visible = False
         Shape3.Visible = False
      ElseIf Shape2.Visible Then
         Shape3.Visible = True
         Shape2.Visible = False
         Shape1.Visible = False
      ElseIf Shape3.Visible Then
         Shape1.Visible = True
         Shape2.Visible = False
         Shape3.Visible = False
      End If
    ProgressBar1.Value = ProgressBar1.Value + 5
    Label2.Caption = "Loading"
    Label3.Caption = ProgressBar1.Value & "%"
      If (ProgressBar1.Value = ProgressBar1.Max) Then
         Timer1.Enabled = False
         mdiMain.Show
         mdiMain.Label18.Caption = txtInput(0).Text
         Unload Me
      End If
End Sub
