VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   6600
   ClientTop       =   4230
   ClientWidth     =   9660
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4060.369
   ScaleMode       =   0  'User
   ScaleWidth      =   9074.258
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   7320
      MaskColor       =   &H00FFFF00&
      TabIndex        =   0
      Top             =   5280
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer:-"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Ravindra Nath"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Yogesh Kumar Jailiya"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Satyen kumar jatiya"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Garam     Masala         Restaurant"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3000
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   5892
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.543
      X2              =   5309.286
      Y1              =   2195.148
      Y2              =   2195.148
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   169.086
      X2              =   8479.64
      Y1              =   2174.484
      Y2              =   2174.484
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblTitle.Caption = "RESTURANT MANAGEMENT "
End Sub

