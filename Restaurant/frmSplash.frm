VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7935
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   7935
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Zameer Khan"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   705
      Left            =   6690
      TabIndex        =   3
      Top             =   6270
      Width           =   4275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Suresh Bhadana"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   705
      Left            =   6720
      TabIndex        =   2
      Top             =   5550
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By:-"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   4080
      TabIndex        =   1
      Top             =   4710
      Width           =   3795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Garam masala management"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   150
      Width           =   8775
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmLogin.Show
End Sub

Private Sub Form_Load()
    OpenConnection
End Sub

Private Sub Timer1_Timer()
        Unload Me
        frmLogin.Show
End Sub
