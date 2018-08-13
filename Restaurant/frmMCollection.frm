VERSION 5.00
Begin VB.Form frmMCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Collection Report"
   ClientHeight    =   2970
   ClientLeft      =   8070
   ClientTop       =   6570
   ClientWidth     =   5445
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMCollection.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   5445
   Begin VB.TextBox txtInput 
      Height          =   390
      Index           =   1
      Left            =   2160
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   390
      Index           =   0
      Left            =   2160
      MaxLength       =   12
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Collection "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmMCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
     Dim arr() As String
    If txtInput(0).Text = "" Or txtInput(0).Text = "" Then
            MsgBox "Date Is required"
            Exit Sub
    End If
    arr = Split(txtInput(0).Text, "-")
    If UBound(arr) < 2 Then
        MsgBox "please Enter complete Date"
        Exit Sub
    End If
    If Not (arr(0) >= 1 And arr(0) <= 31) Then
        MsgBox "Please Enter valid Date between 1 to 31"
        
        Exit Sub
    End If
     If Not (arr(1) >= 1 And arr(1) <= 12) Then
        MsgBox "Please Enter valid Month between 1 to 12"
        Exit Sub
    End If
     If Not IsNumeric(arr(2)) Then
        MsgBox "Please Enter valid Year"
        Exit Sub
    End If
    
     arr1 = Split(txtInput(1).Text, "-")
    If UBound(arr1) < 2 Then
        MsgBox "please Enter complete Date"
        Exit Sub
    End If
    If Not (arr1(0) >= 1 And arr1(0) <= 31) Then
        MsgBox "Please Enter valid Date between 1 to 31"
        
        Exit Sub
    End If
     If Not (arr1(1) >= 1 And arr1(1) <= 12) Then
        MsgBox "Please Enter valid Month between 1 to 12"
        Exit Sub
    End If
     If Not IsNumeric(arr1(2)) Then
        MsgBox "Please Enter valid Year"
        Exit Sub
    End If
    If DataEnvironment1.rscmdMCollection_Grouping.State Then
        DataEnvironment1.rscmdMCollection_Grouping.Close
    End If
    DataEnvironment1.cmdMCollection_Grouping CDate(txtInput(0).Text), CDate(txtInput(1).Text)
    DRMCollection.Show
End Sub

Private Sub Form_Load()
    CenterWindow Me
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 45 Or KeyAscii = 8) Then
            KeyAscii = 0
            MsgBox "Please Enter Date In DD-MM-YY Formate"
            Exit Sub
            End If
End Sub
