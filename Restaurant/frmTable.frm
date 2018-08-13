VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tables"
   ClientHeight    =   5565
   ClientLeft      =   6810
   ClientTop       =   4650
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmTable.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigation"
      Height          =   1455
      Left            =   480
      TabIndex        =   19
      Top             =   2760
      Width           =   5775
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   495
         Index           =   3
         Left            =   4320
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   6240
      TabIndex        =   18
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3135
         Width           =   1455
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2565
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1395
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   825
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&AddNew"
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1965
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   5835
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   1
         Left            =   2430
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Rent"
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Table No."
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   570
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Table Information"
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
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Function clearControls()
    txtInput(0).Text = ""
    txtInput(1).Text = ""
End Function

Private Function EDControls(mode As Boolean)
    txtInput(1).Enabled = mode
    cmdAddNew.Enabled = Not mode
    cmdCancel.Enabled = mode
    cmdModify.Enabled = Not mode
    cmdFind.Enabled = Not mode
    cmdDelete.Enabled = Not mode
End Function

Private Function EDNavigate(mode As Boolean)
    fraNavigator.Enabled = mode
End Function

Private Function showData()
    txtInput(0).Text = rec.Fields("Table_No").Value
    txtInput(1).Text = rec.Fields("Rent").Value
End Function

Private Sub cmdAddNew_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = True
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = True
    clearControls
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * From Tables", cn, adOpenKeyset, adLockOptimistic
    If t.RecordCount > 0 Then
        t.MoveLast
        txtInput(0).Text = t.Fields(0).Value + 1
    Else
        txtInput(0).Text = 1
    End If
    Set t = Nothing
     txtInput(1).SetFocus
End Sub

Private Sub cmdCancel_Click()
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
    showData
End Sub

Private Sub cmdDelete_Click()
    Dim choice As Integer
    choice = MsgBox("Do you want to Delete the Record", vbYesNo + vbQuestion, "confirmation")
    If choice = vbYes Then
        If rec.EOF = False And rec.BOF = False Then
            rec.Delete
            rec.MoveNext
            If rec.EOF Then rec.MoveLast
            showData
        End If
        End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim s As String
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        s = InputBox("Enter The table no")
        If s <> "" Then
            rec.Find "Table_No='" & s & "'"
            If Not rec.EOF Then
            Else
                MsgBox "Table Not Found"
                rec.MoveLast
            End If
        End If
        showData
    End If

End Sub

Private Sub cmdModify_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = False
    cmdUpdate.Enabled = True
    txtInput(1).SetFocus
End Sub

Private Sub cmdNavigate_click(Index As Integer)
    If Index = 0 Then
        rec.MovePrevious
        If (rec.BOF = True) Then
            MsgBox "You are already at the First Record"
        End If
        
        rec.MoveFirst
    ElseIf Index = 1 Then
        rec.MovePrevious
        If rec.BOF Then
              MsgBox "You are already at the First Record"
              rec.MoveFirst
        End If
                
    ElseIf Index = 2 Then
        rec.MoveNext
        If rec.EOF Then
            MsgBox "you are already at the last record"
            rec.MoveLast
          End If
        
      Else
        rec.MoveNext
        If (rec.EOF = True) Then
            MsgBox "You are already at the last Record"
        End If
        rec.MoveLast
      End If
        showData
End Sub

Private Sub cmdSave_Click()
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "please provide the data"
        txtInput(1).SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtInput(1).Text) Then
        MsgBox "Price must be Numeric value"
        txtInput(1).Text = ""
        txtInput(1).SetFocus
        Exit Sub
    ElseIf txtInput(1).Text < 0 Then
         MsgBox "Price must be Positive value"
        txtInput(1).Text = ""
        txtInput(1).SetFocus
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    rec.AddNew
    rec.Fields("Table_No").Value = txtInput(0).Text
    rec.Fields("Rent").Value = txtInput(1).Text
    rec.Update
End Sub

Private Sub cmdUpdate_Click()
     If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "Please provide the data"
        Exit Sub
    End If
    
    If Not IsNumeric((txtInput(1).Text) Or (txtInput(1).Text)) Then
        MsgBox "Price must be Numeric value"
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdUpdate.Enabled = False
    rec.Fields("Table_No").Value = txtInput(0).Text
    rec.Fields("Rent").Value = txtInput(1).Text
    rec.Update
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    EDControls False
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
     txtInput(0).Enabled = False
End Sub

Private Sub Form_Activate()
    CenterWindow Me
    rec.Open "Select * from Tables", cn, adOpenKeyset, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        showData
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

