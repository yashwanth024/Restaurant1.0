VERSION 5.00
Begin VB.Form frmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category"
   ClientHeight    =   5640
   ClientLeft      =   7020
   ClientTop       =   4650
   ClientWidth     =   8160
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
   Picture         =   "frmCategory.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigation"
      Height          =   1215
      Left            =   480
      TabIndex        =   18
      Top             =   3630
      Width           =   5775
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   495
         Index           =   3
         Left            =   4350
         TabIndex        =   12
         Top             =   450
         Width           =   1365
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   495
         Index           =   2
         Left            =   2940
         TabIndex        =   11
         Top             =   450
         Width           =   1365
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   495
         Index           =   1
         Left            =   1530
         TabIndex        =   10
         Top             =   450
         Width           =   1365
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4665
      Left            =   6240
      TabIndex        =   17
      Top             =   840
      Width           =   1815
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&AddNew"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   955
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2205
         Width           =   1575
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1580
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2850
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   480
      TabIndex        =   14
      Top             =   870
      Width           =   5775
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   570
         Width           =   2295
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1740
         Width           =   2895
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Category ID"
         Height          =   495
         Index           =   0
         Left            =   330
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   19
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmCategory"
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

Private Sub cmdCancel_Click()
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
    showData
End Sub

Private Sub cmdDelete_Click()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "Select * from Item_Table  where Cat_ID=" & txtInput(0).Text, cn, adOpenKeyset, adLockOptimistic
    
    If t.RecordCount > 0 Then
            MsgBox "You Can not Delete the Record because Related Records Exists in Item Table"
            Exit Sub
    End If
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
        Set t = Nothing
End Sub

Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub cmdFind_Click()
Dim s As String
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        s = InputBox("Enter The Category name")
        If s <> "" Then
            rec.Find "Cat_Name='" & s & "'"
            If Not rec.EOF Then
            Else
                MsgBox "Record Not Found"
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
        Exit Sub
    End If
    
    If Not IsNumeric(txtInput(0).Text) Then
        MsgBox "Price must be Numeric value"
        Exit Sub
    End If
    
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    rec.AddNew
    rec.Fields("cat_id").Value = txtInput(0).Text
    rec.Fields("cat_name").Value = txtInput(1).Text
    rec.Update
End Sub

Private Sub cmdUpdate_Click()
     If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "Please provide the data"
        Exit Sub
    End If
    
    If Not IsNumeric(txtInput(0).Text) Then
        MsgBox "Price must be Numeric value"
        Exit Sub
    End If
    
    EDControls False
    EDNavigate True
    cmdUpdate.Enabled = False
    rec.Fields("cat_id").Value = txtInput(0).Text
    rec.Fields("cat_name").Value = txtInput(1).Text
    rec.Update
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    rec.Open "Select * from Category", cn, adOpenKeyset, adLockOptimistic
    EDControls False
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
    txtInput(0).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Query
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

Private Function showData()
    txtInput(0).Text = rec.Fields("cat_id").Value
    txtInput(1).Text = rec.Fields("cat_name").Value
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
    t.Open "select * From Category", cn, adOpenKeyset, adLockOptimistic
    If t.RecordCount > 0 Then
        t.MoveLast
        txtInput(0).Text = t.Fields(0).Value + 1
    Else
        txtInput(0).Text = 1
    End If
    Set t = Nothing
   txtInput(1).SetFocus
End Sub

Private Function loadCategory()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * from Category", cn, adOpenKeyset, adLockOptimistic
    
    If t.RecordCount > 0 Then
        While Not t.EOF
            txtInput(0).Text = t.Fields("cat_id").Value
            txtInput(1).Text = t.Fields("cat_name").Value
            t.MoveNext
        Wend
    End If
     Set t = Nothing
End Function

Private Sub Form_Activate()
    CenterWindow Me
    loadCategory
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        showData
    End If
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 1 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
            KeyAscii = 0
            MsgBox "Please Enter Proper Category Name"
        End If
    End If
End Sub
