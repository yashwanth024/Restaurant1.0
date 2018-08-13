VERSION 5.00
Begin VB.Form frmItemTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5310
   ClientLeft      =   6165
   ClientTop       =   3585
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmItemTable.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigation"
      Height          =   1125
      Left            =   720
      TabIndex        =   15
      Top             =   3360
      Width           =   6375
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   465
         Index           =   3
         Left            =   4800
         TabIndex        =   11
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   465
         Index           =   2
         Left            =   3250
         TabIndex        =   10
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   465
         Index           =   1
         Left            =   1700
         TabIndex        =   9
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   465
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   450
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3975
      Left            =   7080
      TabIndex        =   14
      Top             =   960
      Width           =   1815
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1725
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   1575
      End
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
      Height          =   2385
      Left            =   720
      TabIndex        =   13
      Top             =   960
      Width           =   6375
      Begin VB.ComboBox cboCatID 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2340
         TabIndex        =   1
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   1
         Left            =   2310
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   0
         Left            =   2325
         TabIndex        =   2
         Top             =   990
         Width           =   3135
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1920
      TabIndex        =   19
      Top             =   240
      Width           =   5715
   End
End
Attribute VB_Name = "frmItemTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Function clearControls()
    txtInput(0).Text = ""
    txtInput(1).Text = ""
    If cboCatID.ListCount > 0 Then
        cboCatID.ListIndex = 0
    End If
End Function

Private Function EDControls(mode As Boolean)
    txtInput(0).Enabled = mode
    txtInput(1).Enabled = mode
    cboCatID.Enabled = mode
    cmdAddNew.Enabled = Not mode
    cmdCancel.Enabled = mode
    cmdFind.Enabled = Not mode
    cmdDelete.Enabled = Not mode
End Function

Private Function EDNavigate(mode As Boolean)
    fraNavigator.Enabled = mode
End Function

Private Function showData()
    Dim i As Integer
    txtInput(0).Text = rec.Fields("Item_name").Value
    txtInput(1).Text = rec.Fields("rate").Value
    For i = 0 To cboCatID.ListCount - 1
        If rec.Fields("Cat_id").Value = cboCatID.ItemData(i) Then
            cboCatID.ListIndex = i
            Exit For
        End If
    Next
End Function

Private Sub cmdAddNew_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    clearControls
    txtInput(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
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
        s = InputBox("Enter The Item Name")
        If s <> "" Then
            rec.Find "item_name='" & s & "'"
            If Not rec.EOF Then
            Else
                MsgBox "Item Not Found"
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
    txtInput(0).SetFocus
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
    
    If Not IsNumeric(txtInput(1).Text) Then
        MsgBox "Price must be Numeric value"
        Exit Sub
    End If
    
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    rec.AddNew
    rec.Fields("Item_name").Value = txtInput(0).Text
    rec.Fields("Rate").Value = txtInput(1).Text
    rec.Fields("Cat_id").Value = cboCatID.ItemData(cboCatID.ListIndex)
    rec.Update
End Sub

Private Sub cmdUpdate_Click()
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "Please provide the data"
        Exit Sub
    End If
    
    If Not IsNumeric(txtInput(1).Text) Then
        MsgBox "Price must be Numeric And Positive value"
        Exit Sub
    End If
    
    EDControls False
    EDNavigate True
    cmdUpdate.Enabled = False
    rec.Fields("Item_name").Value = txtInput(0).Text
    cn.Execute "update "
    rec.Fields("Price").Value = txtInput(1).Text
    rec.Fields("Cat_id").Value = cboCatID.ItemData(cboCatID.ListIndex)
    rec.Update
End Sub

Private Sub Form_Activate()
    CenterWindow Me
    loadCategory
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        showData
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    rec.Open "Select * from Item_Table", cn, adOpenKeyset, adLockOptimistic
    EDControls False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

Private Function loadCategory()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * from Category", cn, adOpenKeyset, adLockOptimistic
    cboCatID.Clear
    If t.RecordCount > 0 Then
        While Not t.EOF
            cboCatID.AddItem t.Fields("cat_name").Value
            cboCatID.ItemData(cboCatID.NewIndex) = t.Fields("cat_id").Value
            t.MoveNext
        Wend
    End If
     Set t = Nothing
End Function

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
            KeyAscii = 0
            MsgBox "Please Enter Proper Item Name"
        End If
    ElseIf Index = 1 Then
         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
        End If
    End If
End Sub
