VERSION 5.00
Begin VB.Form frmCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAddress.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   8910
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6960
      TabIndex        =   17
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigation"
      Height          =   1335
      Left            =   360
      TabIndex        =   27
      Top             =   6120
      Width           =   6495
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   495
         Index           =   3
         Left            =   4800
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   495
         Index           =   1
         Left            =   1760
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   495
         Index           =   2
         Left            =   3280
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   5295
      Left            =   6240
      TabIndex        =   26
      Top             =   840
      Width           =   2175
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1060
         Width           =   1575
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1760
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&update"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2460
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   3180
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3860
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&AddNew"
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraAddress 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Customer Address"
      Height          =   5295
      Left            =   360
      TabIndex        =   19
      Top             =   840
      Width           =   5895
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   615
         Index           =   0
         Left            =   2280
         TabIndex        =   1
         Top             =   270
         Width           =   3375
      End
      Begin VB.TextBox txtInput 
         Height          =   615
         Index           =   2
         Left            =   2310
         TabIndex        =   3
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtInput 
         Height          =   615
         Index           =   3
         Left            =   2280
         TabIndex        =   4
         Top             =   2790
         Width           =   3375
      End
      Begin VB.TextBox txtInput 
         Height          =   615
         Index           =   4
         Left            =   2280
         TabIndex        =   5
         Top             =   3630
         Width           =   3375
      End
      Begin VB.TextBox txtInput 
         Height          =   615
         Index           =   5
         Left            =   2280
         TabIndex        =   6
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtInput 
         Height          =   615
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Colony"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "House No."
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer No"
         Height          =   615
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Home Delivery"
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
      Height          =   615
      Left            =   2280
      TabIndex        =   18
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Option Explicit

Private Function clearControls()
    txtInput(0).Text = ""
    txtInput(1).Text = ""
    txtInput(2).Text = ""
    txtInput(3).Text = ""
    txtInput(4).Text = ""
    txtInput(5).Text = ""
End Function

Private Function EDControls(mode As Boolean)
    txtInput(1).Enabled = mode
    txtInput(2).Enabled = mode
    txtInput(3).Enabled = mode
    txtInput(4).Enabled = mode
    txtInput(5).Enabled = mode
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
    Dim i As Integer
    txtInput(0).Text = rec.Fields("C_No").Value
    txtInput(1).Text = rec.Fields("C_Name").Value
    txtInput(2).Text = rec.Fields(2).Value
    txtInput(3).Text = rec.Fields("Colony").Value
    txtInput(4).Text = rec.Fields("City").Value
    txtInput(5).Text = rec.Fields("Ph_no").Value
End Function

Private Sub cmdAddNew_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = True
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = True
    cmdFind.Enabled = False
    clearControls
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * From Customer", cn, adOpenKeyset, adLockOptimistic
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
    cno = rec.Fields("c_no").Value
    Unload Me
End Sub
    Private Function loadCName()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * from Customer", cn, adOpenKeyset, adLockOptimistic
    cboCName.Clear
    If t.RecordCount > 0 Then
        While Not t.EOF
            cboCName.AddItem t.Fields("C_Name").Value
            cboCName.ItemData(cboCName.NewIndex) = t.Fields("C_No").Value
            t.MoveNext
        Wend
    End If
     t.MoveLast
     Set t = Nothing
End Function
Private Sub cmdFind_Click()
    Dim posi As Variant
    Dim s As String
    If rec.RecordCount > 0 Then
        s = InputBox("Enter The Customer name")
        If s <> "" Then
            posi = rec.Bookmark
            rec.Find "C_Name='" & s & "'", 0, adSearchForward, adBookmarkFirst
            If rec.EOF Then
                MsgBox "Record Not Found"
                rec.Bookmark = posi
            End If
            showData
        Else
            MsgBox "Value not specified for search"
        End If
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
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Or txtInput(2).Text = "" Or txtInput(3).Text = "" Or txtInput(4).Text = "" Or txtInput(5).Text = "" Then
        MsgBox "please provide the data"
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    rec.AddNew
    rec.Fields("C_No").Value = txtInput(0).Text
    rec.Fields("C_Name").Value = txtInput(1).Text
    rec.Fields(2).Value = txtInput(2).Text
    rec.Fields(3).Value = txtInput(3).Text
    rec.Fields(4).Value = txtInput(4).Text
    rec.Fields(5).Value = txtInput(5).Text
    rec.Update
    cmdFind.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Or txtInput(2).Text = "" Or txtInput(3).Text = "" Or txtInput(4).Text = "" Or txtInput(5).Text = "" Then
        MsgBox "Please provide the data"
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdUpdate.Enabled = False
    rec.Fields(0).Value = txtInput(0).Text
    rec.Fields(1).Value = txtInput(1).Text
    rec.Fields(2).Value = txtInput(2).Text
    rec.Fields(3).Value = txtInput(3).Text
    rec.Fields(4).Value = txtInput(4).Text
    rec.Fields(5).Value = txtInput(5).Text
    rec.Update
End Sub

Private Sub Form_Activate()
    CenterWindow Me
    rec.Open "Select * from Customer", cn, adOpenKeyset, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        showData
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    EDControls False
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
    txtInput(0).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Select Case UnloadMode
            Case vbFormCode
            Case vbFormControlMenu
            Case vbFormMDIForm
                MsgBox "First close the form"
                Cancel = True
            Case vbAppTaskManager
                MsgBox "First close the application then shut Down"
                Cancel = True
            Case vbAppWindows
                MsgBox "First close the Application"
                Cancel = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
            KeyAscii = 0
            MsgBox "Please Enter Proper Name"
            Exit Sub
        End If
    ElseIf Index = 3 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
            KeyAscii = 0
            MsgBox "Please Enter Proper Colony Name"
        End If
    ElseIf Index = 4 Then
        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32) Then
            KeyAscii = 0
            MsgBox "Please Enter Proper city Name"
        End If
    ElseIf Index = 5 Then
         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 45)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
        End If
    End If
End Sub
