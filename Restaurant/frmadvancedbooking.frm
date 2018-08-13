VERSION 5.00
Begin VB.Form frmAdvancedBooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Booking"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10140
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmadvancedbooking.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   10140
   Begin VB.Frame fraControls 
      BackColor       =   &H00E0E0E0&
      Height          =   4185
      Left            =   6210
      TabIndex        =   25
      Top             =   1200
      Width           =   3525
      Begin VB.TextBox txtRent 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1620
         TabIndex        =   5
         Top             =   2115
         Width           =   1665
      End
      Begin VB.ComboBox cboTableNo 
         Height          =   390
         Left            =   1620
         TabIndex        =   4
         Top             =   1650
         Width           =   1665
      End
      Begin VB.TextBox txtHours 
         Height          =   390
         Left            =   1650
         TabIndex        =   7
         Top             =   3030
         Width           =   1665
      End
      Begin VB.ComboBox cboTimeFrom 
         Height          =   390
         ItemData        =   "frmadvancedbooking.frx":2B11A
         Left            =   1650
         List            =   "frmadvancedbooking.frx":2B142
         TabIndex        =   6
         Top             =   2565
         Width           =   1665
      End
      Begin VB.CommandButton cmdAddCustomer 
         Caption         =   "&Add Customer"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1110
         Width           =   3165
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   405
         Left            =   1770
         TabIndex        =   8
         Top             =   3600
         Width           =   1635
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   405
         Left            =   150
         TabIndex        =   16
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rent"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2130
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Table No"
         Height          =   435
         Left            =   120
         TabIndex        =   29
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label lblTimeFrom 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time From"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblHours 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   3060
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   9645
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4215
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   6105
      Begin VB.ComboBox cboCName 
         Height          =   390
         Left            =   2880
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ListBox lstTables 
         Height          =   1140
         ItemData        =   "frmadvancedbooking.frx":2B16D
         Left            =   1500
         List            =   "frmadvancedbooking.frx":2B16F
         MultiSelect     =   2  'Extended
         TabIndex        =   14
         Top             =   2280
         Width           =   1395
      End
      Begin VB.ListBox lstTimeStart 
         Height          =   1140
         ItemData        =   "frmadvancedbooking.frx":2B171
         Left            =   2910
         List            =   "frmadvancedbooking.frx":2B173
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   2280
         Width           =   2955
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   2
         Left            =   2880
         TabIndex        =   9
         Top             =   3630
         Width           =   2985
      End
      Begin VB.TextBox txtInput 
         Height          =   390
         Index           =   1
         Left            =   2910
         MaxLength       =   8
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtInput 
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   2880
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblTimeStart 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Start "
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2880
         TabIndex        =   26
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tables"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   24
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Advanced Payment"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Booking No"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Booking"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2280
      TabIndex        =   17
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmAdvancedBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim FirstTime As Boolean

Private Function EDControls(mode As Boolean)
    txtInput(1).Enabled = mode
    txtInput(2).Enabled = mode
    txtHours.Enabled = mode
    cboCName.Enabled = mode
    lstTables.Enabled = mode
    cmdRemove.Enabled = mode
    cmdAdd.Enabled = mode
    cboTableNo.Enabled = mode
    cmdAddNew.Enabled = Not mode
    cmdCancel.Enabled = mode
End Function

Private Sub cboTableNo_Click()
    If (cboTableNo.List(cboTableNo.ListIndex)) Then
        txtRent.Text = cboTableNo.ItemData(cboTableNo.ListIndex)
    End If
    cmdAdd.Enabled = True
End Sub

Private Sub cmAddCustomer_Click()
    frmCustomer.Show
End Sub

Private Sub cmdAdd_Click()
        If txtHours.Text = "" Then
            MsgBox "Please Provide Hours"
            txtHours.SetFocus
            Exit Sub
        End If
        
        lstTables.AddItem (cboTableNo.List(cboTableNo.ListIndex))
        lstTables.ItemData(lstTables.NewIndex) = cboTableNo.ItemData(cboTableNo.ListIndex)
        lstTimeStart.AddItem (cboTimeFrom.List(cboTimeFrom.ListIndex))
        lstTimeStart.ItemData(lstTimeStart.NewIndex) = txtHours.Text
        cboTableNo.RemoveItem (cboTableNo.ListIndex)
        cboTimeFrom.ListIndex = -1
        txtRent.Text = ""
        txtHours.Text = ""
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
End Sub

Private Sub cmdAddCustomer_Click()
    frmCustomer.Show
End Sub

Private Sub cmdAddNew_Click()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * From Advance_Booking", cn, adOpenKeyset, adLockOptimistic
    If t.RecordCount > 0 Then
        t.MoveLast
        txtInput(0).Text = t.Fields(0).Value + 1
    Else
        txtInput(0).Text = 1
    End If
    Set t = Nothing
    lstTables.Clear
    lstTimeStart.Clear
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    EDControls True
    cmdAddCustomer.Enabled = True
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    txtInput(2).Text = ""
    txtInput(1).Text = Format(Date, "dd-MM-yy")
    cboTimeFrom.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    txtInput(1).Enabled = False
    txtInput(2).Enabled = False
    cmdSave.Enabled = False
    EDControls False
    txtInput(2).Text = ""
     Dim i As Integer
     
    For i = lstTables.ListCount - 1 To 0 Step -1
            cboTableNo.AddItem (lstTables.List(i))
            cboTableNo.ItemData(cboTableNo.NewIndex) = lstTables.ItemData(i)
            lstTables.RemoveItem (i)
    Next
    lstTables.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    
    For i = lstTables.ListCount - 1 To 0 Step -1
        If lstTables.Selected(i) = True Then
            cboTableNo.AddItem (lstTables.List(i))
            cboTableNo.ItemData(cboTableNo.NewIndex) = lstTables.ItemData(i)
            lstTables.RemoveItem (i)
            lstTimeStart.RemoveItem (i)
        End If
    Next
    cmdRemove.Enabled = False
End Sub

Private Sub cmdSave_Click()
    Dim arr() As String
    If txtInput(1).Text = "" Then
            MsgBox "Date Is required"
            Exit Sub
    End If
    
    arr = Split(txtInput(1).Text, "-")
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
    
    If txtInput(2).Text = "" Then
        MsgBox "Please Provide The Data"
        Exit Sub
    End If
    
    If txtInput(2).Text <= 0 Then
        MsgBox "Please Enter Positive Value"
        Exit Sub
    End If
    
    If Not IsNumeric(txtInput(2).Text) Then
        MsgBox "Please Enter Numeric Data"
        Exit Sub
    End If
    
    If lstTables.ListCount = 0 Then
        MsgBox "Please Allocate the Tables"
        Exit Sub
    End If
    
    If cboCName.Text = "" Then
        MsgBox "Please Enter Customer Name"
        Exit Sub
    End If
    
    Dim rec1 As ADODB.Recordset
    Dim j As Integer
    Dim rest As ADODB.Recordset
    Set rest = New ADODB.Recordset
    rest.Open "select * from Customer", cn, adOpenKeyset, adLockOptimistic
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    rec.AddNew
    rec.Fields("Booking_No").Value = txtInput(0).Text
    rec.Fields("Booking_Date").Value = txtInput(1).Text
    rec.Fields("Adv_payment").Value = txtInput(2).Text
    rec.Fields("No_Of_Tables").Value = lstTables.ListCount
    rec.Fields("C_No").Value = cboCName.ItemData(cboCName.ListIndex)
    rec.Update
    Set rec1 = New ADODB.Recordset
    rec1.Open " select * from Booking_table", cn, adOpenKeyset, adLockOptimistic
    
    For j = 0 To lstTables.ListCount - 1
        rec1.AddNew
        rec1.Fields(0).Value = txtInput(0).Text
        rec1.Fields(1).Value = lstTables.List(j)
        rec1.Fields(2).Value = lstTables.ItemData(j)
        rec1.Fields(3).Value = lstTimeStart.List(j)
        rec1.Fields(4).Value = lstTimeStart.ItemData(j)
        rec1.Update
    Next
    
    Set rec1 = Nothing
    EDControls False
End Sub

Private Sub Form_Activate()
Dim i As Integer
    
    If FirstTime Then
        CenterWindow Me
        loadCName
        loadTableNo
        rec.Open "Select * from Advance_Booking", cn, adOpenKeyset, adLockOptimistic
        cmdAddNew_Click
        EDControls False
        cmdSave.Enabled = False
        cmdAddCustomer.Enabled = False
        FirstTime = False
    Else
        If cno > 0 Then
            loadCName
            
            For i = 0 To cboCName.ListCount - 1
                If cboCName.ItemData(i) = cno Then
                    cboCName.ListIndex = i
                    Exit For
                End If
            Next
            
            cno = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    txtRent.Enabled = False
    txtInput(0).Enabled = False
    Set rec = New ADODB.Recordset
    EDControls False
    FirstTime = True
    cboTimeFrom.Enabled = False
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
     Set t = Nothing
End Function

Private Function loadTableNo()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    Dim h As Integer
    h = Hour(Time)
    t.Open "select * from Tables where Table_No not in(select distinct Table_No from Booking_Table where time_from < " & h & " or  time_from + hours > " & h & " and booking_no  in (Select booking_no from advance_booking where booking_date =#" & Format(Date, "dd-mm-yy") & "#))", cn, adOpenKeyset, adLockOptimistic
    cboTableNo.Clear
    
    If t.RecordCount > 0 Then
            While Not t.EOF
                cboTableNo.AddItem t.Fields(0).Value
                cboTableNo.ItemData(cboTableNo.NewIndex) = t.Fields(1).Value
                t.MoveNext
            Wend
    End If
     Set t = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

Private Sub lstTables_Click()
    If lstTables.ListIndex >= 0 Then
        cmdRemove.Enabled = True
    End If
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 45 Or KeyAscii = 8) Then
            KeyAscii = 0
            MsgBox "Please Enter Date In DD-MM-YY Formate"
            Exit Sub
        End If
    End If
End Sub
