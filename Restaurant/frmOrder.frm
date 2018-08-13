VERSION 5.00
Begin VB.Form frmOrder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order"
   ClientHeight    =   8940
   ClientLeft      =   4695
   ClientTop       =   1035
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOrder.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11760
   Begin VB.Frame fraAction 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Action"
      Height          =   2475
      Left            =   9180
      TabIndex        =   30
      Top             =   4860
      Width           =   2205
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   465
         Left            =   360
         TabIndex        =   8
         Top             =   585
         Width           =   1470
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   465
         Left            =   360
         TabIndex        =   9
         Top             =   1350
         Width           =   1470
      End
   End
   Begin VB.Frame fraPayment 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment"
      Height          =   4065
      Left            =   9150
      TabIndex        =   26
      Top             =   810
      Width           =   2220
      Begin VB.TextBox txtNetPayment 
         Enabled         =   0   'False
         Height          =   390
         Left            =   330
         TabIndex        =   7
         Top             =   3315
         Width           =   1695
      End
      Begin VB.TextBox txtDiscount 
         Height          =   390
         Left            =   315
         TabIndex        =   6
         Top             =   2130
         Width           =   1695
      End
      Begin VB.TextBox txtAmount 
         Height          =   390
         Left            =   330
         TabIndex        =   5
         Text            =   "0"
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label lblNetPayment 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Payment"
         Height          =   435
         Left            =   360
         TabIndex        =   29
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblDiscount 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   435
         Left            =   360
         TabIndex        =   28
         Top             =   1545
         Width           =   1185
      End
      Begin VB.Label lblAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   435
         Left            =   360
         TabIndex        =   27
         Top             =   390
         Width           =   1155
      End
   End
   Begin VB.Frame fraItems 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Items Information"
      Height          =   2490
      Left            =   180
      TabIndex        =   22
      Top             =   4830
      Width           =   9045
      Begin VB.TextBox txtqty 
         Height          =   390
         Left            =   4875
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1665
         Width           =   1245
      End
      Begin VB.TextBox txtPrice 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2700
         TabIndex        =   12
         Top             =   1665
         Width           =   1365
      End
      Begin VB.ComboBox cboItems 
         Height          =   390
         Left            =   2685
         TabIndex        =   3
         Top             =   1185
         Width           =   3465
      End
      Begin VB.ComboBox cbocategory 
         Height          =   390
         Left            =   2700
         TabIndex        =   2
         Top             =   645
         Width           =   3450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   345
         Left            =   4215
         TabIndex        =   32
         Top             =   1725
         Width           =   825
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   525
         Left            =   300
         TabIndex        =   25
         Top             =   1695
         Width           =   1335
      End
      Begin VB.Label lblAddItems 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         Height          =   525
         Left            =   330
         TabIndex        =   24
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label lblCategory 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   525
         Left            =   330
         TabIndex        =   23
         Top             =   705
         Width           =   1335
      End
   End
   Begin VB.Frame fraTables 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tables  Information"
      Height          =   1875
      Left            =   150
      TabIndex        =   18
      Top             =   3015
      Width           =   9045
      Begin VB.TextBox txttime 
         Height          =   420
         Left            =   5310
         TabIndex        =   33
         Top             =   1035
         Width           =   1410
      End
      Begin VB.TextBox txtRent 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2685
         TabIndex        =   21
         Top             =   1035
         Width           =   2385
      End
      Begin VB.ComboBox cboTables 
         Height          =   390
         Left            =   120
         TabIndex        =   1
         Top             =   1050
         Width           =   2385
      End
      Begin VB.Label lblTimeFrom 
         BackStyle       =   0  'Transparent
         Caption         =   "Time From"
         Height          =   315
         Left            =   5250
         TabIndex        =   31
         Top             =   690
         Width           =   2355
      End
      Begin VB.Label lblRent 
         BackStyle       =   0  'Transparent
         Caption         =   "Rent"
         Height          =   405
         Left            =   2685
         TabIndex        =   20
         Top             =   675
         Width           =   2175
      End
      Begin VB.Label lblAddTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Tables"
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   690
         Width           =   2175
      End
   End
   Begin VB.Frame fraAddOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add Order"
      Height          =   2160
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   9045
      Begin VB.CommandButton cmdAddCustomer 
         Caption         =   "Add Customer"
         Height          =   330
         Left            =   5490
         TabIndex        =   34
         Top             =   1440
         Width           =   1995
      End
      Begin VB.ComboBox cbocustomer 
         Height          =   390
         Left            =   2925
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   1395
         Width           =   2355
      End
      Begin VB.TextBox txtOrderNo 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2880
         TabIndex        =   10
         Top             =   420
         Width           =   2385
      End
      Begin VB.TextBox txtOrderDate 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2910
         MaxLength       =   8
         TabIndex        =   11
         Top             =   900
         Width           =   2385
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   465
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   1455
         Width           =   1305
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
         Height          =   465
         Index           =   1
         Left            =   165
         TabIndex        =   16
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   465
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   390
         Width           =   1305
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Information"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3180
      TabIndex        =   13
      Top             =   60
      Width           =   4935
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstTime As Boolean
Dim rec As ADODB.Recordset
Private Sub cboCategory_click()
    loadItems
    txtPrice.Text = ""
End Sub

Private Sub cbocustomer_KeyPress(KeyAscii As Integer)
Call capsonly(KeyAscii)
End Sub

Private Sub cboTables_Click()
txtRent.Text = ""
Dim r As New Recordset
r.Open "select * from Tables where Table_No=" & Val(Trim(cboTables.Text)) & " ", cn, adOpenDynamic, adLockOptimistic
If r.EOF = True Then
MsgBox "There is no table "
Exit Sub
Else
txtRent.Text = r.Fields("Rent").Value
End If
End Sub

Private Sub cmdAddCustomer_Click()
Dim cname As String, r As New Recordset, cid As Integer
cname = UCase(InputBox("Enter Customer name"))
cbocustomer.AddItem cname
r.Open "select * from Customer", cn, adOpenDynamic, adLockOptimistic
If r.EOF = True Then
cid = 1
cnn.Execute "insert into Customer(C_No,Cname) values(" & cid & " ,'" & cname & "')"
Exit Sub
Else
r.MoveLast
cid = r.Fields("C_No").Value + 1
If cname = "" Then Exit Sub
cn.Execute "insert into Customer(C_No,C_name) values(" & cid & " ,'" & cname & "')"
End If
r.Close
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
      If Val(txtAmount.Text) <= Val(txtDiscount.Text) Then
            MsgBox "Please Enter Discount less than Amount"
            txtDiscount.SetFocus
            Exit Sub
        End If
    If cbocustomer.Text = "" Then
        MsgBox "Please Enter Customer Name"
        Exit Sub
    End If
    If txtDiscount.Text = "" Then
        MsgBox "Please Enter Discount"
        txtDiscount.SetFocus
        Exit Sub
    End If
    Dim rest As ADODB.Recordset
    Set rest = New ADODB.Recordset
    rest.Open "select * from morder", cn, adOpenKeyset, adLockOptimistic
    rest.AddNew
    rest.Fields("O_No").Value = txtOrderNo.Text
    rest.Fields("O_Date").Value = txtOrderDate.Text
    rest.Fields("C_Name").Value = cbocustomer.Text
    rest.Fields("Amount").Value = txtAmount.Text
    rest.Fields("Discount").Value = txtDiscount.Text
   rest.Update
   MsgBox "Save successfuly"
End Sub

Private Sub cmdUpdate_Click()
        If cbocustomer.Text = "" Then
        MsgBox "Please Enter Customer Name"
        Exit Sub
        End If
    If txtDiscount.Text = "" Then
        MsgBox "Please Enter Discount"
        txtDiscount.SetFocus
        Exit Sub
        End If
    If Val(txtAmount.Text) <= Val(txtDiscount.Text) Then
        MsgBox "Please Enter Discount less than Amount"
        txtDiscount.SetFocus
        Exit Sub
        End If
End Sub

Public Sub capsonly(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    Else
        KeyAscii = KeyAscii
    End If
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 txtNetPayment = Val(txtAmount.Text) - Val(txtDiscount.Text)
End Sub

Private Sub Form_Load()
    Call generateOID
    txttime.Text = Time()
     txtOrderDate.Text = Format(Date, "DD-MM-YY")
     Call loadCName
     Call loadTableNo
     Call loadCategory
     End Sub

Private Function loadTableNo()
    Dim t As ADODB.Recordset
    Dim h As Integer
    h = Hour(Time)
    Set t = New ADODB.Recordset
    t.Open "select * from Tables where Table_No not in(select distinct Table_No from Booking_Table where time_from < " & h & " or  time_from + hours > " & h & " and booking_no  in (Select booking_no from advance_booking where booking_date =#" & Format(Date, "dd-mm-yy") & "#))", cn, adOpenKeyset, adLockOptimistic
    cboTables.Clear
    If t.RecordCount > 0 Then
            While Not t.EOF
                cboTables.AddItem t.Fields(0).Value
                cboTables.ItemData(cboTables.NewIndex) = t.Fields(1).Value
                t.MoveNext
            Wend
    End If
    Set t = Nothing
    cboTables.ListIndex = 0
    End Function

Private Function loadCategory()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * from Category", cn, adOpenKeyset, adLockOptimistic
    cbocategory.Clear
    If t.RecordCount > 0 Then
        While Not t.EOF
            cbocategory.AddItem t.Fields("cat_name").Value
            cbocategory.ItemData(cbocategory.NewIndex) = t.Fields("cat_id").Value
            t.MoveNext
        Wend
    End If
     Set t = Nothing
     cbocategory.ListIndex = 0
    End Function

Private Function loadItems()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    If cbocategory.ListIndex >= 0 Then
        t.Open "select * from Item_Table where Cat_Id=" & cbocategory.ItemData(cbocategory.ListIndex), cn, adOpenKeyset, adLockOptimistic
        cboItems.Clear
        If t.RecordCount > 0 Then
            While Not t.EOF
                cboItems.AddItem t.Fields(0).Value
                cboItems.ItemData(cboItems.NewIndex) = t.Fields(1).Value
                t.MoveNext
            Wend
        End If
    End If
     Set t = Nothing
     cboItems.ListIndex = 0
End Function

Private Sub cboItems_Click()
    If (cboItems.ListIndex >= 0) Then
       txtPrice.Text = Val(cboItems.ItemData(cboItems.ListIndex))
    Else
        txtPrice.Text = ""
    End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rec = Nothing
End Sub

Private Function loadCName()
    Dim t As ADODB.Recordset
    Set t = New ADODB.Recordset
    t.Open "select * from customer", cn, adOpenKeyset, adLockOptimistic
    cbocustomer.Clear
    If t.RecordCount > 0 Then
        While Not t.EOF
            cbocustomer.AddItem t.Fields("C_Name").Value
            cbocustomer.ItemData(cbocustomer.NewIndex) = t.Fields("C_No").Value
            t.MoveNext
        Wend
    End If
    Set t = Nothing
    cbocustomer.ListIndex = 0
End Function

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
     If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscount_LostFocus()
    txtNetPayment = Val(txtAmount.Text) - Val(txtDiscount.Text)
End Sub

Private Function clearControls()
     cbocustomer.ListIndex = -1
    optConsumed.Value = True
    LstTables.Clear
    lstStartTime.Clear
    lstItems.Clear
    lstquantity.Clear
    cboTables.ListIndex = -1
    cboTimeFrom.ListIndex = -1
    cbocategory.ListIndex = -1
    cboItems.ListIndex = -1
    txtRent.Text = ""
    txtHours.Text = ""
    txtPrice.Text = ""
    txtAmount.Text = ""
    txtDiscount.Text = ""
    txtNetPayment.Text = ""
End Function
Private Sub txtDiscount_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtNetPayment = Val(txtAmount.Text) - Val(txtDiscount.Text)
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
     If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
        MsgBox "Please Enter Numerc Value"
        Exit Sub
    End If
End Sub
Private Sub generateOID()
    Dim r As ADODB.Recordset
    Set r = New ADODB.Recordset
r.Open "select * from morder", cn, adOpenKeyset, adLockOptimistic
If r.EOF = True Then
txtOrderNo.Text = 1
Exit Sub
Else
r.MoveLast
txtOrderNo.Text = r.Fields("O_No").Value + 1
End If
End Sub

Private Sub txtqty_LostFocus()
txtAmount.Text = ""
If txtqty.Text = "" Then
txtAmount.Text = Val(txtPrice.Text) * 1
Else
txtAmount.Text = Val(txtPrice.Text) * Val(txtqty.Text) + Val(txtRent.Text)
End If
End Sub
