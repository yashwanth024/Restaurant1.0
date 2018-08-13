VERSION 5.00
Begin VB.Form frmHDBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill For Home Delivery"
   ClientHeight    =   5700
   ClientLeft      =   6810
   ClientTop       =   4650
   ClientWidth     =   7755
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
   Picture         =   "frmGenerateBill.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   7755
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
      Height          =   4305
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   6585
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   405
         Left            =   600
         TabIndex        =   4
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdGenerateBill 
         Caption         =   "&Generate Bill"
         Height          =   405
         Left            =   2040
         TabIndex        =   3
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   4035
         Left            =   4440
         TabIndex        =   8
         Top             =   120
         Width           =   1995
         Begin VB.TextBox txtNetpayment 
            Height          =   405
            Left            =   150
            TabIndex        =   16
            Text            =   "0"
            Top             =   3480
            Width           =   1515
         End
         Begin VB.TextBox txtDiscount 
            Height          =   405
            Left            =   120
            TabIndex        =   6
            Text            =   "0"
            Top             =   1560
            Width           =   1515
         End
         Begin VB.TextBox txtAdvanced 
            Height          =   405
            Left            =   150
            TabIndex        =   7
            Text            =   "0"
            Top             =   2520
            Width           =   1515
         End
         Begin VB.TextBox txtAmount 
            Height          =   405
            Left            =   120
            TabIndex        =   5
            Text            =   "0"
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label lblNetpayment 
            BackStyle       =   0  'Transparent
            Caption         =   " Net Payment"
            Height          =   345
            Left            =   120
            TabIndex        =   17
            Top             =   3000
            Width           =   1725
         End
         Begin VB.Label lblDiscount 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            Height          =   345
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label lblAdvance 
            BackStyle       =   0  'Transparent
            Caption         =   "Advanced"
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1425
         End
         Begin VB.Label lblAmount 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   345
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.ListBox lstCustomer 
         Height          =   2490
         ItemData        =   "frmGenerateBill.frx":A809
         Left            =   1575
         List            =   "frmGenerateBill.frx":A80B
         TabIndex        =   1
         Top             =   1170
         Width           =   1245
      End
      Begin VB.ListBox lstOrderDate 
         Height          =   2490
         ItemData        =   "frmGenerateBill.frx":A80D
         Left            =   2970
         List            =   "frmGenerateBill.frx":A80F
         TabIndex        =   2
         Top             =   1155
         Width           =   1245
      End
      Begin VB.ListBox lstOrderNo 
         Height          =   2490
         ItemData        =   "frmGenerateBill.frx":A811
         Left            =   180
         List            =   "frmGenerateBill.frx":A813
         TabIndex        =   0
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         Height          =   375
         Left            =   2940
         TabIndex        =   11
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order NO"
         Height          =   375
         Left            =   60
         TabIndex        =   10
         Top             =   540
         Width           =   1635
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Generation"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmHDBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerateBill_Click()
    If lstOrderNo.ListIndex < 0 Then
        MsgBox "Please Select Any Order No. "
        Exit Sub
    End If
    If DataEnvironment1.rscmdHDBill_Grouping.State Then
        DataEnvironment1.rscmdHDBill_Grouping.Close
    End If
    DataEnvironment1.cmdHDBill_Grouping CInt(lstOrderNo.List(lstOrderNo.ListIndex))
    DRHDBill.Show
End Sub

Private Sub Form_Activate()
    CenterWindow Me
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    rec.Open "select * from mOrder where o_type='" & "No" & "'", cn, adOpenKeyset, adLockOptimistic
     loadOrders
End Sub

Private Function loadOrders()
    lstOrderNo.Clear
    lstOrderDate.Clear
    lstCustomer.Clear
    If rec.RecordCount > 0 Then
        While Not rec.EOF
            lstOrderNo.AddItem (rec.Fields(0).Value)
            lstOrderDate.AddItem (rec.Fields(1).Value)
            lstCustomer.AddItem (rec.Fields(2).Value)
            rec.MoveNext
        Wend
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub

Private Sub lstCustomer_Click()
    If (lstCustomer.ListIndex) >= 0 Then
        lstOrderDate.ListIndex = (lstCustomer.ListIndex)
        lstOrderNo.ListIndex = (lstCustomer.ListIndex)
        rec.MoveFirst
        rec.Find "O_No=" & Val(lstOrderNo.List(lstOrderNo.ListIndex))
        If Not rec.EOF Then
            If rec.Fields("O_No") = Val(lstOrderNo.List(lstOrderNo.ListIndex)) Then
                txtAmount.Text = rec.Fields("Amount").Value
                txtDiscount.Text = rec.Fields("Discount").Value
                txtAdvanced.Text = rec.Fields("Adv_payment").Value
                txtNetpayment.Text = Val(txtAmount.Text) - (Val(txtDiscount.Text) + Val(txtAdvanced.Text))
            End If
        End If
    End If
End Sub

Private Sub lstOrderDate_Click()
    If (lstOrderDate.ListIndex) >= 0 Then
        lstOrderNo.ListIndex = (lstOrderDate.ListIndex)
        lstCustomer.ListIndex = (lstOrderDate.ListIndex)
        rec.Find "o_No=" & lstOrderNo.List(lstOrderNo.ListIndex)
        rec.MoveFirst
        If Not rec.EOF Then
            If rec.Fields("O_No") = lstOrderNo.List(lstOrderNo.ListIndex) Then
                txtAmount.Text = rec.Fields("Amount").Value
                txtDiscount.Text = rec.Fields("Discount").Value
                txtAdvanced.Text = rec.Fields("Adv_payment").Value
                txtNetpayment.Text = Val(txtAmount.Text) - (Val(txtDiscount.Text) + Val(txtAdvanced.Text))
            End If
        End If
    End If
End Sub

Private Sub lstOrderNo_Click()
    If (lstOrderNo.ListIndex) >= 0 Then
        lstOrderDate.ListIndex = (lstOrderNo.ListIndex)
        lstCustomer.ListIndex = (lstOrderNo.ListIndex)
        rec.Find "o_No=" & lstOrderNo.List(lstOrderNo.ListIndex)
        rec.MoveFirst
        If Not rec.EOF Then
            If rec.Fields("O_No") = lstOrderNo.List(lstOrderNo.ListIndex) Then
                txtAmount.Text = rec.Fields("Amount").Value
                txtDiscount.Text = rec.Fields("Discount").Value
                txtAdvanced.Text = rec.Fields("Adv_payment").Value
                txtNetpayment.Text = Val(txtAmount.Text) - (Val(txtDiscount.Text) + Val(txtAdvanced.Text))
            End If
        End If
    End If
End Sub

