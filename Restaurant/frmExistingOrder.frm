VERSION 5.00
Begin VB.Form frmExistingOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Existing Order"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
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
   Picture         =   "frmExistingOrder.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   6945
   Begin VB.Frame fraExistingOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Existing Order"
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ListBox lstOrderDate 
         Height          =   2220
         ItemData        =   "frmExistingOrder.frx":26A02
         Left            =   2430
         List            =   "frmExistingOrder.frx":26A04
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.ListBox lstCustomer 
         Height          =   2220
         ItemData        =   "frmExistingOrder.frx":26A06
         Left            =   4590
         List            =   "frmExistingOrder.frx":26A08
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.ListBox lstOrderNo 
         Height          =   2220
         ItemData        =   "frmExistingOrder.frx":26A0A
         Left            =   240
         List            =   "frmExistingOrder.frx":26A0C
         TabIndex        =   0
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblOrderDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblCustomer 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblOrderNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Orders Detail"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmExistingOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub cmdLoad_Click()
    If (lstOrderNo.ListIndex) > 0 Then
        loadOno = Val(lstOrderNo.List(lstOrderNo.ListIndex))
    Else
        MsgBox "Select Any Order For Loading"
        Exit Sub
    End If
    LoadTime = True
    Unload Me
    frmOrder.Show
End Sub

Private Sub Form_Activate()
    CenterWindow Me
    loadOrders
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    rec.Open "select * from mOrder", cn, adOpenKeyset, adLockOptimistic
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
    End If
End Sub

Private Sub lstCustomer_DblClick()
    cmdLoad_Click
End Sub

Private Sub lstOrderDate_Click()
    If (lstOrderDate.ListIndex) >= 0 Then
        lstOrderNo.ListIndex = (lstOrderDate.ListIndex)
        lstCustomer.ListIndex = (lstOrderDate.ListIndex)
    End If
End Sub

Private Sub lstOrderDate_DblClick()
    cmdLoad_Click
End Sub

Private Sub lstOrderNo_Click()
    If (lstOrderNo.ListIndex) >= 0 Then
        lstOrderDate.ListIndex = (lstOrderNo.ListIndex)
        lstCustomer.ListIndex = (lstOrderNo.ListIndex)
    End If
End Sub

Private Sub lstOrderNo_DblClick()
    cmdLoad_Click
End Sub
