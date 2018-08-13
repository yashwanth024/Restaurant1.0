VERSION 5.00
Begin VB.Form frmRBooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report For Bookings"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
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
   Picture         =   "frmRBooking.frx":0000
   ScaleHeight     =   2535
   ScaleWidth      =   4950
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   375
      Left            =   2310
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cboDates 
      Height          =   390
      ItemData        =   "frmRBooking.frx":1B118
      Left            =   2400
      List            =   "frmRBooking.frx":1B11A
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblAdvanceBooking 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Booking"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Dates"
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmRBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
     If cboDates.ListIndex < 0 Then
        MsgBox "Please Select Any Order No. "
        Exit Sub
    End If
    If DataEnvironment1.rscmdAdvBooking_Grouping.State Then
        DataEnvironment1.rscmdAdvBooking_Grouping.Close
    End If
    
    DataEnvironment1.cmdAdvBooking_Grouping CDate(cboDates.List(cboDates.ListIndex))
    DRAdvBooking.Show
End Sub

Private Sub Form_Activate()
    CenterWindow Me
    loadOrders
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    rec.Open "select distinct Booking_Date from Advance_Booking", cn, adOpenKeyset, adLockOptimistic
End Sub

Private Function loadOrders()
    cboDates.Clear
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        While Not rec.EOF
            cboDates.AddItem (rec.Fields("Booking_Date").Value)
            rec.MoveNext
        Wend
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    rec.Close
    Set rec = Nothing
End Sub



