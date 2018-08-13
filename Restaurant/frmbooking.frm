VERSION 5.00
Begin VB.Form frmBooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking Information"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
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
   ScaleHeight     =   6885
   ScaleWidth      =   8115
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   5775
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Booking No"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Table No"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   6000
      TabIndex        =   6
      Top             =   600
      Width           =   2055
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&AddNew"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   1575
      End
   End
   Begin VB.Frame fraNavigate 
      Caption         =   "Navigation"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   5775
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   495
         Index           =   2
         Left            =   3120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   495
         Index           =   3
         Left            =   4560
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Booking Information"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
