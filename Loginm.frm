VERSION 5.00
Begin VB.Form Loginm 
   Caption         =   "Loginm"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   8895
      Begin VB.CommandButton cmd2 
         Caption         =   "Admin"
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Student"
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Loginm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Logins.Show
Unload Loginm
End Sub

Private Sub cmd2_Click()
login.Show
Unload Loginm
End Sub
