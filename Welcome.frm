VERSION 5.00
Begin VB.Form Welcome 
   Caption         =   "Welcome"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form2"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9120
      TabIndex        =   2
      Top             =   6600
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   7080
      Picture         =   "Welcome.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   2640
      Width           =   6375
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "WELCOME TO INFORMATION TECHNOLOGY DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18855
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Loginm.Show
Unload Welcome
End Sub

