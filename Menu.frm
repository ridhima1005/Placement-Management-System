VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Menu"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   2040
      Picture         =   "Menu.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   6555
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "LOGOUT"
      Height          =   615
      Left            =   12840
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "ELIGIBLE STUDENTS"
      Height          =   975
      Left            =   16440
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "COMPANY INFORMATION"
      Height          =   855
      Left            =   12960
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "STUDENT INFORMATION"
      Height          =   855
      Left            =   9600
      TabIndex        =   0
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "PLACEMENT DATA INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   18855
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'student information command button
Private Sub cmd1_Click()
Studentm.Show
Menu.Hide
Unload Menu
End Sub

'company information command button
Private Sub cmd2_Click()
Companym.Show
Menu.Hide
Unload Menu
End Sub

'search student
Private Sub cmd3_Click()
Searchs.Show
Menu.Hide
Unload Menu
End Sub

'ELIGIBLE STUDENTS
Private Sub cmd4_Click()
CompanySelect.Show
Menu.Hide
Unload Menu
End Sub

'logout command button
Private Sub cmd5_Click()
End
End Sub
