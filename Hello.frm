VERSION 5.00
Begin VB.MDIForm Welcome 
   BackColor       =   &H80000007&
   Caption         =   "Welcome"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   4095
      Left            =   0
      Picture         =   "Hello.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   0
      Width           =   11760
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

End Sub
