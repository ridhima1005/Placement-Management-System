VERSION 5.00
Begin VB.Form Studentm 
   Caption         =   "Student Menu"
   ClientHeight    =   6960
   ClientLeft      =   -390
   ClientTop       =   180
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11760
   Begin VB.CommandButton cmd6 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "STUDENT INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "Studentm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'edit
Private Sub cmd3_Click()
EditStudent.Show
Unload Studentm
End Sub

'records
Private Sub cmd4_Click()
StudentDetails.Show
End Sub

'home
Private Sub cmd6_Click()
Menu.Show
Unload Studentm
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Student"

End Sub

'new
Private Sub cmd1_Click()
NewStudent.Show
Unload Studentm
End Sub

'search
Private Sub cmd2_Click()
ViewStudent.Show
Unload Studentm
End Sub

'delete
Private Sub cmd5_Click()
Dim view As String

If (MsgBox("Are you sure to delete...", vbYesNo) = vbYes) Then
On Error GoTo l1
view = InputBox("enter the full name")
Set rs = New ADODB.Recordset
On Error GoTo l1

rs.Open "select * from Student where name = '" & view & "'", con, adOpenKeyset, adLockPessimistic

rs.Delete
con.Execute "commit"
rs.Close
Set rs = Nothing
MsgBox "Deleted Succesfully..."
Exit Sub
l1:
MsgBox "Rec not found"
End If
End Sub
