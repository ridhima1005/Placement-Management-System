VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Logins 
   Caption         =   "Logins"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form2"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   10680
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txt2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   10680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   7800
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   11400
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   4800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Login"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Login"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "MODERN COLLEGE OF ENGINEERING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   18975
   End
   Begin VB.Label lbl2 
      Caption         =   "Username"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lbl3 
      Caption         =   "Password"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "Logins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'login command button
Private Sub cmd1_Click()
rs.Open "insert into login values('" & txt1.Text & "','" & txt2.Text & "')", con, adOpenDynamic, adLockOptimistic

If (rs.State = 1) Then
rs.Close
End If

If txt1.Text = "student" And txt2.Text = "student123" Then
 Menus.Show
 login.Hide
 Unload login
Else
MsgBox "Enter correct username or Password !"
txt1.Text = " "
txt2.Text = " "
End If
End Sub

'cancel command button
Private Sub cmd3_Click()
End
End Sub

'clear command button
Private Sub cmd2_Click()
txt1.Text = " "
txt2.Text = " "
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Login"

Adodc1.Visible = False
End Sub

