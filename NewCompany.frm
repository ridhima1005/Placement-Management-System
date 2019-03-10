VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form NewCompany 
   Caption         =   "NewCompany"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   20
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   19
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16680
      TabIndex        =   18
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox txt8 
      Height          =   615
      Left            =   10800
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txt7 
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txt6 
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txt5 
      Height          =   495
      Left            =   14040
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txt3 
      Height          =   495
      Left            =   14040
      TabIndex        =   13
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txt4 
      Height          =   735
      Left            =   5040
      TabIndex        =   12
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox txt2 
      Height          =   525
      Left            =   4920
      TabIndex        =   11
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
      Top             =   8760
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Company"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Company"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbl10 
      Caption         =   "Live Backlogs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lbl9 
      Caption         =   "Dead Backlogs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lbl8 
      Caption         =   "Aggregrate %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   18975
   End
   Begin VB.Label lbl6 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lbl7 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lbl4 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lbl3 
      Caption         =   "Website"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lbl2 
      Caption         =   "Company ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   18975
   End
End
Attribute VB_Name = "NewCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

Private Sub cmd1_Click()
Dim id As Integer

'if any field is empty
If txt1.Text = "" Or txt2.Text = "" Or txt4.Text = "" Or txt5.Text = "" Or txt6.Text = "" Or txt7.Text = "" Or txt8.Text = "" Then
MsgBox "No field should be left blank!"

Else
'con.CursorLocation = adUseClient
'con.Open
'cmd.ActiveConnection = con
'cmd.CommandType = adCmdText

'autoincrement of id
'cmd.CommandText = "select max(id) from Company"
'On Error GoTo l1
'Set rs = cmd.Execute
'id = rs.Fields(0)
'txt1.Text = id + 1

'insert values into table
'cmd.CommandText = "insert into Student values('" & txt1.Text & "','" & txt3.Text & "','" & txt2.Text & "','" & txt4.Text & "','" & txt5.Text & "','" & txt6.Text & "','" & txt7.Text & "','" & txt8.Text & "','" & txt9.Text & "','" & txt10.Text & "','" & txt11.Text & "','" & txt12.Text & "','" & txt13.Text & "','" & txt14.Text & "' )"
rs.Open "insert into Company values('" & txt1.Text & "','" & txt3.Text & "','" & txt2.Text & "','" & txt4.Text & "','" & txt5.Text & "','" & txt6.Text & "','" & txt7.Text & "','" & txt8.Text & "' )", con, adOpenDynamic, adLockOptimistic
If (rs.State = 1) Then
rs.Close
End If
'cmd.Execute
MsgBox "Successful!!!"
'con.Close
End If

'Exit Sub
'l1:
End Sub

'clear
Private Sub cmd2_Click()
txt1.Text = " "
txt2.Text = " "
txt3.Text = " "
txt4.Text = " "
txt5.Text = " "
txt6.Text = " "
txt7.Text = " "
txt8.Text = " "
End Sub

'back
Private Sub cmd3_Click()
Companym.Show
Unload NewCompany
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Company"
Adodc1.Visible = False
End Sub

'id validate
Private Sub txt1_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Only numbers are allowed..."
txt1.Text = ""
End If
End Sub

'code of valdiation for name
Private Sub txt2_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 127 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8) Then
MsgBox "Only alphabets are allowed..."
KeyAscii = 0
txt2.Text = ""
End If
End Sub

'code of valdiation for address
Private Sub txt4_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 127 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 127 Or (KeyAscii >= 48 And KeyAscii <= 57)) Then
MsgBox "enter correct address"
KeyAscii = 0
txt4.Text = ""
End If
End Sub

'contact
Private Sub txt5_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Enter proper contact..."
txt5.Text = ""
End If
End Sub

'code of valdiation for contact number length
Private Sub txt5_lostfocus()
If Len(txt5.Text) <> 10 Then
MsgBox "Contact number should be 10 digits"
txt5.Text = ""
End If
End Sub

'aggregrate
Private Sub txt6_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt6.Text = ""
End If
End Sub

'dead backlog
Private Sub txt7_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Enter number only..."
txt7.Text = ""
End If
End Sub

'live backlog
Private Sub txt8_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Enter number only..."
txt8.Text = ""
End If
End Sub
