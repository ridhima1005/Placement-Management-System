VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ViewCompany 
   Caption         =   "ViewCompany"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      Height          =   525
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox txt4 
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txt3 
      Height          =   495
      Left            =   14040
      TabIndex        =   7
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txt5 
      Height          =   495
      Left            =   14040
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txt6 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txt7 
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox txt8 
      Height          =   615
      Left            =   10800
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
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
      TabIndex        =   2
      Top             =   8640
      Width           =   1575
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
      TabIndex        =   1
      Top             =   8640
      Width           =   1455
   End
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
      TabIndex        =   0
      Top             =   8520
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
      Top             =   8520
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
      TabIndex        =   20
      Top             =   0
      Width           =   18975
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
      TabIndex        =   19
      Top             =   840
      Width           =   1695
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
      TabIndex        =   18
      Top             =   1920
      Width           =   1095
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
      TabIndex        =   17
      Top             =   1920
      Width           =   2175
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
      TabIndex        =   16
      Top             =   2880
      Width           =   2055
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
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   18975
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
      TabIndex        =   13
      Top             =   5160
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
      TabIndex        =   12
      Top             =   6840
      Width           =   2055
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
      TabIndex        =   11
      Top             =   6720
      Width           =   1815
   End
End
Attribute VB_Name = "ViewCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'search
Private Sub cmd1_Click()
Dim view As Integer
view = InputBox("Enter the id")
On Error GoTo l1
rs.Open "select id,name,website,address,contact,aggregate,dead_backlog,live_backlog from Company where id='" & view & "'", con, adOpenDynamic, adLockOptimistic
On Error GoTo l1
txt1.Text = rs.Fields("id")
txt2.Text = rs.Fields("name")
txt3.Text = rs.Fields("website")
txt4.Text = rs.Fields("address")
txt5.Text = rs.Fields("contact")
txt6.Text = rs.Fields("aggregate")
txt7.Text = rs.Fields("dead_backlog")
txt8.Text = rs.Fields("live_backlog")

If (rs.State = 1) Then
rs.Close
End If
MsgBox "Search successful"
Exit Sub
l1: MsgBox "Record not found"
End Sub

'back
Private Sub cmd3_Click()
Companym.Show
Unload ViewCompany
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Company"

End Sub

