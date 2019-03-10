VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form NewStudent 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Student"
   ClientHeight    =   10230
   ClientLeft      =   -390
   ClientTop       =   720
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   Begin VB.TextBox txt19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16320
      TabIndex        =   46
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txt18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   44
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox txt17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   43
      Top             =   6840
      Width           =   3735
   End
   Begin VB.TextBox txt13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15720
      TabIndex        =   41
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txt12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11520
      TabIndex        =   39
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txt11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      TabIndex        =   37
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txt10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   36
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   32
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   31
      Top             =   2040
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   13800
      TabIndex        =   30
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   133365761
      CurrentDate     =   41905
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   9240
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
      Connect         =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Student"
      OLEDBString     =   "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Student"
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
   Begin VB.TextBox txt14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   29
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txt16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11640
      TabIndex        =   28
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txt15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txt9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txt8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11520
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txt7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txt6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   22
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txt4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6840
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7680
      TabIndex        =   19
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   18
      Top             =   1080
      Width           =   3615
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
      Left            =   14760
      TabIndex        =   17
      Top             =   9000
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
      Left            =   8880
      TabIndex        =   16
      Top             =   9120
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
      Left            =   3480
      TabIndex        =   2
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000B&
      Caption         =   "Native Place"
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
      Left            =   13800
      TabIndex        =   45
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email-id"
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
      Left            =   240
      TabIndex        =   42
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 8th sem%"
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
      Left            =   13440
      TabIndex        =   40
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 7th sem%"
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
      Left            =   9000
      TabIndex        =   38
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 6th sem%"
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
      Left            =   4320
      TabIndex        =   35
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 5th sem%"
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
      Left            =   0
      TabIndex        =   34
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Gender"
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
      TabIndex        =   33
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lbl16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
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
      Left            =   5880
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lbl15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aggregate %"
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
      TabIndex        =   14
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lbl14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label lbl13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   9120
      TabIndex        =   12
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 4th sem%"
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
      Left            =   13440
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lbl10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 2nd sem%"
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
      Left            =   4320
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lbl11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 3rd sem%"
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
      Left            =   9240
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "B.E 1st sem%"
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
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Diploma %"
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
      Left            =   10920
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "12 %"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "10 %"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      TabIndex        =   4
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOB"
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
      Left            =   11760
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRNo"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   18855
   End
End
Attribute VB_Name = "NewStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

'clear
Private Sub cmd2_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
txt6.Text = ""
txt7.Text = ""
txt8.Text = ""
txt9.Text = ""
txt10.Text = ""
txt11.Text = ""
txt12.Text = ""
txt13.Text = ""
txt14.Text = ""
txt15.Text = ""
txt16.Text = ""
txt17.Text = ""
txt18.Text = ""
txt19.Text = ""
End Sub

'back
Private Sub cmd3_Click()
Menus.Show
Unload NewStudent
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command

con.Open "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Student"
Adodc1.Visible = False
'con.CursorLocation = adUseClient
'cmd.ActiveConnection = con
'cmd.CommandType = adCmdText
'con.Close

'Exit Sub
End Sub

'submit
Private Sub cmd1_Click()
Dim id As Integer
Dim gen As String
If (Option1.Value = True) Then
gen = "male"
ElseIf (Option2.Value = True) Then
gen = "female"
End If

'if any field is empty
If txt1.Text = "" Or txt2.Text = "" Or txt3.Text Or txt4.Text = "" Or txt5.Text = "" Or txt6.Text = "" Or txt7.Text = "" Or txt8.Text = "" Or txt9.Text = "" Or txt10.Text = "" Or txt11.Text = "" Or txt12.Text = "" Or txt13.Text = "" Or txt14.Text = "" Or txt15.Text = "" Or txt16.Text = "" Or txt17.Text = "" Or txt18.Text = "" Or txt19.Text = "" Then
MsgBox "No field should be left blank!"

Else
'con.CursorLocation = adUseClient
'con.Open
'cmd.ActiveConnection = con
'cmd.CommandType = adCmdText

'autoincrement of id
'cmd.CommandText = "select max(d_id) from Student"
'On Error GoTo l1
'Set rs = cmd.Execute
'id = rs.Fields(0)
'txt3.Text = id + 1

'insert values into table
'cmd.CommandText = "insert into Student values('" & txt1.Text & "','" & txt3.Text & "','" & txt2.Text & "','" & txt4.Text & "','" & txt5.Text & "','" & txt6.Text & "','" & txt7.Text & "','" & txt8.Text & "','" & txt9.Text & "','" & txt10.Text & "','" & txt11.Text & "','" & txt12.Text & "','" & txt13.Text & "','" & txt14.Text & "' )"
rs.Open "insert into Student values('" & txt1.Text & "','" & txt2.Text & "','" & DTPicker1.Value & "','" & gen & "',''" & txt3.Text & "','" & txt4.Text & "','" & txt5.Text & "','" & txt6.Text & "','" & txt7.Text & "','" & txt8.Text & "','" & txt9.Text & "','" & txt10.Text & "','" & txt11.Text & "','" & txt12.Text & "','" & txt13.Text & "','" & txt14.Text & "','" & txt15.Text & "','" & txt16.Text & "','" & txt17.Text & "','" & txt18.Text & "','" & txt19.Text & "' )", con, adOpenDynamic, adLockOptimistic
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

'prno
Private Sub txt1_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 127 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 58)) Then
MsgBox "Only alphabets and numbers are allowed..."
KeyAscii = 0
txt1.Text = ""
End If
End Sub

'7th
Private Sub txt12_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt12.Text = ""
End If
End Sub

'aggregate
Private Sub txt14_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt14.Text = ""
End If
End Sub

'contact
Private Sub txt18_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt18.Text = ""
End If
End Sub

'name
Private Sub txt2_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 127 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8) Then
MsgBox "Only alphabets are allowed..."
KeyAscii = 0
txt2.Text = ""
End If
End Sub

'5th
Private Sub txt10_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt10.Text = ""
End If
End Sub

'6th
Private Sub txt11_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt11.Text = ""
End If
End Sub

'livebacklog
Private Sub txt16_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Only numbers are allowed..."
txt16.Text = ""
End If
End Sub

'dead backlog
Private Sub txt15_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Only numbers are allowed..."
txt15.Text = ""
End If
End Sub

'8th
Private Sub txt13_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt13.Text = ""
End If
End Sub

'email
Private Sub txt17_lostfocus()
Dim str1 As String

str1 = txt17.Text
If ((InStr(str1, "@")) And ((InStr(str1, "gmail.com")) Or (InStr(str1, "hotmail.com")) Or (InStr(str1, "google.com")) Or (InStr(str1, "yahoo.com")) Or (InStr(str1, "yahoo.in")))) Then
MsgBox "Valid email id"
Else
MsgBox "Not valid email id.. Plz enter a valid email id"
txt17.Text = ""
End If
End Sub

'10
Private Sub txt3_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Only numbers are allowed..."
txt3.Text = ""
End If
End Sub

'12
Private Sub txt4_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8) Then
MsgBox "Enter proper contact..."
txt4.Text = ""
End If
End Sub

'contact length
Private Sub txt18_LostFocus()
If Len(txt4.Text) <> 10 Then
MsgBox "Contact number should be 10 digits"
txt18.Text = ""
End If
End Sub

'1st
Private Sub txt6_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt6.Text = ""
End If
End Sub

'2nd
Private Sub txt7_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt7.Text = ""
End If
End Sub

'3rd
Private Sub txt8_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt8.Text = ""
End If
End Sub

'4th
Private Sub txt9_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46) Then
MsgBox "Enter neatly..."
txt9.Text = ""
End If
End Sub
