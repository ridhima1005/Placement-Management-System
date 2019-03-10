VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CompanySelect 
   Caption         =   "CompanySelect"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "CompanySelect.frx":0000
      Left            =   6240
      List            =   "CompanySelect.frx":001F
      TabIndex        =   1
      Text            =   "Student Criteria"
      Top             =   720
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   3480
      TabIndex        =   0
      Top             =   2280
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Students"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CompanySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

con.ConnectionString = "Provider=MSDASQL.1;Password=ridhima;Persist Security Info=True;User ID=system;Data Source=Student"

con.CursorLocation = adUseClient
con.Open
cmd.ActiveConnection = con
cmd.CommandType = adCmdText
'view = InputBox("Enter the company name")
'On Error GoTo l1
'rs.Open "select id,name,website,address,contact,aggregrate,dead_backlog,live_backlog from Company where id='" & view & "'", con, adOpenDynamic, adLockOptimistic
'On Error GoTo l1

'cmd.CommandText = "select name from Student where "
 If Combo1.ListIndex = 0 Then
        cmd.CommandText = "select * from Student where aggregate>='66' And dead_backlog='0' And live_backlog='0'"
        End If
         
         If Combo1.ListIndex = 1 Then
        cmd.CommandText = "select * from Student where aggregate>='66' And dead_backlog>='1' Or live_backlog>='1'"
        End If
        
        If Combo1.ListIndex = 2 Then
        cmd.CommandText = "select * from Student where aggregate>='60' And dead_backlog='0' And live_backlog='0'"
        End If
       
If Combo1.ListIndex = 3 Then
        cmd.CommandText = "select * from Student where aggregate>='60' And dead_backlog>='1' Or live_backlog>='1'"
        End If
        
        If Combo1.ListIndex = 4 Then
        cmd.CommandText = "select * from Student where aggregate>='55' And dead_backlog='0' And live_backlog='0'"
        End If
        
        If Combo1.ListIndex = 5 Then
        cmd.CommandText = "select * from Student where aggregate>='55' And dead_backlog>='1' And live_backlog>='1'"
        End If
        
        If Combo1.ListIndex = 6 Then
        cmd.CommandText = "select * from Student where aggregate>='50' And dead_backlog='0' And live_backlog='0'"
        End If
        
        If Combo1.ListIndex = 7 Then
        cmd.CommandText = "select * from Student where aggregate>='50' And dead_backlog>='1' And live_backlog>='1'"
        End If
        
        If Combo1.ListIndex = 8 Then
        cmd.CommandText = "select * from Student where aggregate<'50'"
        End If
        
Set rs = cmd.Execute
Set DataGrid1.DataSource = rs
End Sub

Private Sub Form_Load()



End Sub
