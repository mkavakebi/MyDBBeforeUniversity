VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form studentspresentation_form 
   Caption         =   "ÍÖæÑ æ ÛíÇÈ ÏÇäÔ ÂãæÒÇä"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   8790
      Picture         =   "studentspay_form.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   435
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   6960
      Picture         =   "studentspay_form.frx":05A7
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   435
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   7875
      Picture         =   "studentspay_form.frx":0A0E
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   435
      Width           =   915
   End
   Begin VB.Data data_date 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "security_time_backups"
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1410
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Data Data_groups 
      Caption         =   "data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "groups"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data_fari 
      Caption         =   "Data_fari"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6390
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "presentations"
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Data Data_asli 
      Caption         =   "Data_asli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "students"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   285
      Top             =   855
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM presentations"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DG1 
      Align           =   2  'Align Bottom
      Bindings        =   "studentspay_form.frx":0FEC
      Height          =   5730
      Left            =   0
      TabIndex        =   6
      Top             =   1770
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   10107
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16744576
      ForeColor       =   8388608
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÝÑã ÍÖæÑ æ ÛíÇÈ"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "code"
         Caption         =   "code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "phone"
         Caption         =   "Phone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "teacher"
         Caption         =   "Teacher"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "group"
         Caption         =   "Group"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   734.74
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox text1 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "####/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox text1 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "####/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ": ÊÇÑíÎ ÇíÇä ÊÑã"
      Height          =   240
      Index           =   3
      Left            =   2520
      TabIndex        =   5
      Top             =   1095
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ": ÊÇÑíÎ ÔÑæÚ ÊÑã"
      Height          =   240
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   615
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ": Ñæå"
      Height          =   240
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ": ãÏÑÓ"
      Height          =   240
      Index           =   0
      Left            =   5760
      TabIndex        =   2
      Top             =   1095
      Width           =   525
   End
End
Attribute VB_Name = "studentspresentation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Adodc1.Refresh
''''''''''''teachers name''''''''''''''''''''''''''''''''
Data_groups.Recordset.MoveFirst                         '
For i = 1 To Combo1.ListIndex                           '
Data_groups.Recordset.MoveNext                          '
Next                                                    '
Text22.Text = Data_groups.Recordset.Fields("teacher")   '
'''''''''''''delet records'''''''''''''''''''''''''''''''
On Error GoTo 1
Data_fari.Recordset.MoveFirst
Y:
Data_fari.Recordset.Delete
Data_fari.Recordset.MoveFirst
If Data_fari.Recordset.EOF = False Then GoTo Y
1:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Data_asli.Recordset.FindFirst ("group='" + Combo1.List(Combo1.ListIndex) + "'")
Dim group_name As String
While Data_asli.Recordset.NoMatch = False
group_name = Data_asli.Recordset.Fields("group")
n_f = Data_asli.Recordset.Fields("name") & " " & Data_asli.Recordset.Fields("family")
Data_fari.Recordset.AddNew
Data_fari.Recordset.Fields("name") = n_f
Data_fari.Recordset.Fields("code") = Data_asli.Recordset.Fields("code")
Data_fari.Recordset.Fields("phone") = Data_asli.Recordset.Fields("phone")
Data_fari.Recordset.Fields("vs") = Data_asli.Recordset.Fields("vs")
Data_fari.Recordset.Fields("group") = group_name
Data_fari.Recordset.Fields("a_date") = Text1(0).Text
Data_fari.Recordset.Fields("b_date") = Text1(1).Text
Data_fari.Recordset.Fields("teacher") = Text22.Text
Data_asli.Recordset.FindNext ("group='" + Combo1.List(Combo1.ListIndex) + "'")
Data_fari.Recordset.Update
Wend
Adodc1.RecordSource = "SELECT * FROM presentations"
Adodc1.Refresh
MsgBox "Students listed!" + vbclrf + "Press OK", , "Progress info"
Adodc1.RecordSource = "SELECT * FROM presentations"
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
CrystalReport1.Destination = crptToFile
'CrystalReport1.PrintFileName = "c:\" + "Presentation Group(" + Data_groups.Recordset.Fields("name") + ")" + ".rtf"
For i = 1 To 10
CrystalReport1.RetrieveDataFiles
Next
CrystalReport1.PrintReport
End Sub


Private Sub Command2_Click()
CrystalReport1.Destination = crptToWindow
For i = 1 To 10
CrystalReport1.RetrieveDataFiles
Next
CrystalReport1.PrintReport
End Sub

Private Sub Command3_Click()
CrystalReport1.Destination = crptToPrinter
'CrystalReport1.PrintFileName = "c:\" + "Presentation Group(" + Data_groups.Recordset.Fields("name") + ")" + ".rtf"
For i = 1 To 10
CrystalReport1.RetrieveDataFiles
Next
CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''
CrystalReport1.ReportFileName = App.Path + "\report1.rpt"
Data_asli.DatabaseName = App.Path + "\db1.mdb"
Data_fari.DatabaseName = App.Path + "\db1.mdb"
Data_groups.DatabaseName = App.Path + "\db1.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db1.mdb;Persist Security Info=False"
'''''''''''''''''''''''''''''''''''''''''''
data_date.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Activate()
'''''''''combo1.validate{groups}
On Error Resume Next
Combo1.Clear
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
Combo1.AddItem (Data_groups.Recordset.Fields("name"))
Data_groups.Recordset.MoveNext
Wend
''''''''''''''''''
'''''''''''''''''
If data_date.Recordset.RecordCount > 0 Then
data_date.Recordset.MoveFirst
s = data_date.Recordset.Fields("cls_begin")
If IsNull(s) Then s = ""
Text1(0).Text = s
s = data_date.Recordset.Fields("cls_end")
If IsNull(s) Then s = ""
Text1(1).Text = s
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 12 Then
Load Me
Exit Sub
End If
Me.Hide
mainshow
End Sub
