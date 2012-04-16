VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Karname2 
   Caption         =   "ﬂ«—‰«„Â ê—ÊÂÌ"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   990
      Left            =   5925
      Picture         =   "Karname2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   465
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   990
      Left            =   5010
      Picture         =   "Karname2.frx":05DE
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   465
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   990
      Left            =   6840
      Picture         =   "Karname2.frx":0A45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   465
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   3015
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1095
      Width           =   1530
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Karname2.frx":0FEC
      Height          =   4515
      Left            =   3765
      TabIndex        =   12
      Top             =   1740
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   7964
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
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
         DataField       =   "family"
         Caption         =   "family"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   675.213
         EndProperty
      EndProperty
   End
   Begin VB.ListBox Combo1 
      Height          =   1425
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   195
      TabIndex        =   18
      Top             =   3855
      Width           =   3495
      Begin MSMask.MaskEdBox text4 
         Height          =   300
         Left            =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox Check1 
         Caption         =   "«„—Ê“"
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin MSMask.MaskEdBox text5 
         Height          =   300
         Left            =   585
         TabIndex        =   8
         Top             =   1140
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   " «—ÌŒ  ÕÊÌ· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2235
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6240
         TabIndex        =   20
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label9 
         Caption         =   " «—ÌŒ «„ Õ«‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1140
         Width           =   855
      End
   End
   Begin VB.Data Data_asli 
      Caption         =   "Data_asli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   5145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "students"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data_fari 
      Caption         =   "Data_fari"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   5070
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "karname"
      Top             =   3465
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data_groups 
      Caption         =   "data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5295
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "groups"
      Top             =   4395
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text22 
      Height          =   345
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   450
      Width           =   1530
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5100
      Top             =   3090
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
      CommandType     =   1
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
      RecordSource    =   "SELECT * FROM karname"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "Report5.rpt"
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport_kodakan 
      Left            =   240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "Report6.rpt"
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   195
      TabIndex        =   22
      Top             =   1650
      Width           =   3495
      Begin VB.TextBox Text3 
         DataField       =   "a_lis"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "b_rid"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "c_rit"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "d_spk"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "e_pst"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   5
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "‰„—« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   28
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label6 
         Caption         =   "A) Listening Comprehension :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "B) Reading Comprehension :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "E) Class Activity :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1785
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "D) Speaking Skill :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "C) Structure &&  Written expression :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   210
      TabIndex        =   29
      Top             =   1650
      Width           =   3495
      Begin VB.TextBox Text3 
         DataField       =   "a_lis"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   13
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "b_rid"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   8
         Left            =   2640
         TabIndex        =   14
         Top             =   690
         Width           =   735
      End
      Begin VB.TextBox Text3 
         DataField       =   "c_rit"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   9
         Left            =   2640
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "C) Listening Skill :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   1095
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "B) Paper Test grade :"
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "A) Class Activity :"
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   31
         Top             =   345
         Width           =   2910
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "‰„—« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   30
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Level :"
      Height          =   195
      Index           =   1
      Left            =   3525
      TabIndex        =   35
      Top             =   870
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Teacher :"
      Height          =   195
      Index           =   0
      Left            =   3465
      TabIndex        =   17
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "Karname2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Group_Type
Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then text4.Text = Replace(CStr(Date), "-", "/")
If Check1.Value = 0 Then text4.Text = ""
End Sub

Private Sub Combo1_Click()
List_them
Select Case Group_Type
Case "kids"
Frame4.Visible = True
Frame2.Visible = False
Case "adults"
Frame2.Visible = True
Frame4.Visible = False
Case Else
Frame2.Visible = True
Frame4.Visible = False
End Select
End Sub

Private Sub Command1_Click()
Dim p As Object
Select Case Group_Type
Case "kids"
Set p = CrystalReport_kodakan
Case "adults"
Set p = CrystalReport1
Case Else
 MsgBox "Select a group first"
 Exit Sub
End Select
Adodc1.Recordset.Update
p.Destination = crptToFile
'p.PrintFileName = "c:\" + "Karname Group(" + Data_groups.Recordset.Fields("name") + ") " + Group_Type + ".rtf"
For i = 1 To 10
p.RetrieveDataFiles
Next
p.PrintReport
End Sub

Private Sub Command2_Click()
Dim p As Object
Select Case Group_Type
Case "kids"
Set p = CrystalReport_kodakan
Case "adults"
Set p = CrystalReport1
Case Else
 MsgBox "Select a group first"
 Exit Sub
End Select
Adodc1.Recordset.Update
p.Destination = crptToWindow
For i = 1 To 10
p.RetrieveDataFiles
Next
p.PrintReport
End Sub

Private Sub Command3_Click()
Dim p As Object
Select Case Group_Type
Case "kids"
Set p = CrystalReport_kodakan
Case "adults"
Set p = CrystalReport1
Case Else
 MsgBox "Select a group first"
 Exit Sub
End Select
Adodc1.Recordset.Update
p.Destination = crptToPrinter
'p.PrintFileName = "c:\" + "Karname Group(" + Data_groups.Recordset.Fields("name") + ") " + Group_Type + ".rtf"
For i = 1 To 10
p.RetrieveDataFiles
Next
p.PrintReport
End Sub

Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''
CrystalReport1.ReportFileName = App.Path + "\report5.rpt"
CrystalReport_kodakan.ReportFileName = App.Path + "\report6.rpt"
Data_asli.DatabaseName = App.Path + "\db1.mdb"
Data_fari.DatabaseName = App.Path + "\db1.mdb"
Data_groups.DatabaseName = App.Path + "\db1.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db1.mdb;Persist Security Info=False"
'''''''''''''''''''''''''''''''''''''''''''
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
For i = 1 To 10
Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM karname"
Adodc1.Refresh
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 12 Then
Load Me
Exit Sub
End If
Me.Hide
mainshow
End Sub

Private Sub List_them()
Adodc1.RecordSource = "SELECT * FROM karname"
Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
Adodc1.Refresh
''''''''''''teachers name''''''''''''''''''''''''''''''''
Data_groups.Recordset.MoveFirst                         '
For i = 1 To Combo1.ListIndex                           '
Data_groups.Recordset.MoveNext                          '
Next                                                    '
Text22.Text = Data_groups.Recordset.Fields("teacher")   '
Text1.Text = Data_groups.Recordset.Fields("level")   '
'''''''''''''delete records'''''''''''''''''''''''''''''''
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
Data_fari.Recordset.AddNew
Data_fari.Recordset.Fields("name") = Data_asli.Recordset.Fields("name")
Data_fari.Recordset.Fields("family") = Data_asli.Recordset.Fields("family")
Data_fari.Recordset.Fields("tt") = Data_asli.Recordset.Fields("tt")
Data_fari.Recordset.Fields("shsh") = Data_asli.Recordset.Fields("shsh")
Data_fari.Recordset.Fields("code") = Data_asli.Recordset.Fields("code")
Data_fari.Recordset.Fields("nationality") = Data_asli.Recordset.Fields("nationality")
Data_fari.Recordset.Fields("group") = group_name
Data_fari.Recordset.Fields("ts") = Data_asli.Recordset.Fields("ts")
'Data_fari.Recordset.Fields("a_lis") = Data_asli.Recordset.Fields("a_lis")
'Data_fari.Recordset.Fields("b_rid") = Data_asli.Recordset.Fields("b_rid")
'Data_fari.Recordset.Fields("c_rit") = Data_asli.Recordset.Fields("c_rit")
'Data_fari.Recordset.Fields("d_spk") = Data_asli.Recordset.Fields("d_spk")
'Data_fari.Recordset.Fields("e_pst") = Data_asli.Recordset.Fields("e_pst")
Data_fari.Recordset.Fields("date_test") = text5.Text
Data_fari.Recordset.Fields("date_get") = text4.Text
'Data_fari.Recordset.Fields("teacher") = Text22.Text
Data_fari.Recordset.Fields("term") = Data_asli.Recordset.Fields("term")
Data_fari.Recordset.Fields("level") = Data_groups.Recordset.Fields("level")
Data_fari.Recordset.Fields("phone") = Data_asli.Recordset.Fields("phone")
Data_fari.Recordset.Fields("gender") = Data_asli.Recordset.Fields("gender")
Data_fari.Recordset.Fields("film") = Data_groups.Recordset.Fields("film")
Data_fari.Recordset.Fields("book") = Data_groups.Recordset.Fields("book")
Data_asli.Recordset.FindNext ("group='" + Combo1.List(Combo1.ListIndex) + "'")
Data_fari.Recordset.Update
Wend

Adodc1.RecordSource = "SELECT * FROM karname"
Adodc1.Refresh
MsgBox "Students listed!" + vbclrf + "Press OK", , "Progress info"
Adodc1.RecordSource = "SELECT * FROM karname"
Adodc1.Refresh
'load_un Me
'Dim o As New Karname2
'Load o
'Unload Me
'o.Show
'Form_Unload (12)
Group_Type = Data_groups.Recordset.Fields("type")
End Sub

Private Sub Text3_GotFocus(Index As Integer)
On Error Resume Next
Adodc1.Recordset.Update
End Sub

Private Sub Text3_LostFocus(Index As Integer)
On Error Resume Next
Adodc1.Recordset.Update
End Sub
