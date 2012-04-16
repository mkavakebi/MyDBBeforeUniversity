VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form karname_form 
   Caption         =   " ‰ŸÌ„ ò«—‰«„Â"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
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
   ScaleHeight     =   4590
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   4905
      Picture         =   "karname_form.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3345
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   2895
      Picture         =   "karname_form.frx":0826
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3345
      Width           =   1980
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   885
      Picture         =   "karname_form.frx":1041
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3345
      Width           =   1980
   End
   Begin VB.Data Data_fari 
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
      Left            =   1710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "karname"
      Top             =   3270
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
      Left            =   5175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "groups"
      Top             =   3270
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "students"
      Top             =   3270
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   465
      TabIndex        =   36
      Top             =   2265
      Width           =   6855
      Begin MSMask.MaskEdBox text4 
         Height          =   345
         Left            =   570
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox Check1 
         Caption         =   "«„—Ê“"
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
         Left            =   600
         TabIndex        =   7
         Top             =   225
         Width           =   735
      End
      Begin MSMask.MaskEdBox text5 
         Height          =   345
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   0
         Width           =   390
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
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4065
      TabIndex        =   27
      Top             =   30
      Width           =   3255
      Begin VB.CheckBox Check2 
         Caption         =   "‰„«Ì‘ òœÂ«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
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
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "œ«‰‘ ¬„Ê“"
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
         Left            =   2280
         TabIndex        =   35
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   ": òœ œ«‰‘ ¬„Ê“"
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
         Left            =   2160
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Or"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   ": ‰«„ œ«‰‘ ¬„Ê“"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   465
      TabIndex        =   12
      Top             =   30
      Width           =   3495
      Begin VB.TextBox Text3 
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
         Index           =   4
         Left            =   2640
         TabIndex        =   5
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "C) Structure &&  Written expression :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "D) Speaking Skill :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "E) Class Activity :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "B) Reading Comprehension :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "A) Listening Comprehension :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2175
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
         TabIndex        =   13
         Top             =   0
         Width           =   450
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1050
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "Report5.rpt"
      Destination     =   2
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport_kodakan 
      Left            =   180
      Top             =   3375
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
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   465
      TabIndex        =   19
      Top             =   30
      Width           =   3495
      Begin VB.TextBox Text3 
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
         Index           =   9
         Left            =   2640
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   8
         Left            =   2640
         TabIndex        =   21
         Top             =   690
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Index           =   7
         Left            =   2640
         TabIndex        =   20
         Top             =   330
         Width           =   735
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
         TabIndex        =   26
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label6 
         Caption         =   "A) Class Activity :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   25
         Top             =   345
         Width           =   2910
      End
      Begin VB.Label Label6 
         Caption         =   "B) Paper Test grade :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "C) Listening Skill :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1095
         Width           =   2535
      End
   End
End
Attribute VB_Name = "karname_form"
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

Private Sub Check2_Click()
If Check2.Value = 1 Then List2.ZOrder (0)
If Check2.Value = 0 Then List1.ZOrder (0)
End Sub

Private Sub Command1_Click()
a = Train_Karname
If a = "error" Then Exit Sub
''''''''''''''''''''''
Dim p As Object
Select Case Group_Type
Case "kids"
Set p = CrystalReport_kodakan
Case "adualts"
Set p = CrystalReport1
Case Else
Exit Sub
End Select

For Each i In Text3
i.Text = ""
Next
p.Destination = crptToFile
'CrystalReport1.PrintFileName = "c:\" + Data_asli.Recordset.Fields("name") + "Karname " + Group_Type + ".rtf"
For i = 1 To 10
p.RetrieveDataFiles
Next
p.PrintReport
End Sub

Private Sub Command2_Click()
a = Train_Karname
If a = "error" Then Exit Sub
'''''''''''''''''''''
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
For Each i In Text3
i.Text = ""
Next
p.Destination = crptToWindow
For i = 1 To 10
p.RetrieveDataFiles
Next
p.PrintReport
End Sub

Private Sub Command3_Click()
a = Train_Karname
If a = "error" Then Exit Sub
''''''''''''''''''''''
Dim p As Object
Select Case Group_Type
Case "kids"
Set p = CrystalReport_kodakan
Case "adualts"
Set p = CrystalReport1
Case Else
Exit Sub
End Select

For Each i In Text3
i.Text = ""
Next
p.Destination = crptToPrinter
'CrystalReport1.PrintFileName = "c:\" + Data_asli.Recordset.Fields("name") + "Karname " + Group_Type + ".rtf"
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
mainshow
End Sub

Private Sub List1_Click()
Text1.Text = ""
Text1.Text = List2.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
Text1.Text = List2.List(List2.ListIndex)
End Sub

Private Sub Text1_Change()
code = Text1.Text
group_name = find_group(code)
Group_Type = find_group_type(group_name)
Select Case Group_Type
Case "kids"
Frame4.ZOrder 0
Case "adults"
Frame2.ZOrder 0
Case Else
Frame2.ZOrder 0
End Select
End Sub

Private Sub Text2_Change()
List1.Clear
List2.Clear
Data_asli.Recordset.MoveFirst
While Data_asli.Recordset.EOF = False
Nam = Data_asli.Recordset.Fields("name")
family = Data_asli.Recordset.Fields("family")
code = Data_asli.Recordset.Fields("code")
esm = family + " " + Nam
''''''''''''''''''''''''''''
If InStr(1, LCase(esm), Text2.Text) Then
List2.AddItem code
List1.AddItem esm
End If
Data_asli.Recordset.MoveNext
Wend
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
ind = List1.ListIndex
If ind > -1 Then
List1.RemoveItem (ind)
List2.RemoveItem (ind)
End If
End If
End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
ind = List2.ListIndex
If ind > -1 Then
List1.RemoveItem (ind)
List2.RemoveItem (ind)
End If
End If
End Sub

 
Function find_group(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_group = Data_asli.Recordset.Fields("group")
Else
find_group = "error"
End If
End Function

Function find_family(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_family = Data_asli.Recordset.Fields("family")
Else
find_family = "error"
End If
End Function

Function find_name(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_name = Data_asli.Recordset.Fields("name")
Else
find_name = "error"
End If
End Function

Function find_fname(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_fname = Data_asli.Recordset.Fields("fname")
Else
find_fname = "error"
End If
End Function

Function find_tt(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_tt = Data_asli.Recordset.Fields("tt")
Else
find_tt = "error"
End If
End Function

Function find_madrak(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_madrak = Data_asli.Recordset.Fields("madrak")
Else
find_madrak = "error"
End If
End Function

Function find_shsh(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_shsh = Data_asli.Recordset.Fields("shsh")
Else
find_shsh = "error"
End If
End Function

Function find_nat(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_nat = Data_asli.Recordset.Fields("nationality")
Else
find_nat = "error"
End If
End Function

Function find_phone(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_phone = Data_asli.Recordset.Fields("phone")
Else
find_phone = "error"
End If
End Function

Function find_gender(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_gender = Data_asli.Recordset.Fields("gender")
Else
find_gender = "error"
End If
End Function

Function find_info(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
f = Data_asli.Recordset.Fields("information")
If IsNull(f) Then f = ""
find_info = f
Else
find_info = "error"
End If
End Function

Function find_ts(ByVal code As String) As String
Data_asli.Recordset.FindFirst ("code='" + code + "'")
If Data_asli.Recordset.NoMatch = False Then
find_ts = Data_asli.Recordset.Fields("ts")
Else
find_ts = "error"
End If
End Function

Function find_level(ByVal group As String) As String
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
If Data_groups.Recordset.Fields("name") = group Then
find_level = Data_groups.Recordset.Fields("level")
GoTo k
End If
Data_groups.Recordset.MoveNext
Wend
find_level = "error"
k:
End Function

Function find_group_type(ByVal group As String) As String
On Error Resume Next
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
If Data_groups.Recordset.Fields("name") = group Then
find_group_type = Data_groups.Recordset.Fields("type")
GoTo k
End If
Data_groups.Recordset.MoveNext
Wend
find_group_type = "error"
k:
End Function

Function find_term(ByVal group As String) As String
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
If Data_groups.Recordset.Fields("name") = group Then
find_term = Data_groups.Recordset.Fields("term")
GoTo k
End If
Data_groups.Recordset.MoveNext
Wend
find_term = "error"
k:
End Function

Function find_book(ByVal group As String) As String
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
If Data_groups.Recordset.Fields("name") = group Then
find_book = Data_groups.Recordset.Fields("book")
GoTo k
End If
Data_groups.Recordset.MoveNext
Wend
find_book = "error"
k:
End Function

Function find_film(ByVal group As String) As String
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
If Data_groups.Recordset.Fields("name") = group Then
find_film = Data_groups.Recordset.Fields("film")
GoTo k
End If
Data_groups.Recordset.MoveNext
Wend
find_film = "error"
k:
End Function

Function Train_Karname() As String
'''''check the spaces''''''''''''
If Text1.Text = "" Then
MsgBox "òœ ò«—»—Ì Ê«—œ ‰‘œÂ «” .", vbInformation
Text1.SetFocus
Train_Karname = "error"
Exit Function
End If
'''''''''''''''''''''''''''''''''
code = Text1.Text
group_name = find_group(code)
'''''''''''''
If Group_Type <> "kids" Then
    For i = 0 To 4
    If Text3(i).Text = "" Then
    MsgBox "ÌòÌ «“ ‰„—«  Ê«—œ ‰‘œÂ «” ." + vbCrLf + "·ÿ›« ¬‰ —« Ê«—œ ò‰Ìœ.", vbInformation
    Text3(i).SetFocus
    Train_Karname = "error"
    Exit Function
    End If
    Next
Else
    For i = 7 To 9
    If Text3(i).Text = "" Then
    MsgBox "ÌòÌ «“ ‰„—«  Ê«—œ ‰‘œÂ «” ." + vbCrLf + "·ÿ›« ¬‰ —« Ê«—œ ò‰Ìœ.", vbInformation
    Text3(i).SetFocus
    Train_Karname = "error"
    Exit Function
    End If
    Next

End If
'''''''''''''''''''''''''''''''''
With Data_fari.Recordset
On Error Resume Next
.MoveFirst
While .EOF = False
.Delete
.MoveFirst
Wend
.AddNew
.Fields("name") = find_name(code)
.Fields("family") = find_family(code)
.Fields("tt") = find_tt(code)
.Fields("shsh") = find_shsh(code)
.Fields("code") = code
.Fields("nationality") = find_nat(code)
.Fields("group") = group_name
.Fields("ts") = find_ts(code)
.Fields("date_test") = text5.Text
.Fields("date_get") = text4.Text
.Fields("level") = find_level(group_name)
.Fields("term") = find_term(group_name)
.Fields("phone") = find_phone(code)
.Fields("gender") = find_gender(code)
.Fields("book") = find_book(group_name)
.Fields("film") = find_film(group_name)
Select Case Group_Type
Case "kids"
.Fields("a_lis") = Text3(7).Text
.Fields("b_rid") = Text3(8).Text
.Fields("c_rit") = Text3(9).Text
.Fields("d_spk") = ""
.Fields("e_pst") = ""
Case "adults"
.Fields("a_lis") = Text3(0).Text
.Fields("b_rid") = Text3(1).Text
.Fields("c_rit") = Text3(2).Text
.Fields("d_spk") = Text3(3).Text
.Fields("e_pst") = Text3(4).Text
Case Else
End Select
.Update
End With
End Function



