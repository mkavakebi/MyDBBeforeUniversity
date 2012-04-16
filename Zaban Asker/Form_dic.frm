VERSION 5.00
Begin VB.Form Form_dic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   1305
      Left            =   3330
      TabIndex        =   2
      Top             =   30
      Width           =   120
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "ZAsker.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1110
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Loghat(EtoP)"
      Top             =   585
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Data DataDIC 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "Dic.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Latin2Farsi"
      Top             =   75
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Data DataAli 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "ali.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ali"
      Top             =   30
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   495
      Top             =   150
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   3270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   3285
   End
End
Attribute VB_Name = "Form_dic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IsOpen As Boolean
Private Sub Command1_Click()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Activate()
Make_New_Dic
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Make_New_Dic
End Sub

Private Sub Make_New_Dic()
Left = Screen.Width - Width
Top = Screen.Height - Height - 500
ZOrder 0
Dim H As Currency
Dim G As Integer
Dim i As Currency
Randomize Timer
With DataDIC.Recordset
.MoveLast
H = Rnd * .RecordCount
.MoveFirst
For i = 1 To H - 1
.MoveNext
Next
Label1.Caption = .Fields("latin")
Javab = .Fields("farsi")
Label2.Caption = Javab
Length = Len(Javab)
End With
End Sub
