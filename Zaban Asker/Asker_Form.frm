VERSION 5.00
Begin VB.Form Asker_Form 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mr Bagheri Class "
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   495
      Top             =   135
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
      Top             =   570
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Height          =   945
      Left            =   3330
      TabIndex        =   1
      Top             =   15
      Width           =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "add"
      Height          =   210
      Left            =   45
      TabIndex        =   0
      Top             =   960
      Width           =   3405
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
      Height          =   345
      Left            =   15
      TabIndex        =   3
      Top             =   45
      Width           =   3285
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
      Height          =   510
      Left            =   15
      TabIndex        =   2
      Top             =   435
      Width           =   3270
   End
End
Attribute VB_Name = "Asker_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IsOpen As Boolean
Private Sub Command1_Click()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command3_Click()
Timer1.Enabled = fasle
Add_Form.Show
End Sub

Private Sub Form_Activate()
Make_New_Dic
End Sub

Private Sub Form_Load()
WindowPos Me, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Make_New_Dic
End Sub

Private Sub Make_New_Dic()
Label1.Caption = ""
Label2.Caption = ""

Left = Screen.Width - Width
Top = Screen.Height - Height - 500
ZOrder 0
Dim H As Currency
Dim G As Integer
Dim i As Currency
Randomize Timer
With Data1.Recordset
.MoveLast
H = Rnd * .RecordCount
.MoveFirst
For i = 1 To H
.MoveNext
Next
aa = .Fields("english")
bb = .Fields("farsi")
Label1.Caption = IIf(IsNull(aa), "", aa)
Label2.Caption = IIf(IsNull(bb), "", bb)
End With
End Sub

