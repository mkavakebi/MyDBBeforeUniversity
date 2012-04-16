VERSION 5.00
Begin VB.Form Add_Form 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delet"
      Height          =   345
      Left            =   1800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1275
      Width           =   930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&New"
      Height          =   345
      Left            =   825
      TabIndex        =   2
      Top             =   1275
      Width           =   930
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   300
      Index           =   3
      Left            =   3150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   930
      Width           =   765
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   300
      Index           =   2
      Left            =   2370
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   930
      Width           =   765
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   300
      Index           =   1
      Left            =   1590
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   930
      Width           =   765
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   300
      Index           =   0
      Left            =   825
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Width           =   765
   End
   Begin VB.TextBox Text2 
      DataField       =   "farsi"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   825
      TabIndex        =   1
      Top             =   540
      Width           =   3090
   End
   Begin VB.TextBox Text1 
      DataField       =   "english"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   825
      TabIndex        =   0
      Top             =   150
      Width           =   3090
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "ZAsker.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Loghat(EtoP)"
      Top             =   1290
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Persian :"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   585
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "English :"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   195
      Width           =   690
   End
End
Attribute VB_Name = "Add_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
s = MsgBox("are youe sure yoe want to delet it? ", vbQuestion + vbYesNo + vbDefaultButton2)
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
End Sub

Private Sub command_Click()

End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command3_Click(Index As Integer)
On Error Resume Next

Select Case Index
Case 0: Data1.Recordset.MoveFirst
Case 1: Data1.Recordset.MovePrevious
Case 2: Data1.Recordset.MoveNext
Case 3: Data1.Recordset.MoveLast
End Select
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\ZAsker.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Asker_Form.Show
Asker_Form.Timer1.Enabled = True
End Sub
