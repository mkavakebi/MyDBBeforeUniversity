VERSION 5.00
Begin VB.Form add_groupform 
   Caption         =   "«÷«›Â ”«“Ì ê—ÊÂ"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
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
   ScaleHeight     =   4605
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   105
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   105
      Width           =   4500
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
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   90
         TabIndex        =   4
         Top             =   630
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "add_groupform.frx":0000
         Left            =   2250
         List            =   "add_groupform.frx":0016
         TabIndex        =   1
         Top             =   630
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2250
         TabIndex        =   0
         Top             =   270
         Width           =   1335
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
         Left            =   2250
         TabIndex        =   2
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ":  ‰«„ ê—ÊÂ"
         Height          =   255
         Index           =   1
         Left            =   3690
         TabIndex        =   21
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   ": ‰«„ „⁄·„"
         Height          =   240
         Index           =   0
         Left            =   1500
         TabIndex        =   20
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ": ‘Â—ÌÂ"
         Height          =   240
         Index           =   1
         Left            =   1500
         TabIndex        =   19
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   ": ”ÿÕ"
         Height          =   240
         Index           =   1
         Left            =   3690
         TabIndex        =   18
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ":  —„"
         Height          =   240
         Index           =   0
         Left            =   3690
         TabIndex        =   17
         Top             =   975
         Width           =   345
      End
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "ê—ÊÂ ﬂÊœﬂ«‰"
      Height          =   240
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1710
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÊÌ—«Ì‘"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   795
      Picture         =   "add_groupform.frx":0076
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3525
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ–›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1860
      Picture         =   "add_groupform.frx":0695
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3525
      Width           =   1020
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2955
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "À»   ê—ÊÂ "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   300
      Picture         =   "add_groupform.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2595
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Left            =   2085
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "teachers"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
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
      Height          =   345
      Left            =   1815
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "groups"
      Top             =   3570
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   1425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "groups"
      Top             =   3165
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   825
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1545
      Width           =   4485
      Begin VB.TextBox Book_Text 
         BackColor       =   &H80000000&
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
         Left            =   2295
         TabIndex        =   12
         Top             =   435
         Width           =   1335
      End
      Begin VB.TextBox Film_Text 
         BackColor       =   &H80000000&
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
         TabIndex        =   11
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ": ‰«„ ﬂ «»"
         Height          =   240
         Index           =   2
         Left            =   3720
         TabIndex        =   14
         Top             =   435
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   ": ‰«„ ›Ì·„"
         Height          =   240
         Index           =   3
         Left            =   1530
         TabIndex        =   13
         Top             =   435
         Width           =   600
      End
   End
   Begin VB.Label Label18 
      Caption         =   ":  ‰«„ ê—ÊÂ"
      Height          =   255
      Index           =   0
      Left            =   3465
      TabIndex        =   7
      Top             =   2610
      Width           =   825
   End
End
Attribute VB_Name = "add_groupform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_bol As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
Frame1.Enabled = True
Book_Text.BackColor = vbWhite
Film_Text.BackColor = vbWhite
Else
Frame1.Enabled = False
Book_Text.BackColor = vbScrollBars
Film_Text.BackColor = vbScrollBars
Book_Text.Text = ""
Film_Text.Text = ""
End If
End Sub

Private Sub Command1_Click()
'''''''''error handling
If Combo3.Text = "" Then
MsgBox ".‰«„ ê—ÊÂ Ê«—œ ‰‘œÂ «” ", vbInformation
Combo3.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''
If Combo2.ListIndex = -1 Then
MsgBox ".›Ì·œ ”ÿÕ Œ«·Ì «” ", vbInformation
Combo3.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''
If Combo1.ListIndex = -1 Then
MsgBox ".‰«„ „⁄·„ «‰ Œ«» ‰‘œÂ «” ", vbInformation
Combo3.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Data1.Recordset.MoveFirst
While Data1.Recordset.EOF = False
If Data1.Recordset.Fields("name") = Combo3.Text Then
    If edit_bol <> True Then
        msg_yn = MsgBox("«Ì‰ ê—ÊÂ «“ ﬁ»· «÷«›Â ‘œÂ" + vbCrLf + "¬Ì¬ ¬‰ —« »Â —Ê“ ò „", vbQuestion + vbYesNo)
        If msg_yn = vbYes Then
            Data1.Recordset.Delete
            GoTo k
        Else
            Exit Sub
        End If
    Else
        Data1.Recordset.Delete
        GoTo k
    End If
End If
Data1.Recordset.MoveNext
Wend
k:
'''''''''''''''''''''''''''''''''''''''
Data1.Recordset.AddNew
Data1.Recordset.Fields("name") = Combo3.Text
Data1.Recordset.Fields("teacher") = Combo1.Text
Data1.Recordset.Fields("pay") = Text1.Text
Data1.Recordset.Fields("term") = Text2.Text
Data1.Recordset.Fields("level") = Combo2.Text
Data1.Recordset.Fields("book") = Book_Text.Text
Data1.Recordset.Fields("film") = Film_Text.Text
Data1.Recordset.Fields("type") = IIf(Check1.Value, "kids", "adults")
Data1.Recordset.Update
''''''''''''''''''''''''''''''''''''''''
comment_now (".«÷«›Â ”«“Ì ê—ÊÂ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ")
Unload Me
mainshow
End Sub

Private Sub Command2_Click()
If List1.ListIndex < 0 Then
MsgBox "you're had to choose an goup item first."
Exit Sub
End If
a = MsgBox("are you really decided to delete?", vbQuestion + vbYesNo)
If a = vbYes Then
With Data_groups.Recordset
f = List1.List(List1.ListIndex)
''''''''''''.FindFirst (f)
    Data_groups.Recordset.MoveFirst
    While Data_groups.Recordset.Fields("name") <> f
    Data_groups.Recordset.MoveNext
    Wend
''''''''''''''''''''''''''
.Delete
List1.RemoveItem (List1.ListIndex)
End With
End If
End Sub

Private Sub Command3_Click()
If List1.ListIndex < 0 Then
MsgBox "you should select a goup first."
Exit Sub
End If
With Data_groups.Recordset
f = List1.List(List1.ListIndex)
''''''.FindFirst ("name='" + List1.List(List1.ListIndex) + "'")
    Data_groups.Recordset.MoveFirst
    While Data_groups.Recordset.Fields("name") <> f
    Data_groups.Recordset.MoveNext
    Wend
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Combo1.Text = .Fields("teacher")
Combo2.Text = .Fields("level")
Combo3.Text = .Fields("name")
Text2.Text = .Fields("term")
Text1.Text = .Fields("pay")
If .Fields("type") = "kids" Then
Check1.Value = 1
Film_Text = .Fields("film")
Book_Text = .Fields("book")
Else
Check1.Value = 0
Film_Text = ""
Book_Text = ""
End If
edit_bol = True
End With
End Sub

Private Sub Form_Activate()
On Error Resume Next
Combo1.Clear
Data2.Recordset.MoveFirst
While Data2.Recordset.EOF = False
Combo1.AddItem (Data2.Recordset.Fields("family"))
Data2.Recordset.MoveNext
Wend
'''''''''list1.validate{groups}
On Error Resume Next
List.Clear
Data_groups.Recordset.MoveFirst
While Data_groups.Recordset.EOF = False
List1.AddItem (Data_groups.Recordset.Fields("name"))
Data_groups.Recordset.MoveNext
Wend
''''''''''''''''''
End Sub

Private Sub Form_Load()
For i = 65 To 90
Combo3.AddItem Chr(i)
Next
'''''''''''''''''''''''''
Data1.DatabaseName = App.Path + "\db1.mdb"
Data2.DatabaseName = App.Path + "\db1.mdb"
Data_groups.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainshow
End Sub

