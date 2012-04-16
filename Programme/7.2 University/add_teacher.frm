VERSION 5.00
Begin VB.Form add_teacher 
   Caption         =   "«÷«›Â ”«“Ì „⁄·„"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4920
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
   ScaleHeight     =   4575
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   " «ÌÌœ"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   90
      Picture         =   "add_teacher.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   840
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
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   465
      Width           =   1335
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
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   465
      Width           =   1335
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
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1335
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "teachers"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox Check1 
      Caption         =   "„⁄·„ »Ì‘ —"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   9
      Top             =   930
      Width           =   1305
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
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   1500
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "ÊÌ—«Ì‘"
         Height          =   1035
         Left            =   210
         Picture         =   "add_teacher.frx":0415
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1515
         Width           =   1890
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ–›"
         Height          =   870
         Left            =   210
         Picture         =   "add_teacher.frx":0A4B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   585
         Width           =   1890
      End
      Begin VB.ListBox List1 
         Height          =   1980
         Left            =   2445
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   2085
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   ": ‰«„ „⁄·„"
         Height          =   240
         Left            =   3210
         TabIndex        =   16
         Top             =   270
         Width           =   630
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
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   4695
      Begin VB.TextBox RichTextBox1 
         Alignment       =   1  'Right Justify
         Height          =   2205
         Left            =   165
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   645
         Width           =   4305
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Description :"
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
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   ":  Ê÷ÌÕ« "
         Height          =   240
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Label Label18 
      Caption         =   ": ‰«„ „⁄·„"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ":  ·›‰"
      Height          =   240
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ": ¬œ—”"
      Height          =   240
      Index           =   0
      Left            =   1560
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ": ‰«„ Œ«‰Ê«œêÌ „⁄·„"
      Height          =   240
      Left            =   3600
      TabIndex        =   11
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ": ÕﬁÊﬁ œ— 20 Ã·”Â"
      Height          =   240
      Index           =   1
      Left            =   3570
      TabIndex        =   10
      Top             =   840
      Width           =   1380
   End
   Begin VB.Menu Edit 
      Caption         =   "&ÊÌ—«Ì‘"
   End
End
Attribute VB_Name = "add_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_bol As Boolean
Dim edit_num As Integer
Private Sub Command1_Click()
For i = 0 To Text1.UBound
If Text1(i).Text = "" Then r = True
Next
If r = True Then
a = MsgBox("You're having any empty places.", vbInformation, "tell me what")
End If
a = MsgBox("Do you sure you want to save?", vbQuestion + vbYesNo, "tell me what")
If a = vbYes Then
If edit_bol = True Then
Data1.Recordset.MoveFirst
For i = 1 To edit_num
Data1.Recordset.MoveNext
Next
Data1.Recordset.Edit
Else
Data1.Recordset.AddNew
End If
Data1.Recordset.Fields("name") = Text1(0)
Data1.Recordset.Fields("family") = Text1(1)
Data1.Recordset.Fields("phone") = Text1(2)
Data1.Recordset.Fields("adress") = Text1(3)
Data1.Recordset.Fields("sallary") = Text1(4)
Data1.Recordset.Fields("description") = RichTextBox1.Text
Data1.Recordset.Update
End If
comment_now (".«÷«›Â”«“Ì ê—ÊÂ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ")
If Check1.Value = 0 Then
Unload Me
mainshow
End If
edit_bol = False
End Sub

Private Sub Command2_Click()
edit_bol = False
If List1.ListIndex > -1 Then
a = MsgBox("¬Ì¬ „ÿ„∆‰ Â” Ìœ òÂ «Ì‰ „⁄·„ —« Õ–› ò‰„", vbQuestion + vbYesNo)
If a = vbYes Then
Data1.Recordset.MoveFirst
For i = 1 To List1.ListIndex
Data1.Recordset.MoveNext
Next
Data1.Recordset.Delete
List1.RemoveItem (List1.ListIndex)
comment_now (".„⁄·„ «‰ Œ«» ‘œÂ «“ ·Ì”  Œ«—Ã ‘œ")
End If
End If
End Sub

Private Sub Command3_Click()
If List1.ListIndex > -1 Then
a = MsgBox("¬Ì¬ „«Ì· »Â ÊÌ—«Ì‘ «Ì‰ „⁄·„ Â” Ìœ", vbQuestion + vbYesNo)
If a = vbYes Then
edit_num = List1.ListIndex
Data1.Recordset.MoveFirst
For i = 1 To List1.ListIndex
Data1.Recordset.MoveNext
Next
Frame2.Visible = False
edit_bol = True
Text1(0) = Data1.Recordset.Fields("name")
Text1(1) = Data1.Recordset.Fields("family")
Text1(2) = Data1.Recordset.Fields("phone")
Text1(3) = Data1.Recordset.Fields("adress")
Text1(4) = Data1.Recordset.Fields("sallary")
RichTextBox1.Text = Data1.Recordset.Fields("description")
comment_now (".„‘Œ’«  „⁄·„ «‰ Œ«» ‘œÂ »« „Ê›ﬁÌ  ÊÌ—«Ì‘ ‘œ")
End If
End If
End Sub

Private Sub Edit_Click()
Frame2.Visible = Not Frame2.Visible
edit_bol = False
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainshow
End Sub

Private Sub Form_Activate()
On Error Resume Next
List1.Clear
Data1.Recordset.MoveFirst
While Data1.Recordset.EOF = False
List1.AddItem (Data1.Recordset.Fields("family"))
Data1.Recordset.MoveNext
Wend
End Sub

