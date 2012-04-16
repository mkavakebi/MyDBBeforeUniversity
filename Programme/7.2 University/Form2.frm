VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form add_student 
   Caption         =   "À»  „‘Œ’«  œ«‰‘ ¬„Ê“"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1785
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2175
      Width           =   1305
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "students"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   2685
      Left            =   105
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   105
      Width           =   5070
      Begin VB.CheckBox Check1 
         Caption         =   "ç‰œ œ«‰‘ ¬„Ê“"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3330
         TabIndex        =   36
         Top             =   2115
         Width           =   1185
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
         Index           =   8
         Left            =   2655
         TabIndex        =   8
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2655
         TabIndex        =   2
         Top             =   585
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
         Index           =   5
         Left            =   2655
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   135
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   135
         TabIndex        =   5
         Top             =   960
         Width           =   1335
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
         Left            =   2655
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
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
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2655
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mask1 
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   1740
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox masked 
         Height          =   345
         Left            =   135
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": ‰«„ Œ«‰Ê«œêÌ"
         Height          =   255
         Left            =   1575
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":  «—ÌŒ  Ê·œ"
         Height          =   255
         Left            =   1575
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   ": Ã‰”Ì "
         Height          =   240
         Left            =   4095
         TabIndex        =   33
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   ": ‘.‘"
         Height          =   240
         Index           =   0
         Left            =   4095
         TabIndex        =   32
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": „·Ìˆ¯ "
         Height          =   285
         Left            =   1575
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": òœ “»«‰ ¬„Ê“"
         Height          =   255
         Left            =   1575
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": ê—ÊÂ"
         Height          =   255
         Index           =   0
         Left            =   4095
         TabIndex        =   29
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ": ‰«„"
         Height          =   255
         Index           =   1
         Left            =   4095
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":  «—ÌŒ À»  ‰«„"
         Height          =   240
         Index           =   1
         Left            =   1575
         TabIndex        =   27
         Top             =   1740
         Width           =   1020
      End
      Begin VB.Line Line1 
         X1              =   735
         X2              =   615
         Y1              =   2325
         Y2              =   2325
      End
      Begin VB.Line Line2 
         X1              =   735
         X2              =   735
         Y1              =   1965
         Y2              =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": ‘„«—Â  ·›‰"
         Height          =   240
         Index           =   1
         Left            =   4095
         TabIndex        =   26
         Top             =   1740
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2775
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "ÊÌ—«Ì‘"
         Height          =   855
         Left            =   120
         Picture         =   "Form2.frx":0409
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   1455
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
         Height          =   2535
         Left            =   1680
         TabIndex        =   12
         Top             =   170
         Width           =   3255
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
            TabIndex        =   16
            Top             =   840
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
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1935
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
            TabIndex        =   14
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
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
            TabIndex        =   13
            Top             =   1320
            Width           =   975
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
            TabIndex        =   17
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label12 
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
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label11 
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
            TabIndex        =   20
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label10 
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
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
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
            TabIndex        =   18
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ–›"
         Height          =   795
         Left            =   120
         Picture         =   "Form2.frx":08DF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
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
      Left            =   3465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "groups"
      Top             =   2220
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "add_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st_vir As Boolean
Dim edit_bol As Boolean
Dim j
Private Sub Check3_Click()
If Check3.Value = 1 Then List2.ZOrder (0)
If Check3.Value = 0 Then List1.ZOrder (0)
End Sub

Private Sub Command1_Click()
For i = 0 To Text1.UBound
If i = 3 Or i = 8 Or i = 10 Or i = 4 Then GoTo 1
If Text1(i).Text = "" Then r = True
1: Next
If r = True Then
a = MsgBox(".Â‰Ê“ Ã«Â«Ì Œ«·Ì œÌê—Ì „«‰œÂ «” ", vbInformation, "tell me what")
End If
a = MsgBox(".¬Ì¬ „ÿ„∆‰ Â” Ìœ òÂ –ŒÌ—Â ”«“Ì «‰Ã«„ ‘Êœ", vbQuestion + vbYesNo, "tell me what")
If a = vbYes Then
If edit_bol = True Then
edit_bol = False
Data1.Recordset.FindFirst ("code='" + Text1(6).Text + "'")
Data1.Recordset.Delete
End If
Data1.Recordset.AddNew
Data1.Recordset.Fields("name") = Text1(0)
Data1.Recordset.Fields("family") = Text1(1)
Data1.Recordset.Fields("gender") = Text1(2)
Data1.Recordset.Fields("tt") = masked.Text
Data1.Recordset.Fields("shsh") = Text1(5)
Data1.Recordset.Fields("code") = Text1(6)
Data1.Recordset.Fields("nationality") = Text1(7)
Data1.Recordset.Fields("group") = Combo1.Text
Data1.Recordset.Fields("ts") = mask1.Text
Data1.Recordset.Fields("phone") = Text1(8).Text
Data1.Recordset.Update
End If
''''''''''''pay form shower
'If Check1.Value = 0 Then
'Unload Me
'pay_st_form.Show
'pay_st_form.code_defualt = True
'End If
'''''''''''''''''''''
End Sub

Private Sub Command2_Click()
If Text1(3).Text <> "" Then
With Data1.Recordset
.FindFirst ("code='" + Text1(3) + "'")
If .NoMatch = False Then
.Delete
 '
 If List1.ListIndex >= 0 Then b = List1.ListIndex
 If List2.ListIndex >= 0 Then b = List2.ListIndex
 List1.RemoveItem (b)
 List2.RemoveItem (b)
 '
 edit_bol = True
Else
MsgBox ".òœ œ«‰‘ ¬„Ê“ „‘ò· œ«—œ", vbCritical
End If
End With
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1(3).Text <> "" Then
With Data1.Recordset
    .FindFirst ("code='" + Text1(3) + "'")
    If .NoMatch = False Then
        Frame2.Visible = False
        edit_bol = True
        Text1(0) = Data1.Recordset.Fields("name")
        Text1(1) = Data1.Recordset.Fields("family")
        Text1(2) = Data1.Recordset.Fields("gender")
        masked.Text = Data1.Recordset.Fields("tt")
        Text1(5) = Data1.Recordset.Fields("shsh")
        Text1(6) = Data1.Recordset.Fields("code")
        Text1(7) = Data1.Recordset.Fields("nationality")
        Combo1.List(Combo1.ListIndex) = Data1.Recordset.Fields("group")
        mask1.Text = Data1.Recordset.Fields("ts")
        Text1(8).Text = Data1.Recordset.Fields("phone")
        Combo1.Text = Data1.Recordset.Fields("group")
        Frame2.Visible = False
        edit_bol = True
        If List1.ListIndex >= 0 Then b = List1.ListIndex
        If List2.ListIndex >= 0 Then b = List2.ListIndex
        List1.RemoveItem (b)
        List2.RemoveItem (b)
    Else
        MsgBox ".òœ œ«‰‘ ¬„Ê“ „‘ò· œ«—œ", vbCritical
    End If
End With
End If
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
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\db1.mdb"
Data_groups.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainshow
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then mask1.Text = Replace(CStr(Date), "-", "/")
If Check2.Value = 0 Then mask1.Text = ""
End Sub
 
Function code_exist(ByVal code As String) As Boolean
Data1.Recordset.FindFirst ("code='" + code + "'")
code_exist = Not Data1.Recordset.NoMatch
End Function

Private Sub List1_Click()
Text1(3).Text = List2.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
Text1(3).Text = List2.List(List2.ListIndex)
End Sub
Private Sub Text2_Change()
List1.Clear
List2.Clear
On Error Resume Next
Data1.Recordset.MoveFirst
While Data1.Recordset.EOF = False
Nam = Data1.Recordset.Fields("name")
family = Data1.Recordset.Fields("family")
code = Data1.Recordset.Fields("code")
esm = family + " " + Nam
''''''''''''''''''''''''''''
If InStr(1, LCase(esm), Text2.Text) Then
List2.AddItem code
List1.AddItem esm
End If
Data1.Recordset.MoveNext
Wend
End Sub

