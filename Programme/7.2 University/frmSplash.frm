VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00004080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4110
   ClientLeft      =   3990
   ClientTop       =   1095
   ClientWidth     =   7350
   ClipControls    =   0   'False
   DrawMode        =   1  'Blackness
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_date 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "security_time_backups"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3105
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3225
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   165
      Left            =   -120
      Picture         =   "frmSplash.frx":12ED2
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   7575
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language school"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   4185
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "prince soft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   7
      Top             =   360
      Width           =   1845
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Platform"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5190
      TabIndex        =   6
      Top             =   2040
      Width           =   1785
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Company                     Prince soft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2310
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright                October 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2100
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ": ·ÿ›« —„“ ⁄»Ê— —« Ê«—œ ‰„«ÌÌœ"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4125
      Left            =   0
      Picture         =   "frmSplash.frx":2277C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7365
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
SaveSetting "mk", "mk", "mk", "I'm running"
window_style = GetSetting("mk", "mk", "style")
data_date.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "mk", "mk", "mk", "I'm close"
End
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
a = LCase(Text1.Text)
If KeyCode = 13 Then
If data_date.Recordset.RecordCount = 0 Then GoTo 3
data_date.Recordset.MoveFirst
pass = data_date.Recordset.Fields("entrance_password")
If IsNull(pass) Then pass = ""
If decoding(pass) = a Then
3: Me.Hide
mainshow
End If
End If
End Sub
