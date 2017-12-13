VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8970
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   15345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "user info1.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   7680
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "user info1.frx":4A9C2
      Left            =   5760
      List            =   "user info1.frx":4A9C4
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   5760
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "mobile number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "user name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "user id"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TICKET COLLECTOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Combo1_Click()
'code to match the data in list box and display in text box
rs.MoveFirst
Do While Not rs.EOF
If Combo1.List(Combo1.ListIndex) = rs(0) Then
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
Text4.Text = rs(3)
Exit Sub
Else
rs.MoveNext
End If
Loop
End Sub

Private Sub Command1_Click()
'update button code
rs(0) = Val(Text1.Text)
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
rs.Update
MsgBox " Records are updated", vbInformation + vbOKOnly, "update"
End Sub
Private Sub Command2_Click()
'Add button CODE
rs.AddNew
Command3.Visible = True
Command2.Visible = False
MsgBox "records added", vbInformation + vbOKOnly, "add"
End Sub

Private Sub command3_click()
'Save Button code
rs(0) = Val(Text1.Text)
rs(1) = Text2.Text
rs(2) = Text3.Text
rs(3) = Text4.Text
 rs.Update
MsgBox "record are saved", vbInformation + vbOKOnly, "save"
End Sub

Private Sub Command4_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
'code to open database and table,make sure that u give the path correctly and table name
Command3.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\edu.softy\5th semester\project\mts\db2.mdb;Persist Security Info=False"
rs.Open "select* from userinfo", db, adOpenDynamic, adLockOptimistic
MsgBox "open"
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop

End Sub






