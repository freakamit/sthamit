VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000009&
   Caption         =   "Form7"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   Picture         =   "movie details.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   15465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "movie details.frx":2B991
      Left            =   3600
      List            =   "movie details.frx":2B99B
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIE DETAILS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Movie  ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
'code to delete
rs.MoveFirst
Do While Not rs.EOF
If Combo1.List(Combo1.ListIndex) = rs(0) Then
rs.Delete
MsgBox "records are deleted"
Exit Sub
Else
rs.MoveNext
End If
Loop
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
'code to open database and table, make sure that you give the path correctly and table name
'command4. enabled = false
'command2. enabled = false
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\edu.softy\5th semester\project\mts\db2.mdb;Persist Security Info=False"
rs.Open "select * from moviedetails", db, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop

End Sub

