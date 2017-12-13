VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000009&
   Caption         =   "Form4"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16035
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "movie detais.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   16035
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00404040&
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
      Left            =   3840
      TabIndex        =   9
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00404040&
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
      Left            =   3480
      TabIndex        =   8
      Top             =   5880
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00404040&
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
      Left            =   3480
      TabIndex        =   7
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
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
      Left            =   3480
      TabIndex        =   6
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
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
      Left            =   3840
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Actress"
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
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Actor"
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
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Movie Name"
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
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Movie ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIE DETAILS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
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

Private Sub Command5_Click()
DataReport3.Show
End Sub

Private Sub Form_Load()
'code to open database and table,make sure that u give the path correctly and table name
Command3.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\edu.softy\5th semester\project\mts\db2.mdb;Persist Security Info=False"
rs.Open "select* from moviedetails", db, adOpenDynamic, adLockOptimistic
MsgBox "open"
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop

End Sub







