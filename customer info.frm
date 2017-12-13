VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000008&
   Caption         =   "Form8"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15315
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "customer info.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6840
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3960
      MaskColor       =   &H0000FF00&
      TabIndex        =   3
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
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
      Left            =   7920
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "customer id"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Form8"
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
MsgBox "Do you really want to delete", vbCritical + vbOKCancel, "Delete"
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
rs.Open "select * from customerdetails", db, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop

End Sub


