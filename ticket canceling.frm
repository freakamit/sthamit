VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H80000007&
   Caption         =   "Form12"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form12"
   Picture         =   "ticket canceling.frx":0000
   ScaleHeight     =   9345
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TICKET CANCEL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      TabIndex        =   0
      Top             =   8280
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVIE TICKET CANCELATION "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seat number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
'code to cancel ticket
rs.MoveFirst
Do While Not rs.EOF
If Combo1.List(Combo1.ListIndex) = rs(10) Then
MsgBox "Do you really want to delete", vbCritical + vbOKCancel, "Delete"
rs.Delete
MsgBox "records are cancel your ticket"
Exit Sub
Else
rs.MoveNext
End If
Loop
End Sub

Private Sub Command2_Click()
MDIForm1.Show
End Sub

Private Sub Form_Load()
'code to open database and table, make sure that you give the path correctly and table name
'command4. enabled = false
'command2. enabled = false
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\mts\db2.mdb;Persist Security Info=False"
rs.Open "select * from billingdetails", db, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(10)
rs.MoveNext
Loop
End Sub



