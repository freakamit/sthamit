VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000012&
   Caption         =   "Form6"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   Picture         =   "user info.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
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
      Height          =   435
      Left            =   8880
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
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
      Height          =   435
      Left            =   5160
      TabIndex        =   0
      Top             =   4740
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "TICKET COLLECTOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   5040
      TabIndex        =   2
      Top             =   720
      Width           =   4185
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "USERID"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
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
rs.Open "select * from userinfo", db, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop
End Sub


