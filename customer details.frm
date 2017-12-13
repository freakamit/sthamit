VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   " "
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Picture         =   "customer details.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   15735
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   13
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000015&
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   6240
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000015&
      Height          =   615
      Left            =   5400
      TabIndex        =   11
      Top             =   4920
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000015&
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000015&
      Height          =   615
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000015&
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of The Movie"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form5"
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
DataReport2.Show
End Sub

Private Sub Form_Load()
'code to open database and table,make sure that u give the path correctly and table name
Command2.Visible = True
Command3.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\edu.softy\5th semester\project\mts\db2.mdb;Persist Security Info=False"
rs.Open "select* from customerdetails", db, adOpenDynamic, adLockOptimistic
MsgBox "open"
rs.MoveFirst
Do While Not rs.EOF
Combo1.AddItem rs(0)
rs.MoveNext
Loop

End Sub






