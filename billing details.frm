VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15390
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   Picture         =   "billing details.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   15390
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   29
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   855
      Left            =   13680
      TabIndex        =   28
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   26
      Top             =   8520
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4560
      TabIndex        =   25
      Top             =   3000
      Width           =   4335
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   9360
      TabIndex        =   24
      Top             =   1200
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16744576
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   66584577
      CurrentDate     =   40975
   End
   Begin VB.TextBox Text6 
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
      Left            =   12240
      TabIndex        =   23
      Top             =   7080
      Width           =   3255
   End
   Begin VB.TextBox Text5 
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
      Left            =   12240
      TabIndex        =   22
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox Text3 
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
      Left            =   4560
      TabIndex        =   21
      Top             =   7080
      Width           =   4335
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   13080
      TabIndex        =   20
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
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
      Left            =   4560
      TabIndex        =   17
      Top             =   6000
      Width           =   4335
   End
   Begin VB.TextBox Text1 
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
      Left            =   4560
      TabIndex        =   16
      Top             =   4920
      Width           =   4335
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
      Left            =   4560
      TabIndex        =   15
      Top             =   3960
      Width           =   4335
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   13080
      TabIndex        =   13
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE BILL"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      TabIndex        =   10
      Top             =   8520
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCEED"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   9
      Top             =   8520
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "cash/card"
      DragMode        =   1  'Automatic
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
      Left            =   4560
      TabIndex        =   4
      Top             =   1920
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Seat No."
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
      Left            =   12600
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF8080&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9120
      TabIndex        =   18
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9120
      TabIndex        =   12
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
      DragMode        =   1  'Automatic
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
      Left            =   360
      TabIndex        =   8
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Count(persons)"
      DragMode        =   1  'Automatic
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
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movie name"
      DragMode        =   1  'Automatic
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
      Left            =   360
      TabIndex        =   3
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movie watching on"
      DragMode        =   1  'Automatic
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
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Booking via"
      DragMode        =   1  'Automatic
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
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BIILLING DETAILS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset

Private Sub Combo1_Click()
'get data from moviedetails
rs2.MoveFirst
Do While Not rs2.EOF
If Combo1.List(Combo1.ListIndex) = rs2(0) Then
Text1.Text = rs2(1)
Exit Sub
Else
rs2.MoveNext
End If
Loop
End Sub
Private Sub Combo2_Click()
'get data from customer details
rs3.MoveFirst
Do While Not rs3.EOF
If Combo2.List(Combo2.ListIndex) = rs3(0) Then
Text5.Text = rs3(1)
Exit Sub
Else
rs3.MoveNext
End If
Loop
End Sub

Private Sub Combo3_Click()
'get data from userinfo
rs4.MoveFirst
Do While Not rs4.EOF
If Combo3.List(Combo3.ListIndex) = rs4(0) Then
Text6.Text = rs4(1)
Exit Sub
Else
rs4.MoveNext
End If
Loop
End Sub

Private Sub Command1_Click()
'processing command
Command1.Enabled = True
rs1.AddNew
MsgBox "billing is processed", vbInformation + vbOKOnly, "processed"
End Sub

Private Sub Command2_Click()
'save bill
Command1.Enabled = False
If Option1.Value = True Then
rs1(0) = "Cash"
Else
If Option2.Value = True Then
rs1(0) = "Card"
End If
End If
rs1(1) = Text7.Text
rs1(2) = Combo1.Text
rs1(3) = Text1.Text
rs1(4) = Val(Text2.Text)
rs1(5) = Text3.Text
rs1(6) = Combo2.Text
rs1(7) = Text5.Text
rs1(8) = Combo3.Text
rs1(9) = Text6.Text
rs1(10) = Text4.Text
rs1.Update
MsgBox "Bill is saved"
Unload Form3
rs1.Close
db.Close
End Sub

Private Sub command3_click()
DataReport1.Show
End Sub

Private Sub Command4_Click()
MDIForm1.Show
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text7.Text = MonthView1.Value
End Sub

Private Sub Text3_GotFocus()
Text3.Text = Val(Text2.Text) * 150
End Sub
Private Sub Form_Load()
Command1.Enabled = True
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\mts\db2.mdb;Persist Security Info=False"
rs1.Open "select * from billingdetails", db, adOpenDynamic, adLockOptimistic
rs2.Open "select * from moviedetails", db, adOpenDynamic, adLockOptimistic
rs3.Open "select * from customerdetails", db, adOpenDynamic, adLockOptimistic
rs4.Open "select * from userinfo", db, adOpenDynamic, adLockOptimistic
MonthView1.Visible = False
rs2.MoveFirst
'to load movie id
Do While Not rs2.EOF
Combo1.AddItem rs2(0)
rs2.MoveNext
Loop
rs3.MoveFirst
'to load movie id
Do While Not rs3.EOF
Combo2.AddItem rs3(0)
rs3.MoveNext
Loop
rs4.MoveFirst
'to load movie id
Do While Not rs4.EOF
Combo3.AddItem rs4(0)
rs4.MoveNext
Loop

End Sub

Private Sub Text7_GotFocus()
MonthView1.Visible = True

End Sub
