VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   16755
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
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
      IMEMode         =   3  'DISABLE
      Left            =   8640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   7560
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
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
      Left            =   8640
      TabIndex        =   2
      Top             =   6600
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN HOME"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   8040
      TabIndex        =   6
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   6720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "logged in successfully", vbOKOnly + vbInformation, "login"
MDIForm1.Show
Text1.Text = ""
Text2.Text = ""
Exit Sub
End If
If Text1.Text <> "aaa" And Text2.Text <> "bbb" Then
MsgBox "Username and Password is Incorrect", vbOKOnly = vbInformation, "login"
Text1.Text = ""
Text2.Text = ""
Else
If Text1.Text <> "aaa" Then
MsgBox "Username is Incorrect", vbOKOnly + vbInformation, "Login"
Text1.Text = ""
Else
If Text2.Text <> "bbb" Then
MsgBox "Password is Incorrect", vbOKOnly + vbInformation, "login"
Text2.Text = ""
End If
End If
End If

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "logged in successfully", vbOKOnly + vbInformation, "login"
MDIForm1.Show
Text1.Text = ""
Text2.Text = ""
Exit Sub
End If
If Text1.Text <> "aaa" And Text2.Text <> "bbb" Then
MsgBox "Username and Password is Incorrect", vbOKOnly = vbInformation, "login"
Text1.Text = ""
Text2.Text = ""
Else
If Text1.Text <> "aaa" Then
MsgBox "Username is Incorrect", vbOKOnly + vbInformation, "Login"
Text1.Text = ""
Else
If Text2.Text <> "bbb" Then
MsgBox "Password is Incorrect", vbOKOnly + vbInformation, "login"
Text2.Text = ""
End If
End If
End If
End If

End Sub
