VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   Caption         =   "WEL COME TO CINEMAX"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form10"
   Picture         =   "wel come.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   15285
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   12
      Top             =   7080
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   11040
      Picture         =   "wel come.frx":2D975
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ASHIM MAHARJAN"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   10560
      Picture         =   "wel come.frx":2E2B7
      Top             =   4080
      Width           =   4320
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "click here"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11520
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPORT AND ENCOURAGEMENT"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   10920
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MADAM LAXMI "
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MADAM KALAI SELVI(HOD)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   6
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "UNDER THE GUIDANCE "
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NEVILLE BAJIKA"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "UMANGA DEEP SHRESTHA"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AMIT SHRESTHA"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMITTED BY"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WELCOME TO FUN CINEMA"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Form1.Show
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label14_Click()
Form1.Show
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Picture1_Click()
Form1.Show
End Sub
