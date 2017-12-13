VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000009&
   Caption         =   "Form6"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15690
   LinkTopic       =   "Form6"
   Picture         =   "hit movie details.frx":0000
   ScaleHeight     =   10215
   ScaleWidth      =   15690
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
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
      Height          =   975
      Left            =   11640
      TabIndex        =   15
      Top             =   9600
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Select the Movie"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   10695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         DownPicture     =   "hit movie details.frx":5A6D8
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         Picture         =   "hit movie details.frx":6E879
         TabIndex        =   16
         Top             =   480
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         Picture         =   "hit movie details.frx":6F8C7
         TabIndex        =   14
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6720
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   2775
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   11
         Top             =   1800
         Width           =   3255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6720
         TabIndex        =   10
         Top             =   1800
         Width           =   3375
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   11640
      Width           =   9735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   360
      TabIndex        =   3
      Top             =   9360
      Width           =   6135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   360
      TabIndex        =   1
      Top             =   6000
      Width           =   6135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   13560
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "RATING"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   10920
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIRECTOR"
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
      Left            =   360
      TabIndex        =   7
      Top             =   8520
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACTRESS"
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
      Left            =   360
      TabIndex        =   6
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACTOR"
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
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HIT MOVIE DETAILS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
MDIForm1.Show
End Sub

Private Sub Option1_Click()
Text1.Text = "SHAHRUKH KHAN"
Text2.Text = "PRIYANKA CHOPRA"
Text3.Text = "FARHAN AKTAR"
End Sub

Private Sub Option2_Click()
Text1.Text = "SALMAN KHAN"
Text2.Text = "KATRINA KAIF"
Text3.Text = "KABIR KHAN"
End Sub

Private Sub Option3_Click()
Text1.Text = "IMRAN HASHMI"
Text2.Text = "ROOMAN KHAN AFZAL AHMED & BABUVA SHARMA"
Text3.Text = "MUKESH BHATT"
End Sub

Private Sub Option4_Click()
Text1.Text = "RANBIR KAPOOR"
Text2.Text = "PRIYANKA CHOPRA & ILLIYANA DISUZA"
Text3.Text = "ANURAG KASHYAP"
End Sub

Private Sub Option5_Click()
Text1.Text = "SHAHRUKH KHAN"
Text2.Text = "ANUSHKA SHARMA & KATRINA KAIF"
Text3.Text = "YASH CHOPRA"
End Sub

Private Sub Option6_Click()
Text1.Text = "AJAY DEVGAN"
Text2.Text = "SONAKSHI"
Text3.Text = "ROHIT SHETTY"
End Sub
