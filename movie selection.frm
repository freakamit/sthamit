VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   Picture         =   "movie selection.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   11640
      Width           =   9735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF8080&
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
      TabIndex        =   10
      Top             =   8640
      Width           =   9855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
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
      TabIndex        =   9
      Top             =   7080
      Width           =   9855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
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
      TabIndex        =   8
      Top             =   5400
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
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
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   10335
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "AYAN"
         Height          =   855
         Left            =   6120
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "7TH SENSE"
         Height          =   855
         Left            =   3480
         TabIndex        =   6
         Top             =   2640
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "AGNIPATH"
         Height          =   855
         Left            =   840
         TabIndex        =   5
         Top             =   2640
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ARRYA 2"
         Height          =   855
         Left            =   6120
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CHINGARI"
         Height          =   855
         Left            =   3480
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DON 2"
         Height          =   855
         Left            =   840
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "RATING"
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   10920
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
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
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
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
      TabIndex        =   13
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
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
      TabIndex        =   12
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "HIT MOVIE DETAILS"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
Text1.Text = "shah rukh khan"
Text2.Text = "priyanka chopra"
Text3.Text = "farhan aktar"
End Sub

Private Sub Option2_Click()
Text1.Text = "darshan"
Text2.Text = "ramya"
Text3.Text = "RRGGGG"
End Sub

Private Sub Option3_Click()
Text1.Text = "allu arjun"
Text2.Text = "kajol agarwal"
Text3.Text = "puri jaganath"
End Sub

Private Sub Option4_Click()
Text1.Text = "hrithik roshan"
Text2.Text = "priyanka chopra"
Text3.Text = "yash chopra"
End Sub

Private Sub Option5_Click()
Text1.Text = "surya"
Text2.Text = "shruti hasan"
Text3.Text = "rajmaoili"
End Sub

Private Sub Option6_Click()
Text1.Text = "surya"
Text2.Text = "thamana"
Text3.Text = "hkghiy"
End Sub
