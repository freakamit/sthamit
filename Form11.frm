VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16290
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   4
      Top             =   7320
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   3
      Text            =   "75"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   2
      Text            =   "50"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   1
      Text            =   "50"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   7500
      Left            =   12000
      Picture         =   "Form11.frx":269B5
      Top             =   0
      Width           =   5070
   End
   Begin VB.Image Image1 
      Height          =   2910
      Index           =   2
      Left            =   8400
      Picture         =   "Form11.frx":2CCA2
      Top             =   2520
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   3180
      Index           =   1
      Left            =   4320
      Picture         =   "Form11.frx":2F1FC
      Top             =   2400
      Width           =   3570
   End
   Begin VB.Image Image1 
      Height          =   2910
      Index           =   0
      Left            =   120
      Picture         =   "Form11.frx":31865
      Top             =   2400
      Width           =   3885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Avilable food in theater"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.Show
End Sub
