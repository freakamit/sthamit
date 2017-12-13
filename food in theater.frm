VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000008&
   Caption         =   "Form6"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16290
   LinkTopic       =   "Form11"
   Picture         =   "food in theater.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   11880
      TabIndex        =   0
      Top             =   7200
      Width           =   3855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.Show
End Sub

