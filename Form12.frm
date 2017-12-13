VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TICKET CANCEL"
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
Me.WindowState = 2
i = 0
End Sub
