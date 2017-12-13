VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000008&
   Caption         =   "MDIForm1"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15345
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Menu bl 
      Caption         =   "BILLING"
      WindowList      =   -1  'True
      Begin VB.Menu bn 
         Caption         =   "Booking new"
      End
      Begin VB.Menu cb 
         Caption         =   "Billing list"
      End
   End
   Begin VB.Menu cd 
      Caption         =   "CUSTOMER DETAILS"
      Begin VB.Menu ad1 
         Caption         =   "Add"
      End
      Begin VB.Menu dt1 
         Caption         =   "Delete"
      End
      Begin VB.Menu ud1 
         Caption         =   "Update"
      End
      Begin VB.Menu rep1 
         Caption         =   "Customer list"
      End
   End
   Begin VB.Menu ui 
      Caption         =   "USER INFO"
      Begin VB.Menu ad2 
         Caption         =   "Add"
      End
      Begin VB.Menu dt2 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu md2 
      Caption         =   "MOVIE DETAILS"
      Begin VB.Menu ad3 
         Caption         =   "Add"
      End
      Begin VB.Menu dt3 
         Caption         =   "Delete"
      End
      Begin VB.Menu ud3 
         Caption         =   "Update"
      End
      Begin VB.Menu ml 
         Caption         =   "Movies list"
      End
   End
   Begin VB.Menu ET 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a12_Click()
Form11.Show
End Sub

Private Sub ad1_Click()
Form5.Show
Form5.Combo1.Visible = False
Form5.Command1.Visible = False
End Sub

Private Sub ad2_Click()
Form2.Show
'Form2.Combo1.Visible = True
'Form2.Command1.Visible = True
End Sub

Private Sub ad3_Click()
Form4.Show
Form4.Combo1.Visible = False
Form4.Command1.Visible = False
End Sub

Private Sub bn_Click()
Form3.Show
End Sub

Private Sub cb_Click()
DataReport1.Show
End Sub

Private Sub dt1_Click()
Form8.Show
End Sub

Private Sub dt2_Click()
Form6.Show
End Sub

Private Sub dt3_Click()
Form7.Show
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub ET_Click()
Unload Me
Form1.Show
End Sub



Private Sub ml_Click()
DataReport3.Show
End Sub

Private Sub rep1_Click()
DataReport2.Show

End Sub

Private Sub ud1_Click()
Form5.Show
'Form5.Command2.Visible = False

End Sub

Private Sub ud2_Click()
Form2.Show
Form2.Command2.Visible = False
End Sub

Private Sub ud3_Click()
Form4.Show
Form4.Command2.Visible = False
End Sub
