VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   " HOSPITAL MANAGEMENT"
   ClientHeight    =   10650
   ClientLeft      =   1110
   ClientTop       =   900
   ClientWidth     =   20175
   LinkTopic       =   "MDIForm1"
   Picture         =   "main form.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu patient 
      Caption         =   "Patient"
   End
   Begin VB.Menu doctor 
      Caption         =   "Doctor"
   End
   Begin VB.Menu nurse 
      Caption         =   "Nurse"
   End
   Begin VB.Menu worktim 
      Caption         =   "WorkTime"
   End
   Begin VB.Menu blood 
      Caption         =   "blood bank"
   End
   Begin VB.Menu billing 
      Caption         =   "billing"
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
Form6.Show
End Sub

Private Sub billing_Click()
Form8.Show
End Sub

Private Sub blood_Click()
Form7.Show
End Sub

Private Sub doctor_Click()
Form3.Show
End Sub

Private Sub exit_Click()
confirmation = MsgBox("Do you want to EXIT", vbYesNo + vbCritical, "EXIT confirmation")
If confirmation = vbYes Then
End
MsgBox "EXIT has been  sucessfully", vbInformation, "Message"
Else
MsgBox " EXIT NOT SUCESSFULLY !!!", vbInformation, "Message"
End If
End Sub

Private Sub nurse_Click()
Form4.Show
End Sub

Private Sub patient_Click()
addmissionpatient.Show
End Sub

Private Sub worktim_Click()
Form5.Show
End Sub
