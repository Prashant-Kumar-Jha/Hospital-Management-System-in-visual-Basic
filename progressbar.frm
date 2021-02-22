VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9720
   ClientLeft      =   -4350
   ClientTop       =   17925
   ClientWidth     =   19260
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   19260
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   9240
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   5160
      Width           =   3015
   End
   Begin VB.PictureBox ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   19200
      TabIndex        =   0
      Top             =   8745
      Width           =   19260
   End
   Begin VB.Image Image2 
      Height          =   2280
      Left            =   18000
      Picture         =   "progressbar.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   0
      Picture         =   "progressbar.frx":3628
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "    PATNA MEDICAL AND REASERCH HOSIPTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   10080
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_click()
Timer1.Enabled = True

End Sub


Private Sub Label1_Click()
Label1.BackColor = vbRed
End Sub


Private Sub Timer1_Timer()
Timer1.Interval = Rnd * 300 + 10
ProgressBar1.Value = ProgressBar1.Value + 2
Label1.capttion = ProgressBar1.Value & "%"
If Label1.Caption = 100 & "%" Then
frmmain.Show
Unload Me
End If
End Sub
