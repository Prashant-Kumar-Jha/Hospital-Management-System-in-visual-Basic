VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H0000FFFF&
   Caption         =   "DOCTOR DEATILS"
   ClientHeight    =   3135
   ClientLeft      =   2895
   ClientTop       =   5025
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11400
      Top             =   8640
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"DOCTOR DEATILS.frx":0000
      OLEDBString     =   $"DOCTOR DEATILS.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DOCTOR"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Height          =   495
      Left            =   14760
      Picture         =   "DOCTOR DEATILS.frx":0146
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   13080
      Picture         =   "DOCTOR DEATILS.frx":59FD
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   11040
      Picture         =   "DOCTOR DEATILS.frx":B18C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   9240
      Picture         =   "DOCTOR DEATILS.frx":10A58
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   7800
      Picture         =   "DOCTOR DEATILS.frx":16303
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   5760
      Picture         =   "DOCTOR DEATILS.frx":1BB8A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4080
      Picture         =   "DOCTOR DEATILS.frx":21414
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2400
      Picture         =   "DOCTOR DEATILS.frx":26CF1
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   720
      Picture         =   "DOCTOR DEATILS.frx":2C59C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9840
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFC0C0&
      DataField       =   "WORKING TIME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   23
      Top             =   9000
      Width           =   4815
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFC0C0&
      DataField       =   "SALLARY"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   21
      Top             =   8040
      Width           =   4815
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFC0C0&
      DataField       =   "DEPARTMENT"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   20
      Top             =   7200
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFC0C0&
      DataField       =   "FATHERS NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   6480
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFC0C0&
      DataField       =   "EMAIL ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   5760
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFC0C0&
      DataField       =   "PHONE NUMBER"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      DataField       =   "ADDRESS"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFC0C0&
      DataField       =   "DATE OF BIRTH"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      DataField       =   "DOCTOR EDU"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      DataField       =   "GENDER"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      DataField       =   "DOCTOR NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      DataField       =   "DOCTOR ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   8865
      Left            =   9360
      Picture         =   "DOCTOR DEATILS.frx":31E26
      Stretch         =   -1  'True
      Top             =   840
      Width           =   10215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WORKING TIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   720
      TabIndex        =   22
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALLARY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   8160
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   720
      TabIndex        =   17
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FATHERS NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMAIL ID @"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PHONE NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE  OF BIRTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOCTOR    EDU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOCTOR NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOCTOR ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Fields("doctor id") = Text1.Text
Adodc1.Recordset.Fields("doctor name") = Text2.Text
Adodc1.Recordset.Fields("gender") = Text3.Text
Adodc1.Recordset.Fields("DOCTOR EDU") = Text4.Text
Adodc1.Recordset.Fields("date of birth") = Text5.Text
Adodc1.Recordset.Fields("address") = Text6.Text
Adodc1.Recordset.Fields("phone number") = Text7.Text
Adodc1.Recordset.Fields("email id") = Text8.Text
Adodc1.Recordset.Fields("fathers name") = Text9.Text
Adodc1.Recordset.Fields("department") = Text10.Text
Adodc1.Recordset.Fields("sallary") = Text11.Text
Adodc1.Recordset.Fields("working time") = Text12.Text
End Sub

Private Sub Command7_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
End Sub

Private Sub Command8_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "delete record confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been delete sucessfully", vbInformation, "Message"
Else
MsgBox " Record not delete !!!", vbInformation, "Message"
End If
Adodc1.Recordset.Delete
End Sub

Private Sub Command9_Click()
frmmain.Show
Form3.Hide
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus

End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus

End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus

End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus

End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus

End Sub
Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text7.SetFocus

End Sub
Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text8.SetFocus

End Sub
Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text9.SetFocus

End Sub
Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text10.SetFocus

End Sub
Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text11.SetFocus

End Sub
Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text12.SetFocus

End Sub
Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus

End Sub
