VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H0000FFFF&
   Caption         =   "nurse deatils"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12840
      Top             =   7680
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      Connect         =   $"nurse details.frx":0000
      OLEDBString     =   $"nurse details.frx":00A3
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
      Left            =   13920
      Picture         =   "nurse details.frx":0146
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   12120
      Picture         =   "nurse details.frx":59FD
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   13920
      Picture         =   "nurse details.frx":B2C9
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   12120
      Picture         =   "nurse details.frx":10A58
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   8280
      Picture         =   "nurse details.frx":16303
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   5160
      Picture         =   "nurse details.frx":1BB8A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3240
      Picture         =   "nurse details.frx":21414
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5160
      Picture         =   "nurse details.frx":26CF1
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   3240
      Picture         =   "nurse details.frx":2C59C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9000
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FF8080&
      DataField       =   "nurse email id"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   8160
      Width           =   4815
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FF8080&
      DataField       =   "NURSE WORKING TIME"
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
      Height          =   525
      Left            =   4200
      TabIndex        =   20
      Top             =   7440
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FF8080&
      DataField       =   "NURSE SALLARY"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   6480
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF8080&
      DataField       =   "DOB"
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
      Left            =   4200
      TabIndex        =   18
      Top             =   5760
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF8080&
      DataField       =   "PHONE NUMBERS"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   4920
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF8080&
      DataField       =   "ADDRESSS"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF8080&
      DataField       =   "nurse department"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF8080&
      DataField       =   "NURSE EDUCATION"
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
      Left            =   4200
      TabIndex        =   14
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF8080&
      DataField       =   "FATHER NAME"
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
      Height          =   525
      Left            =   4200
      TabIndex        =   13
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      DataField       =   "NURSE NAME"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      DataField       =   "NURSE ID"
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
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   8385
      Left            =   9240
      Picture         =   "nurse details.frx":31E26
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10890
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "EMAIL ID"
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
      Left            =   240
      TabIndex        =   10
      Top             =   8160
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SALLERY"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DATE OF BIRTH"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "NURSE EDUCATION"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
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
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "NURSE NAME"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "NURSE ID"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form4"
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
Adodc1.Recordset.Fields("nurse id") = Text1.Text
Adodc1.Recordset.Fields("nurse name") = Text2.Text
Adodc1.Recordset.Fields("father name") = Text3.Text
Adodc1.Recordset.Fields("nurse education") = Text4.Text
Adodc1.Recordset.Fields("nurse department") = Text5.Text
Adodc1.Recordset.Fields("ADDRESSS") = Text6.Text
Adodc1.Recordset.Fields("PHONE NUMBERS") = Text7.Text
Adodc1.Recordset.Fields("DOB") = Text8.Text
Adodc1.Recordset.Fields("NURSE SALLARY") = Text9.Text
Adodc1.Recordset.Fields("NURSE WORKING TIME") = Text10.Text
Adodc1.Recordset.Fields("nurse email id") = Text11.Text
End Sub

Private Sub Command7_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "delete record confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been delete sucessfully", vbInformation, "Message"
Else
MsgBox " Record not delete !!!", vbInformation, "Message"
End If
Adodc1.Recordset.Delete
End Sub

Private Sub Command8_Click()
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
End Sub

Private Sub Command9_Click()
frmmain.Show
Form4.Hide
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
If KeyCode = 13 Then Text1.SetFocus

End Sub
