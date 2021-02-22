VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H0000FFFF&
   Caption         =   "employe work shedule"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13920
      Top             =   8760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   $"employee.frx":0000
      OLEDBString     =   $"employee.frx":0098
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EMPLOYEE "
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
      Left            =   16680
      Picture         =   "employee.frx":0130
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   15000
      Picture         =   "employee.frx":59E7
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   13200
      Picture         =   "employee.frx":B176
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   11520
      Picture         =   "employee.frx":10A42
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   8400
      Picture         =   "employee.frx":162ED
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   5400
      Picture         =   "employee.frx":1BB74
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3720
      Picture         =   "employee.frx":213FE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1920
      Picture         =   "employee.frx":26CDB
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   360
      Picture         =   "employee.frx":2C586
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      DataField       =   "SHIFT"
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      DataField       =   "DATE"
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      DataField       =   "NAME"
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      DataField       =   "EMPLOYEE ID"
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   6285
      Left            =   120
      Picture         =   "employee.frx":31E10
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   7770
   End
   Begin VB.Image Image1 
      Height          =   9945
      Left            =   7920
      Picture         =   "employee.frx":853E0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12720
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SHIFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "NAME "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "EMPLOYEE ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form5"
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
Adodc1.Recordset.Fields("EMPLOYEE ID") = Text1.Text
Adodc1.Recordset.Fields("name") = Text2.Text
Adodc1.Recordset.Fields("department") = Text3.Text
Adodc1.Recordset.Fields("date") = Text4.Text
Adodc1.Recordset.Fields("shift") = Text5.Text

End Sub

Private Sub Command7_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "

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
Form5.Hide
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
If KeyCode = 13 Then Text1.SetFocus
End Sub
