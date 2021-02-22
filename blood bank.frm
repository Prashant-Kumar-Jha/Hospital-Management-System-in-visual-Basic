VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00FF00FF&
   Caption         =   "blood bank"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12840
      Top             =   7920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   $"blood bank.frx":0000
      OLEDBString     =   $"blood bank.frx":0098
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
      Height          =   855
      Left            =   8040
      Picture         =   "blood bank.frx":0130
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      Height          =   2175
      Left            =   9840
      TabIndex        =   18
      Top             =   5160
      Width           =   9375
      Begin VB.CommandButton Command8 
         Height          =   615
         Left            =   6600
         Picture         =   "blood bank.frx":59E7
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   4440
         Picture         =   "blood bank.frx":B176
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Height          =   615
         Left            =   2520
         Picture         =   "blood bank.frx":10A42
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Height          =   615
         Left            =   360
         Picture         =   "blood bank.frx":162ED
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Height          =   2295
      Left            =   840
      TabIndex        =   13
      Top             =   5160
      Width           =   8055
      Begin VB.CommandButton Command4 
         Height          =   735
         Left            =   6120
         Picture         =   "blood bank.frx":1BB74
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Height          =   735
         Left            =   4080
         Picture         =   "blood bank.frx":213FE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   735
         Left            =   2040
         Picture         =   "blood bank.frx":26CDB
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   240
         Picture         =   "blood bank.frx":2C586
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "DONOR'S INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   18495
      Begin VB.TextBox Text6 
         BackColor       =   &H00FF8080&
         DataField       =   "BLOOD GROUP"
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
         Left            =   13320
         TabIndex        =   12
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FF8080&
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
         Left            =   13320
         TabIndex        =   11
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FF8080&
         DataField       =   "AGE"
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
         Left            =   13320
         TabIndex        =   10
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FF8080&
         DataField       =   "PHONE NO"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF8080&
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
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FF8080&
         DataField       =   "DONORS NAME"
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
         Left            =   3480
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Caption         =   "BLOOD GROUP"
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
         Left            =   9720
         TabIndex        =   9
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
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
         Height          =   495
         Left            =   9720
         TabIndex        =   8
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
         Caption         =   "AGE"
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
         Left            =   9720
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "PHONE NO"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H008080FF&
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
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "NAME"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form7"
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
Adodc1.Recordset.Fields("DONORS NAME") = Text1.Text
Adodc1.Recordset.Fields("ADDRESS") = Text2.Text
Adodc1.Recordset.Fields("PHONE NO") = Text3.Text
Adodc1.Recordset.Fields("AGE") = Text4.Text
Adodc1.Recordset.Fields("GENDER") = Text5.Text
Adodc1.Recordset.Fields("BLOOD GROUP") = Text6.Text
End Sub

Private Sub Command7_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
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
Form7.Hide
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
If KeyCode = 13 Then Text1.SetFocus
End Sub
