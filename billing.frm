VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00FFFF00&
   Caption         =   "billing"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   13200
      Top             =   10320
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
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
      Connect         =   $"billing.frx":0000
      OLEDBString     =   $"billing.frx":009C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BILL"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Height          =   975
      Left            =   8640
      Picture         =   "billing.frx":0138
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF00FF&
      Caption         =   "ACTION"
      Height          =   1575
      Left            =   10200
      TabIndex        =   29
      Top             =   7320
      Width           =   9495
      Begin VB.CommandButton Command9 
         Height          =   615
         Left            =   7200
         Picture         =   "billing.frx":59EF
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Height          =   615
         Left            =   4800
         Picture         =   "billing.frx":B17E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   2520
         Picture         =   "billing.frx":10A4A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Height          =   615
         Left            =   120
         Picture         =   "billing.frx":162F5
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF00FF&
      Caption         =   "NAVIGATION"
      Height          =   1575
      Left            =   840
      TabIndex        =   28
      Top             =   7320
      Width           =   8895
      Begin VB.CommandButton Command5 
         Height          =   615
         Left            =   7080
         Picture         =   "billing.frx":1BB7C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   4560
         Picture         =   "billing.frx":21406
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   2280
         Picture         =   "billing.frx":26CE3
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   240
         Picture         =   "billing.frx":2C58E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF00FF&
      Caption         =   "CHARGES"
      Height          =   3615
      Left            =   840
      TabIndex        =   13
      Top             =   3240
      Width           =   18735
      Begin VB.TextBox Text13 
         DataField       =   "TOTAL "
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   9720
         TabIndex        =   27
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TOTAL CHARGES"
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
         Left            =   6000
         TabIndex        =   26
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox Text12 
         DataField       =   "OTHER CHARES"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   15360
         TabIndex        =   25
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text11 
         DataField       =   "HOSPITAL"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   8760
         TabIndex        =   23
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         DataField       =   "MEDICINE"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2640
         TabIndex        =   21
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text9 
         DataField       =   "OT CHARGES"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   15360
         TabIndex        =   19
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         DataField       =   "OT CHARGES"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   8760
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         DataField       =   "PATHOLOGY"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "OTHER CHARGES"
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
         Left            =   12120
         TabIndex        =   24
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "HOSPITAL "
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
         Left            =   5880
         TabIndex        =   22
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "MEDICINE "
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
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "ICU CHARGE"
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
         Left            =   12120
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "OT CHARGES"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PATHOLOGY"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Caption         =   "patient's_info"
      Height          =   2775
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   18735
      Begin VB.TextBox Text6 
         DataField       =   "DISCHARGE DATE"
         DataSource      =   "Adodc1"
         Height          =   360
         Left            =   15360
         TabIndex        =   12
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         DataField       =   "ENTRY DATE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         DataField       =   "DEPARTMENT"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         DataField       =   "PAITENT NAME"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   15360
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         DataField       =   "PAITENT ID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   9360
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         DataField       =   "BILL NO"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "DISCHARGE DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12360
         TabIndex        =   11
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "ENTRY DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "PATIENT NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12360
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "PATIENT ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "BILL NO."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text13.Text = Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
End Sub

Private Sub Command10_Click()
frmmain.Show
Form8.Hide
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Fields("bill no") = Text1.Text
Adodc1.Recordset.Fields("PAITENT ID") = Text2.Text
Adodc1.Recordset.Fields("PAITENT name") = Text3.Text
Adodc1.Recordset.Fields("DEPARTMENT") = Text4.Text
Adodc1.Recordset.Fields("entry date") = Text5.Text
Adodc1.Recordset.Fields("discharge date") = Text6.Text
Adodc1.Recordset.Fields("PATHOLOGY") = Text7.Text
Adodc1.Recordset.Fields("OT CHARGES") = Text8.Text
Adodc1.Recordset.Fields("icu charges") = Text9.Text
Adodc1.Recordset.Fields("MEDICINE") = Text10.Text
Adodc1.Recordset.Fields("hospital") = Text11.Text
Adodc1.Recordset.Fields("OTHER CHARES") = Text12.Text
Adodc1.Recordset.Fields("TOTAL ") = Text13.Text

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
Text12.Text = " "
Text13.Text = " "
End Sub

Private Sub Command9_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "delete record confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been delete sucessfully", vbInformation, "Message"
Else
MsgBox " Record not delete !!!", vbInformation, "Message"
End If
Adodc1.Recordset.Delete
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
If KeyCode = 13 Then Text13.SetFocus

End Sub
Private Sub Text13_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus

End Sub



