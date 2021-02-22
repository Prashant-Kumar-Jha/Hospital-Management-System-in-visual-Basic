VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addmissionpatient 
   Caption         =   "Addmission  Patient"
   ClientHeight    =   9375
   ClientLeft      =   1815
   ClientTop       =   1515
   ClientWidth     =   18915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   18915
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         TabIndex        =   22
         Top             =   8400
         Width           =   18975
         Begin VB.CommandButton Command9 
            Height          =   495
            Left            =   15120
            Picture         =   "patientadd.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Height          =   495
            Left            =   13680
            Picture         =   "patientadd.frx":58B7
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
            Height          =   495
            Left            =   12120
            Picture         =   "patientadd.frx":B17E
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command6 
            Height          =   495
            Left            =   10680
            Picture         =   "patientadd.frx":1090D
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Height          =   495
            Left            =   9240
            Picture         =   "patientadd.frx":161B8
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Height          =   495
            Left            =   7080
            Picture         =   "patientadd.frx":1BA3F
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Height          =   495
            Left            =   5640
            Picture         =   "patientadd.frx":212C9
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Height          =   495
            Left            =   4440
            Picture         =   "patientadd.frx":26B74
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   2880
            Picture         =   "patientadd.frx":2C451
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "PATIENT INFORMATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8415
         Left            =   600
         TabIndex        =   1
         Top             =   0
         Width           =   18975
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   14040
            Top             =   7080
            Visible         =   0   'False
            Width           =   3255
            _ExtentX        =   5741
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
            Connect         =   $"patientadd.frx":31CDB
            OLEDBString     =   $"patientadd.frx":31D7B
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "paitent"
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
         Begin VB.TextBox Text11 
            DataField       =   "patient discharge date"
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
            Left            =   13320
            TabIndex        =   38
            Top             =   7680
            Width           =   3615
         End
         Begin VB.TextBox Text10 
            DataField       =   "patient admit date"
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
            ForeColor       =   &H00C0C000&
            Height          =   615
            Left            =   3360
            TabIndex        =   36
            Top             =   7560
            Width           =   4575
         End
         Begin VB.TextBox Text7 
            DataField       =   "BED NUMBER"
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
            Left            =   3360
            TabIndex        =   34
            Top             =   6720
            Width           =   4575
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "gender"
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
            ItemData        =   "patientadd.frx":31E1B
            Left            =   3360
            List            =   "patientadd.frx":31E25
            TabIndex        =   23
            Top             =   1680
            Width           =   4575
         End
         Begin VB.TextBox Text1 
            DataField       =   "paitent id"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   11
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            DataField       =   "paitent name"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   10
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox Text3 
            DataField       =   "paitent disease"
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
            Left            =   3360
            TabIndex        =   9
            Top             =   2280
            Width           =   4575
         End
         Begin VB.TextBox Text4 
            DataField       =   "doctor fee"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   4320
            Width           =   4575
         End
         Begin VB.TextBox Text5 
            DataField       =   "fathers name"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   7
            Top             =   4920
            Width           =   4575
         End
         Begin VB.TextBox Text6 
            DataField       =   "phone number"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   6
            Top             =   5520
            Width           =   4575
         End
         Begin VB.TextBox Text8 
            DataField       =   "paitent department"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   5
            Top             =   3720
            Width           =   4575
         End
         Begin VB.TextBox Text9 
            DataField       =   "date of birth"
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
            Height          =   375
            Left            =   3360
            TabIndex        =   4
            Top             =   3000
            Width           =   4575
         End
         Begin VB.PictureBox Picture1 
            Height          =   6615
            Left            =   9360
            Picture         =   "patientadd.frx":31E37
            ScaleHeight     =   6555
            ScaleWidth      =   7515
            TabIndex        =   3
            Top             =   240
            Width           =   7575
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "blood group"
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
            ItemData        =   "patientadd.frx":73A50
            Left            =   3360
            List            =   "patientadd.frx":73A60
            TabIndex        =   2
            Top             =   6000
            Width           =   4575
         End
         Begin VB.Label Label4 
            Caption         =   "patient discharge date"
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
            Left            =   9360
            TabIndex        =   37
            Top             =   7680
            UseMnemonic     =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            Caption         =   "patient admit date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   35
            Top             =   7560
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "patient bed number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   33
            Top             =   6720
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
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
            Index           =   8
            Left            =   240
            TabIndex        =   21
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Id"
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
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Name"
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
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Disease"
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
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
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
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "PatientDepartment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   3720
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Fee"
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
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   4320
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
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
            Index           =   6
            Left            =   240
            TabIndex        =   14
            Top             =   4920
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
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
            Index           =   7
            Left            =   240
            TabIndex        =   13
            Top             =   5520
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Blood Group"
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
            Index           =   9
            Left            =   240
            TabIndex        =   12
            Top             =   6120
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "addmissionpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Fields("paitent id") = Text1.Text
Adodc1.Recordset.Fields("paitent name") = Text2.Text
Adodc1.Recordset.Fields("gender") = Combo2
Adodc1.Recordset.Fields("paitent disease") = Text3.Text
Adodc1.Recordset.Fields("date of birth") = Text9.Text
Adodc1.Recordset.Fields("paitent department") = Text8.Text
Adodc1.Recordset.Fields("doctor fee") = Text4.Text
Adodc1.Recordset.Fields("fathers name") = Text5.Text
Adodc1.Recordset.Fields("phone number") = Text6.Text
Adodc1.Recordset.Fields("blood group") = Combo1
Adodc1.Recordset.Fields("BED NUMBER") = Text7.Text
Adodc1.Recordset.Fields("patient admit date") = Text10.Text
Adodc1.Recordset.Fields("patient discharge date") = Text11.Text


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
Combo2.Text = " "
Text3.Text = " "
Text9.Text = " "
Text8.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Combo1.Text = " "
Text7.Text = " "
Text10.Text = " "
Text11.Text = " "
End Sub

Private Sub Command9_Click()
addmissionpatient.Hide
frmmain.Show
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text2.SetFocus

End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo2.SetFocus

End Sub
Private Sub combo2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text3.SetFocus

End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text9.SetFocus

End Sub
Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text8.SetFocus

End Sub
Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text4.SetFocus

End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text5.SetFocus

End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text6.SetFocus
End Sub
Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Combo1.SetFocus

End Sub
Private Sub combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text7.SetFocus

End Sub
Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text10.SetFocus

End Sub
Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text11.SetFocus

End Sub
Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus

End Sub
