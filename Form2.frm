VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Library Management System-Login"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16740
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   16740
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   1320
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Height          =   2295
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   13695
      Begin VB.PictureBox Picture1 
         Height          =   2175
         Left            =   0
         Picture         =   "Form2.frx":25D36
         ScaleHeight     =   2115
         ScaleWidth      =   13635
         TabIndex        =   21
         Top             =   120
         Width           =   13695
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11640
            TabIndex        =   26
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Library Management System "
            BeginProperty Font 
               Name            =   "Cooper Black"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   735
            Left            =   0
            TabIndex        =   24
            Top             =   1320
            Width           =   13455
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PRESTIGE INSTITUTE OF MANAGEMENT"
            BeginProperty Font 
               Name            =   "Cooper Black"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   735
            Left            =   0
            TabIndex        =   23
            Top             =   600
            Width           =   13455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "WELCOME"
            BeginProperty Font 
               Name            =   "Cooper Black"
               Size            =   29.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   735
            Left            =   240
            TabIndex        =   22
            Top             =   0
            Width           =   12375
         End
      End
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   14640
      TabIndex        =   18
      Top             =   7800
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   14640
      TabIndex        =   16
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   14640
      TabIndex        =   14
      Top             =   6840
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   14640
      TabIndex        =   12
      Top             =   4920
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   14640
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      Picture         =   "Form2.frx":3B904
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9360
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   2040
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "----------------------Select-------------------"
      Top             =   5760
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16800
      Picture         =   "Form2.frx":5FACB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login>>>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "Form2.frx":68E85
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6600
      Width           =   4935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   19
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   17
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   15
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "registration"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12840
      TabIndex        =   8
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Enter User ID And Password To Verify"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER LOGIN"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Width           =   6375
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim b As Integer

Private Sub Command1_Click()
If Text2.Text <> "" Then
    rs.Open "select * from staff where username='" & Combo1.Text & " ' ", con, 3, 3
    If Text2.Text = rs.Fields(5) Then
        MDIForm1.Show
        Unload Me
    Else
        MsgBox "Password is Incorrect"
        Text2.SetFocus
    End If
    rs.Close
 Else
    MsgBox "Please enter Password", vbCritical
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub command3_click()
Dim err As Integer
rs.Open "select * from staff", con, 3, 3
For i = 0 To rs.RecordCount - 1
    If Text3.Text = rs(1) Then
        MsgBox "Username Is Not Available", vbInformation
        err = 1
    End If
Next
rs.Close
If err <> 1 Then
    rs.Open "select max(empid) from staff", con, 3, 3
    dummy = rs(0)
    rs.Close
    rs.Open "select * from staff", con, 3, 3
    rs.AddNew
    rs(0) = dummy + 1
    rs(1) = Text3.Text
    rs(2) = Text6.Text
    rs(3) = Text7.Text
    rs(4) = Text4.Text
    rs(5) = Text1.Text
    rs.Update
    rs.Close
    MsgBox "Registration Successful", vbInformation
End If
End Sub

Private Sub Form_Load()
connect
rs.Open "select username from staff", con, 3, 3
rs.MoveFirst
For i = 0 To rs.RecordCount - 1
    Combo1.AddItem rs.Fields(0)
    rs.MoveNext
Next i
rs.Close
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub
Private Sub Text3_LostFocus()
rs.Open "select username from staff", con, 3, 3
rs.MoveFirst
For i = 0 To rs.RecordCount - 1
    If Text3.Text = rs(0) Then
        MsgBox "Username not available", vbInformation
    End If
    rs.MoveNext
Next
rs.Close
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Date
Label7.Caption = Time()
End Sub

Private Sub Timer2_Timer()
b = b + 1
If b = 1 Then
Label4.Caption = "WELCOME"
End If
If b = 2 Then
Label4.Caption = ""
End If


If b = 3 Then
b = 0
End If
End Sub
