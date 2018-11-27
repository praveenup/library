VERSION 5.00
Begin VB.Form frmstud 
   Caption         =   "Student Information"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      TabIndex        =   20
      Top             =   4440
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      TabIndex        =   19
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit>>"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Picture         =   "Form3.frx":25D36
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New Search>>"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Picture         =   "Form3.frx":29190
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update>>"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Picture         =   "Form3.frx":2C5EA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete>>"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Picture         =   "Form3.frx":2FA44
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close>>"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Picture         =   "Form3.frx":32E9E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<&Back"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Picture         =   "Form3.frx":48A6C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8400
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   11
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   12
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT INFORMATION"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   9
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ScholarNo.:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ContactNo.:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MemberID:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "frmstud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Dim temp As Integer


Private Sub Command1_Click()
frmsearch.Show
Unload Me
End Sub

Private Sub Command2_Click()
Text4.Locked = False
Text1.Locked = False
Text6.Locked = False
Text7.Locked = False
Combo1.Locked = False
Combo2.Locked = False
rs.Open "select courseid , coursename from course ", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo1.AddItem rs(1)
    Combo1.ItemData(i - 1) = rs(0)
    rs.MoveNext
Next
rs.Close
Combo2.AddItem "Male"
Combo2.AddItem "Female"
Command3.Enabled = True
Command4.Enabled = True

End Sub

Private Sub command3_click()
If Text4.Text <> "" And Text1.Text <> "" And Text6.Text <> "" And Text7.Text <> "" And Combo2.Text <> "" And Combo1.Text <> "" Then
    rs.Open "select * from student where memberid=" & Text5.Text & "", con, 3, 3
    rs("scholarno") = Text4.Text
    rs("stuname") = Text1.Text
    rs("stucontact") = Text6.Text
    rs("stuaddress") = Text7.Text
    rs("stusex") = Combo2.Text
    If Combo1.ListIndex <> -1 Then
        rs("courseid") = Combo1.ItemData(Combo1.ListIndex)
    End If
    rs.Update
    rs.Close
    MsgBox "Student Record Updated Successful", vbInformation
Else
    MsgBox "Please Fill All Fields", vbCritical
End If
End Sub

Private Sub Command4_Click()
If Text5.Text <> "" Then
    sure = MsgBox("Are You Sure Want To Delete The Student Record", 1)
    If sure = 1 Then
        rs.Open "delete from student where memberid=" & Val(Text5.Text) & " ", con, 3, 3
        rs.Open "delete from member where memberid=" & Val(Text5.Text) & " ", con, 3, 3
        MsgBox "Student Record Successfully Deleted", vbInformation
        frmsearch.Show
        Unload Me
    End If
Else
    MsgBox "Please Click On New Search", vbCritical
End If
End Sub

Private Sub Command5_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub Command6_Click()
frmsearch.Show
Unload Me
End Sub

Private Sub Form_Load()
Command3.Enabled = False
Command4.Enabled = False
rs.Open "select * from student where memberid=" & frmsearch.DataGrid2.Columns(0) & "", con, 3, 3
Text5.Text = rs(0)
Text4.Text = rs(1)
Text1.Text = rs(3)
Text6.Text = rs(5)
Text7.Text = rs(6)
Combo2.Text = rs(4)
temp = rs("courseid")
rs.Close
rs.Open "select coursename from course where courseid=" & temp & "", con, 3, 3
Combo1.Text = rs(0)
rs.Close
Text5.Locked = True
Text4.Locked = True
Text1.Locked = True
Text6.Locked = True
Text7.Locked = True
Combo1.Locked = True
Combo2.Locked = True
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text1.Text = Text1.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text4.Text = Text4.Text & Chr(KeyAscii)
End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(Text6.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            Text6.Text = Text6.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        Text6.Text = Text6.Text & Chr(KeyAscii)
    End If
End If
End Sub
