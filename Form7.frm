VERSION 5.00
Begin VB.Form frmteach 
   Caption         =   "Teacher Information"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   9840
   ScaleWidth      =   13110
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
      Left            =   3000
      TabIndex        =   20
      Top             =   3840
      Width           =   2775
   End
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
      Left            =   3000
      TabIndex        =   19
      Top             =   4560
      Width           =   2775
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
      Left            =   960
      Picture         =   "Form7.frx":25D36
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8640
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
      Left            =   9960
      Picture         =   "Form7.frx":3B904
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8640
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
      Left            =   9840
      Picture         =   "Form7.frx":514D2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
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
      Left            =   9840
      Picture         =   "Form7.frx":5492C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
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
      Left            =   9840
      Picture         =   "Form7.frx":57D86
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
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
      Left            =   9840
      Picture         =   "Form7.frx":5B1E0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
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
      Height          =   1215
      Left            =   3000
      TabIndex        =   12
      Top             =   6000
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
      Left            =   3000
      TabIndex        =   10
      Top             =   1560
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
      Height          =   450
      Left            =   3000
      TabIndex        =   8
      Top             =   5280
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
      Height          =   450
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   4560
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
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   2775
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
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Dept.:-"
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
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TeacherID.:-"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   2
      Top             =   6000
      Width           =   3135
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
      Left            =   1080
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER INFORMATION"
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
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmteach"
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
Text2.Locked = False
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
Set rs = New ADODB.Recordset
If Text4.Text <> "" And Text2.Text <> "" And Text6.Text <> "" And Text7.Text <> "" And Combo1.Text <> "" And Combo2.Text <> "" Then
    rs.Open "select * from teacher where memberid=" & Text5.Text & "", con, 3, 3
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
rs.Open "select * from teacher where memberid=" & frmsearch.DataGrid3.Columns(0) & "", con, 3, 3
Text5.Text = rs(1)
Text4.Text = rs(0)
Text2.Text = rs(3)
Text6.Text = rs(5)
Text7.Text = rs(6)
Combo2.Text = rs(4)
temp = rs("deptid")
rs.Close
rs.Open "select coursename from course where courseid=" & temp & "", con, 3, 3
Combo1.Text = rs(0)
rs.Close
Text5.Locked = True
Text4.Locked = True
Text2.Locked = True
Text6.Locked = True
Text7.Locked = True
Combo1.Locked = True
Combo2.Locked = True
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text2.Text = Text2.Text & Chr(KeyAscii)
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
