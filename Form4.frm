VERSION 5.00
Begin VB.Form frmbook 
   Caption         =   "Book Information"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
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
      Picture         =   "Form4.frx":25D36
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
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
      Left            =   10440
      Picture         =   "Form4.frx":3B904
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7560
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
      Left            =   10320
      Picture         =   "Form4.frx":40524
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Left            =   10320
      Picture         =   "Form4.frx":4397E
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   10320
      Picture         =   "Form4.frx":46DD8
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Left            =   10320
      Picture         =   "Form4.frx":4A232
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4800
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4800
      TabIndex        =   0
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4800
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4800
      TabIndex        =   2
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4800
      TabIndex        =   3
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BOOK INFORMATION"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "AccessionNo.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Edition:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   5760
      Width           =   2055
   End
End
Attribute VB_Name = "frmbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()
frmsearch.Show
Unload Me
End Sub

Private Sub Command2_Click()
Text12.Locked = False
Text13.Locked = False
Text14.Locked = False
Text15.Locked = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub command3_click()
If Text12.Text <> "" And Text13.Text <> "" And Text14.Text <> "" And Text15.Text <> "" Then
    rs.Open "select * from book", con, 3, 3
    rs(1) = Text12.Text
    rs(2) = Text13.Text
    rs(3) = Text14.Text
    rs(4) = Text15.Text
    rs.Update
    rs.Close
    MsgBox "Book Record Updated Successful", vbInformation
Else
    MsgBox "Please Fill All Fields", vbCritical
End If
End Sub

Private Sub Command4_Click()
If Text11.Text <> "" Then
    sure = MsgBox("Are You Sure Want To Delete The Book Record", 1)
    If sure = 1 Then
    rs.Open "delete from book where bookid=" & Val(Text11.Text) & "", con, 3, 3
    MsgBox "Book Record Successfully Deleted", vbInformation
    frmsearch.Show
    Unload Me
    Else
        MsgBox "Book Record Not Deleted", vbInformation
    End If
Else
    MsgBox "Please Click On New Search ", vbCritical
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
rs.Open "select * from book where bookid=" & frmsearch.DataGrid1.Columns(0) & "", con, 3, 3
Text11.Text = rs(0)
Text12.Text = rs(1)
Text13.Text = rs(2)
Text14.Text = rs(3)
Text15.Text = rs(4)
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
Text14.Locked = True
Text15.Locked = True
rs.Close
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text12.Text = Text12.Text & Chr(KeyAscii)
End If
End Sub


Private Sub Text13_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text13.Text = Text13.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text14.Text = Text14.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text15.Text = Text15.Text & Chr(KeyAscii)
End If
End Sub
