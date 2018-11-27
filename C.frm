VERSION 5.00
Begin VB.Form frmissue 
   Caption         =   "BOOK ISSUE"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   12675
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtStudentID 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtBookID 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   600
      TabIndex        =   2
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdissue 
      Caption         =   "Issue "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   1
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6960
      TabIndex        =   0
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "MemberID:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "AccessionNo:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "Issue Date:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
End
Attribute VB_Name = "frmissue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = Date

End Sub

