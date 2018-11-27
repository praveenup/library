VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmissue 
   Caption         =   "BOOK ISSUE"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form8"
   ScaleHeight     =   9045
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   5520
      TabIndex        =   8
      Top             =   3960
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   41746434
      CurrentDate     =   42270
   End
   Begin VB.TextBox txtStudentID 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtBookID 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6960
      TabIndex        =   0
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "MemberID:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "AccessionNo:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "Issue Date:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   3960
      Width           =   2775
   End
End
Attribute VB_Name = "frmissue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
