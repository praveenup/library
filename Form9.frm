VERSION 5.00
Begin VB.Form frmdeposit 
   Caption         =   "Book Deposit"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   9255
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      TabIndex        =   9
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      Left            =   2280
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox text1 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmddeposit 
      Caption         =   "Deposit "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3000
      TabIndex        =   2
      Top             =   4440
      Width           =   2415
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
      Height          =   705
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
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
      Height          =   765
      Left            =   5880
      TabIndex        =   0
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Fine:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "ScholarNo:-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "frmdeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
