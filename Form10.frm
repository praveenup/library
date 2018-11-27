VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmstatus 
   Caption         =   "Status"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
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
      Left            =   18000
      Picture         =   "Form10.frx":25D36
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   10920
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Member"
      TabPicture(0)   =   "Form10.frx":2F0F0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Books"
      TabPicture(1)   =   "Form10.frx":2F10C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Member Information"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -72720
         TabIndex        =   12
         Top             =   3960
         Width           =   11415
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   18
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   17
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   15
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7920
            TabIndex        =   14
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   960
            Width           =   3015
         End
         Begin VB.PictureBox Picture4 
            Height          =   3735
            Left            =   0
            Picture         =   "Form10.frx":2F128
            ScaleHeight     =   3675
            ScaleWidth      =   11355
            TabIndex        =   38
            Top             =   240
            Width           =   11415
            Begin VB.Line Line2 
               X1              =   5760
               X2              =   5760
               Y1              =   240
               Y2              =   3480
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Member Type:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   240
               TabIndex        =   44
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6120
               TabIndex        =   43
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Course:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6120
               TabIndex        =   42
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "MemberID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   240
               TabIndex        =   41
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6120
               TabIndex        =   40
               Top             =   2040
               Width           =   1695
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   240
               TabIndex        =   39
               Top             =   2160
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Caption         =   "Member Information"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   960
         TabIndex        =   9
         Top             =   2640
         Width           =   11415
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   20
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   10
            Top             =   480
            Width           =   3015
         End
         Begin VB.PictureBox Picture3 
            Height          =   2895
            Left            =   0
            Picture         =   "Form10.frx":384E2
            ScaleHeight     =   2835
            ScaleWidth      =   11355
            TabIndex        =   30
            Top             =   240
            Width           =   11415
            Begin VB.Line Line1 
               X1              =   5880
               X2              =   5880
               Y1              =   240
               Y2              =   2640
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6600
               TabIndex        =   37
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Member Type:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   240
               TabIndex        =   36
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label16 
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
               Height          =   375
               Left            =   2280
               TabIndex        =   35
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Book Issued:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   360
               TabIndex        =   34
               Top             =   2280
               Width           =   1695
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   360
               TabIndex        =   33
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   840
               TabIndex        =   32
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Course:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6600
               TabIndex        =   31
               Top             =   960
               Width           =   1695
            End
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72240
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69720
         Picture         =   "Form10.frx":4189C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72240
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search>>"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Picture         =   "Form10.frx":4557A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Show Books"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12720
         Picture         =   "Form10.frx":49258
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5400
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Height          =   9975
         Left            =   0
         Picture         =   "Form10.frx":4DE78
         ScaleHeight     =   9915
         ScaleWidth      =   15075
         TabIndex        =   22
         Top             =   360
         Width           =   15135
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3375
            Left            =   960
            TabIndex        =   45
            Top             =   5880
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   5953
            _Version        =   393216
            BackColor       =   16776960
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Books Issued"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "MemberID:-"
            BeginProperty Font 
               Name            =   "Cambria Math"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Status For Member "
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4680
            TabIndex        =   23
            Top             =   480
            Width           =   4935
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   9975
         Left            =   -75000
         Picture         =   "Form10.frx":63A46
         ScaleHeight     =   9915
         ScaleWidth      =   15075
         TabIndex        =   25
         Top             =   360
         Width           =   15135
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Status For Books"
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5160
            TabIndex        =   29
            Top             =   480
            Width           =   5895
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "AccessionNo.:-"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   480
            TabIndex        =   28
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Issued By:-"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   840
            TabIndex        =   27
            Top             =   3240
            Width           =   3375
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Book Name:-"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   720
            TabIndex        =   26
            Top             =   2400
            Width           =   3375
         End
      End
      Begin VB.Label Label12 
         Height          =   375
         Left            =   -72240
         TabIndex        =   5
         Top             =   3600
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim temp As Integer
Dim temp1 As Integer
Dim c_id As Integer
Dim d_id As Integer



Private Sub Command1_Click()
Set rs = New ADODB.Recordset
If Text1.Text <> "" Then
    rs.Open "select * from issue where bookid=" & Val(Text1.Text) & " and return_date is null", con, 3, 3
    If rs.RecordCount <> 0 Then
        temp1 = rs(1)
        rs.Close
        rs.Open "select * from book where bookid=" & Val(Text1.Text) & "", con, 3, 3
        Text2.Text = rs(1)
        rs.Close
        rs.Open "select mtype_id from member where memberid=" & temp1 & "", con, 3, 3
        temp = rs(0)
        rs.Close
        Select Case (temp)
            
            Case 1:
                    rs.Open "select * from student where memberid = " & temp1 & "", con, 3, 3
                    Text4.Text = "Student"
                    Text5.Text = rs(3)
                    Text9.Text = temp1
                    Label7.Caption = "Course:"
                    c_id = rs(2)
                    Text11.Text = rs(5)
                    Text8.Text = rs(6)
                    rs.Close
                    rs.Open "select coursename from course where courseid=" & c_id & "", con, 3, 3
                    Text10.Text = rs(0)
                    rs.Close
            Case 2:
                    rs.Open "select * from teacher where memberid = " & temp1 & "", con, 3, 3
                    Text4.Text = "Teacher"
                    Text5.Text = rs(3)
                    Text9.Text = temp1
                    Label7.Caption = "Department:"
                    d_id = rs(2)
                    Text11.Text = rs(6)
                    Text8.Text = rs(5)
                    rs.Close
                    rs.Open "select deptname from department where deptid=" & d_id & "", con, 3, 3
                    Text10.Text = rs(0)
                    rs.Close
        
        End Select
    Else
       
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text2.Text = ""
        MsgBox "Book Not Issued To Anyone", vbInformation
        
    End If
Else
    MsgBox "Please enter AccessionNO.", vbCritical
End If
End Sub

Private Sub Command2_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub command3_click()
If Text3.Text <> "" Then
    Set rs = New ADODB.Recordset
    rs.Open "select * from issue where member_id=" & Val(Text3.Text) & " and return_date is null ", con, 3, 3
    If rs.RecordCount <> 0 Then
        Label16.Caption = rs.RecordCount
        rs.Close
        'for taking membertypeid
        rs.Open "select mtype_id from member where memberid=" & Val(Text3.Text) & "", con, 3, 3
        temp = rs(0)
        rs.Close
        Select Case (temp)
            
            Case 1:
                    rs.Open "select * from student where memberid = " & Val(Text3.Text) & "", con, 3, 3
                    Text7.Text = "Student"
                    Text6.Text = rs(3)
                        Label22.Caption = "Course:"
                        c_id = rs(2)
                        Text12.Text = rs(5)
                        Text15.Text = rs(6)
                        rs.Close
                        rs.Open "select coursename from course where courseid=" & c_id & "", con, 3, 3
                        Text13.Text = rs(0)
                    rs.Close
            Case 2:
                    rs.Open "select * from teacher where memberid = " & Val(Text3.Text) & "", con, 3, 3
                    Text7.Text = "Teacher"
                    Text6.Text = rs(3)
                    Label22.Caption = "Department:"
                    d_id = rs(2)
                    Text12.Text = rs(6)
                    Text15.Text = rs(5)
                    rs.Close
                    rs.Open "select deptname from department where deptid=" & d_id & "", con, 3, 3
                    Text13.Text = rs(0)
                    rs.Close
        
        End Select
    
    Else
        rs.Close
        Text6.Text = ""
        Text7.Text = ""
        Text12.Text = ""
        Text15.Text = ""
        Text13.Text = ""
        Text3.Text = ""
        Label16.Caption = ""
        MsgBox "Member Doesn't Exists", vbCritical
    End If
    Set DataGrid1.DataSource = Nothing
Else
    MsgBox "Please Enter MemberID", vbCritical
End If
End Sub

Private Sub Command4_Click()

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text3.Text) & " And  issue.bookid  = book.bookid  and return_date is null", con, 3, 3
If rs.RecordCount > 0 Then
    Label16.Caption = rs.RecordCount
    Set DataGrid1.DataSource = rs
End If
End Sub



Private Sub Form_Load()
con.Close
connect
Text2.Locked = True
Text4.Locked = True
Text5.Locked = True
Text7.Locked = True
Text6.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text1.Text = Text1.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text3.Text = Text3.Text & Chr(KeyAscii)
End If
End Sub
