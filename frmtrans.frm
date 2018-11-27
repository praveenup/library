VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmtrans 
   Caption         =   "Transactions"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmtrans.frx":0000
   ScaleHeight     =   15756.42
   ScaleMode       =   0  'User
   ScaleWidth      =   1.09523e5
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11655
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   20558
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483630
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Issue"
      TabPicture(0)   =   "frmtrans.frx":25D36
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdrefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdissue"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdclose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&Return"
      TabPicture(1)   =   "frmtrans.frx":25D52
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmddeposit"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Member Information"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   480
         TabIndex        =   24
         Top             =   5760
         Width           =   17655
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
            Left            =   2280
            TabIndex        =   29
            Top             =   1320
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
            TabIndex        =   28
            Top             =   600
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
            TabIndex        =   27
            Top             =   3480
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
            Left            =   2280
            TabIndex        =   26
            Top             =   2040
            Width           =   3015
         End
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
            TabIndex        =   25
            Top             =   2760
            Width           =   3015
         End
         Begin VB.PictureBox Picture4 
            Height          =   4575
            Left            =   0
            Picture         =   "frmtrans.frx":25D6E
            ScaleHeight     =   4515
            ScaleWidth      =   17595
            TabIndex        =   30
            Top             =   240
            Width           =   17655
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   3615
               Left            =   6000
               TabIndex        =   38
               Top             =   360
               Width           =   10695
               _ExtentX        =   18865
               _ExtentY        =   6376
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
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Issued Books"
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
               Left            =   360
               TabIndex        =   37
               Top             =   1800
               Width           =   1935
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
               Left            =   360
               TabIndex        =   36
               Top             =   2520
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
               TabIndex        =   35
               Top             =   3240
               Width           =   1695
            End
            Begin VB.Label Label3 
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
               Top             =   3840
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
               TabIndex        =   33
               Top             =   3840
               Width           =   1095
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
               Left            =   360
               TabIndex        =   32
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label1 
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
               Left            =   360
               TabIndex        =   31
               Top             =   1080
               Width           =   1695
            End
         End
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
         Height          =   495
         Left            =   -72120
         TabIndex        =   14
         Top             =   2160
         Width           =   3735
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
         Height          =   495
         Left            =   -72120
         TabIndex        =   13
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69000
         Picture         =   "frmtrans.frx":3B93C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -74040
         Picture         =   "frmtrans.frx":3ED96
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4800
         Width           =   2295
      End
      Begin VB.CommandButton cmddeposit 
         Caption         =   "Deposit "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -71520
         Picture         =   "frmtrans.frx":421F0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Caption         =   "Search Book"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   8280
         TabIndex        =   8
         Top             =   600
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2775
            Left            =   360
            TabIndex        =   9
            Top             =   1920
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   4895
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
            Height          =   495
            Left            =   2160
            TabIndex        =   0
            Top             =   960
            Width           =   3135
         End
         Begin VB.PictureBox Picture3 
            Height          =   4695
            Left            =   0
            Picture         =   "frmtrans.frx":4564A
            ScaleHeight     =   4635
            ScaleWidth      =   9795
            TabIndex        =   21
            Top             =   240
            Width           =   9855
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Book Name:"
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
               Left            =   480
               TabIndex        =   22
               Top             =   720
               Width           =   1575
            End
         End
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
         Height          =   525
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   3255
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
         Height          =   525
         Left            =   2760
         TabIndex        =   2
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   6000
         Picture         =   "frmtrans.frx":4EA04
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton cmdissue 
         Caption         =   "Issue "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3360
         Picture         =   "frmtrans.frx":51E5E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh "
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   600
         Picture         =   "frmtrans.frx":552B8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4320
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         Height          =   11295
         Left            =   0
         Picture         =   "frmtrans.frx":58712
         ScaleHeight     =   11235
         ScaleWidth      =   18915
         TabIndex        =   15
         Top             =   360
         Width           =   18975
         Begin VB.CommandButton Command4 
            Caption         =   "&Show Info."
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            Picture         =   "frmtrans.frx":6E2E0
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   480
            Width           =   1695
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
            Height          =   525
            Left            =   2760
            TabIndex        =   4
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   615
            Left            =   720
            TabIndex        =   39
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
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
            Left            =   360
            TabIndex        =   17
            Top             =   1440
            Width           =   3375
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
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   615
            Left            =   720
            TabIndex        =   16
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   11295
         Left            =   -75000
         Picture         =   "frmtrans.frx":72F00
         ScaleHeight     =   11235
         ScaleWidth      =   18915
         TabIndex        =   18
         Top             =   360
         Width           =   18975
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            Caption         =   "Search Member"
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   8640
            TabIndex        =   60
            Top             =   600
            Width           =   9615
            Begin VB.TextBox Text19 
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
               Left            =   2400
               TabIndex        =   62
               Top             =   720
               Width           =   3135
            End
            Begin MSDataGridLib.DataGrid DataGrid4 
               Height          =   2655
               Left            =   360
               TabIndex        =   61
               Top             =   1560
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   4683
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
            Begin VB.PictureBox Picture6 
               Height          =   4335
               Left            =   0
               Picture         =   "frmtrans.frx":88ACE
               ScaleHeight     =   4275
               ScaleWidth      =   9555
               TabIndex        =   63
               Top             =   240
               Width           =   9615
               Begin VB.Label Label24 
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
                  Height          =   495
                  Left            =   480
                  TabIndex        =   64
                  Top             =   480
                  Width           =   1815
               End
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Show Books"
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            Picture         =   "frmtrans.frx":91E88
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2880
            TabIndex        =   57
            Top             =   2640
            Width           =   3735
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2880
            TabIndex        =   56
            Top             =   3480
            Width           =   3735
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C000&
            Caption         =   "Member Information"
            BeginProperty Font 
               Name            =   "Broadway"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4815
            Left            =   480
            TabIndex        =   40
            Top             =   5520
            Width           =   17655
            Begin VB.TextBox Text16 
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
               TabIndex        =   45
               Top             =   2760
               Width           =   3015
            End
            Begin VB.TextBox Text14 
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
               TabIndex        =   44
               Top             =   2040
               Width           =   3015
            End
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
               TabIndex        =   43
               Top             =   3480
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
               Left            =   2280
               TabIndex        =   42
               Top             =   600
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
               TabIndex        =   41
               Top             =   1320
               Width           =   3015
            End
            Begin VB.PictureBox Picture5 
               Height          =   4575
               Left            =   0
               Picture         =   "frmtrans.frx":96AA8
               ScaleHeight     =   4515
               ScaleWidth      =   17595
               TabIndex        =   46
               Top             =   240
               Width           =   17655
               Begin MSDataGridLib.DataGrid DataGrid3 
                  Height          =   3615
                  Left            =   6360
                  TabIndex        =   47
                  Top             =   240
                  Width           =   10695
                  _ExtentX        =   18865
                  _ExtentY        =   6376
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
                  Caption         =   "Issued Books"
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
               Begin VB.Label Label18 
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
                  Left            =   360
                  TabIndex        =   54
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.Label Label17 
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
                  Left            =   360
                  TabIndex        =   53
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.Label Label15 
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
                  TabIndex        =   52
                  Top             =   3720
                  Width           =   1095
               End
               Begin VB.Label Label12 
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
                  TabIndex        =   51
                  Top             =   3720
                  Width           =   1695
               End
               Begin VB.Label Label11 
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
                  TabIndex        =   50
                  Top             =   3240
                  Width           =   1695
               End
               Begin VB.Label Label10 
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
                  Left            =   360
                  TabIndex        =   49
                  Top             =   2520
                  Width           =   1695
               End
               Begin VB.Label Label9 
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
                  Left            =   360
                  TabIndex        =   48
                  Top             =   1800
                  Width           =   1935
               End
            End
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   615
            Left            =   600
            TabIndex        =   58
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Return Date:-"
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
            Height          =   615
            Left            =   360
            TabIndex        =   55
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
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
            Height          =   735
            Left            =   240
            TabIndex        =   20
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
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
            Height          =   735
            Left            =   600
            TabIndex        =   19
            Top             =   960
            Width           =   2775
         End
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSACTIONS"
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
      Left            =   8280
      TabIndex        =   23
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cmd_deposit As ADODB.Command
Dim check As Boolean 'already issue or not
Dim check1 As Boolean ''already deposited or not
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddeposit_Click()
If check1 = False Then
    Set rs = New ADODB.Recordset
    If Text4.Text <> "" And Text5.Text <> "" Then
        rs.Open "select mtype_id from member where memberid=" & Val(Text4.Text) & "", con, 3, 3
            If rs.RecordCount > 0 Then
                mtypeid = rs(0)
                rs.Close
                
                rs.Open "select rtype_id from book where bookid=" & Val(Text5.Text) & " and issueflag=" & True & "", con, 3, 3
                    If rs.RecordCount > 0 Then
                        rtypeid = rs(0)
                        rs.Close
                        
                        rs.Open "select duration,fine_amt from membersetting where mtype_id=" & mtypeid & " and rtype_id=" & rtypeid & "", con, 3, 3
                        If rs.RecordCount > 0 Then
                                Duration = rs(0)
                                fine = rs(1)
                            rs.Close
                            rs.Open "select issuedate from issue where bookid=" & Val(Text5.Text) & " and member_id=" & Val(Text4.Text) & " and return_date is null", con, 3, 3
                            If rs.RecordCount <> 0 Then
                                    i_date = rs(0)    'storing issue date
                                rs.Close
                                If DateDiff("d", i_date, Date) <= Val(Duration) Then
                                    Set cmd_deposit = New ADODB.Command
                                    cmd_deposit.CommandType = adCmdText
                                    cmd_deposit.ActiveConnection = con
                                    rs.Open "select * from issue where member_id=" & Val(Text4.Text) & " and bookid=" & Val(Text5.Text) & " and return_date is null", con, 3, 3
                                    If rs.RecordCount <> 0 Then
                                        cmd_deposit.CommandText = "update issue set fine=0,return_date='" & Date & "' where member_id=" & Val(Text4.Text) & " and bookid=" & Val(Text5.Text) & ""
                                        cmd_deposit.Execute
                                        cmd_deposit.CommandText = "update book set issueflag=" & False & " where bookid=" & Val(Text5.Text) & ""
                                        cmd_deposit.Execute
                                        rs.Close
                                        check1 = True
                                        MsgBox "Book Deposited", vbInformation
                                        rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text4.Text) & " And  issue.bookid  = book.bookid  and issueflag=" & True & " and return_date is null", con, 3, 3
                                        If rs.RecordCount > 0 Then
                                            Label15.Caption = rs.RecordCount
                                            Set DataGrid3.DataSource = rs
                                        End If
                                        Set rs = New ADODB.Recordset
                                        rs.CursorLocation = adUseClient
                                        rs.Open "select book.bookid as BookID,member_id as MemberID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id like '" & "%" & Text19.Text & "%' And  issue.bookid  = book.bookid  and return_date is null ", con, 3, 3
                                        If rs.RecordCount > 0 Then
                                            Set DataGrid4.DataSource = rs
                                        End If
    
                                    Else
                                        rs.Close
                                        MsgBox "Book is Not Issued", vbCritical
                                    End If
                                Else
                                    'fine
                                    f = DateDiff("d", DateAdd("d", Duration, i_date), Date) * fine
                                    sure = MsgBox("Fine=" & f & "   Are you want to Deposit", 1)
                                    If sure = 1 Then
                                        Set cmd_deposit = New ADODB.Command
                                        cmd_deposit.CommandType = adCmdText
                                        cmd_deposit.ActiveConnection = con
                                        rs.Open "select * from issue where member_id=" & Val(Text4.Text) & " and bookid=" & Val(Text5.Text) & "and return_date is null", con, 3, 3
                                        If rs.RecordCount <> 0 Then
                                            cmd_deposit.CommandText = "update issue set fine=" & f & ",return_date='" & Date & "' where member_id=" & Val(Text4.Text) & " and bookid=" & Val(Text5.Text) & ""
                                            cmd_deposit.Execute
                                            cmd_deposit.CommandText = "update book set issueflag=" & False & " where bookid=" & Val(Text5.Text) & ""
                                            cmd_deposit.Execute
                                            rs.Close
                                            check1 = True
                                            MsgBox "Book Deposited", vbInformation
                                            rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text4.Text) & " And  issue.bookid  = book.bookid  and issueflag=" & True & " and return_date is null", con, 3, 3
                                            If rs.RecordCount > 0 Then
                                                Label15.Caption = rs.RecordCount
                                                Set DataGrid3.DataSource = rs
                                            End If
                                            Set rs = New ADODB.Recordset
                                            rs.CursorLocation = adUseClient
                                            rs.Open "select book.bookid as BookID,member_id as MemberID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id like '" & "%" & Text19.Text & "%' And  issue.bookid  = book.bookid  and return_date is null ", con, 3, 3
                                            If rs.RecordCount > 0 Then
                                                Set DataGrid4.DataSource = rs
                                            End If
                                        Else
                                            rs.Close
                                            MsgBox "Book is Not Issued", vbCritical
                                        End If
                                    Else
                                        MsgBox "Book not  Deposited", vbInformation
                                    End If
                                End If
                            Else
                                MsgBox "Book is Not Issued", vbCritical
                            End If
                        Else
                            MsgBox "Member Settings Not Exists", vbCritical
                        End If
                    Else
                        MsgBox "Book Doesn't Exists or Not Issued To Anyone", vbCritical
                        rs.Close
                    End If
            Else
                MsgBox "Member Doesn't Exists", vbCritical
                rs.Close
            End If
    Else
        MsgBox "Please Fill Above all Fields", vbCritical
    End If
Else
    MsgBox "Book Already deposited", vbCritical
End If
End Sub

Private Sub cmdissue_Click()
Set rs = New ADODB.Recordset
If Text1.Text <> "" And Text3.Text <> "" Then
    rs.Open "select mtype_id from member where memberid=" & Val(Text3.Text) & "", con, 3, 3
        If rs.RecordCount > 0 Then
            mtypeid = rs(0)
            rs.Close
            
            rs.Open "select rtype_id from book where bookid=" & Val(Text1.Text) & " and issueflag=" & False & "", con, 3, 3
                If rs.RecordCount > 0 Then
                    rtypeid = rs(0)
                    rs.Close
                    
                    rs.Open "select qty from membersetting where mtype_id=" & mtypeid & " and rtype_id=" & rtypeid & "", con, 3, 3
                    If rs.RecordCount > 0 Then
                            quantity = rs(0)
                        rs.Close
                        
                        rs.Open "select * from issue,book where book.bookid=issue.bookid and rtype_id=" & rtypeid & " and member_id=" & Text3.Text & " and issueflag=" & True & " and return_date is null", con, 3, 3
                            Cont = rs.RecordCount
                        rs.Close
                        If check = False Then
                            If Val(quantity) > Cont Then
                                Set rs = New ADODB.Recordset
                                rs.Open "select * from book where bookid=" & Val(Text1.Text) & "", con, 3, 3
                                    rs("issueflag") = True
                                    rs.Update
                                rs.Close
                                rs.Open "select max(issueid) from issue", con, 3, 3
                                    temp = rs(0)
                                rs.Close
                                rs.Open "select * from issue", con, 3, 3
                                    rs.AddNew
                                    rs(0) = temp + 1
                                    rs(1) = Text3.Text
                                    rs(2) = CDate(Date)
                                    rs(3) = Text1.Text
                                    rs.Update
                                rs.Close
                                check = True
                                rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text3.Text) & " And  issue.bookid  = book.bookid  and issueflag=" & True & " and return_date is null", con, 3, 3
                                    If rs.RecordCount > 0 Then
                                        Label16.Caption = rs.RecordCount
                                        Set DataGrid2.DataSource = rs
                                    End If
                                Set rs = New ADODB.Recordset
                                    rs.CursorLocation = adUseClient
                                    rs.Open "select * from book where  issueflag=" & False & "", con, 3, 3
                                    If rs.RecordCount > 0 Then
                                        Set DataGrid1.DataSource = rs
                                    End If
                                
                                MsgBox "Book Issued Successful", vbInformation
                            
                            Else
                                MsgBox "Only Limited Number of Books can be Issued", vbCritical
                            End If
                        Else
                            MsgBox "Book Already Issued", vbCritical
                        End If
                    Else
                        MsgBox "Member Settings Not Exists", vbCritical
                    End If
                Else
                    MsgBox "Book Doesn't Exists or Already Issue to Someone", vbCritical
                    rs.Close
                End If
        Else
            MsgBox "Member Doesn't Exists", vbCritical
            rs.Close
        End If
Else
    MsgBox "Please Fill Above all Fields", vbCritical
End If
End Sub

Private Sub cmdrefresh_Click()
Text1.Text = ""
Text3.Text = ""
Text6.Text = ""
Text7.Text = ""
Text12.Text = ""
Text15.Text = ""
Text13.Text = ""
Label16.Caption = ""
DataGrid2.ClearFields
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
End Sub

Private Sub Command1_Click()
Text4.Text = ""
Text5.Text = ""
Text10.Text = ""
Text9.Text = ""
Text11.Text = ""
Text16.Text = ""
Text14.Text = ""
Text18.Text = ""
Label15.Caption = ""
Set DataGrid4.DataSource = Nothing
Set DataGrid3.DataSource = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
Set rs = New ADODB.Recordset
If DataGrid1.Row <> -1 Then
    i = DataGrid1.Row
    DataGrid1.RowBookmark (i)
    Text1.Text = DataGrid1.Columns(0)
    check = False
End If
End Sub

'Private Sub DataGrid3_Click()
'Set rs = New ADODB.Recordset
'If DataGrid3.Row <> -1 Then
'    i = DataGrid3.Row
'    DataGrid3.RowBookmark (i)
'    Text5.Text = DataGrid3.Columns(1)
'    Text18.Text = DataGrid3.Columns(6)
'End If
'End Sub

Private Sub DataGrid4_Click()
Set rs = New ADODB.Recordset
If DataGrid4.Row <> -1 Then
    i = DataGrid4.Row
    DataGrid4.RowBookmark (i)
    Text4.Text = DataGrid4.Columns(1)
    Text5.Text = DataGrid4.Columns(0)
    Text18.Text = DataGrid4.Columns(6)
    check1 = False
End If
End Sub

Private Sub Form_Load()
con.Close
connect
DataGrid2.Enabled = False
Label16.Caption = ""
Label15.Caption = ""
Text8.Text = Date
Text8.Locked = True
Text17.Text = Date
Text17.Locked = True
End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text1.Text = Text1.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text1_LostFocus()
Set rs = New ADODB.Recordset
If Text1.Text <> "" Then
    rs.Open "select * from book where bookid=" & Val(Text1.Text) & "", con, 3, 3
    If rs.RecordCount = 0 Then
        MsgBox "BookID is Incorrect", vbCritical
    Else
        check = False
    End If
End If
End Sub

Private Sub Text19_Change()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select book.bookid as BookID,member_id as MemberID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id like '" & "%" & Text19.Text & "%' And  issue.bookid  = book.bookid  and return_date is null ", con, 3, 3
If rs.RecordCount > 0 Then
    Set DataGrid4.DataSource = rs
End If
End Sub

Private Sub Text2_Change()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price from book where bookname like '" & "%" & Text2.Text & "%' and issueflag=" & False & "", con, 3, 3
If rs.RecordCount > 0 Then
    Set DataGrid1.DataSource = rs
End If
End Sub

Private Sub Command4_Click()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
If Text3.Text <> "" Then
    rs.Open "select * from issue where member_id=" & Val(Text3.Text) & " and return_date is null ", con, 3, 3
    If rs.RecordCount > 0 Then
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
        
    rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text3.Text) & " And  issue.bookid  = book.bookid  and return_date is null ", con, 3, 3
    'If rs.RecordCount > 0 Then
        Set DataGrid2.DataSource = rs
    'End If
    Else
        MsgBox "No Books are Issued ", vbCritical
    End If
Else
    MsgBox "Sorry,MemberID Field is not Fill", vbCritical
End If
End Sub


Private Sub command3_click()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
If Text4.Text <> "" Then
    rs.Open "select * from issue where member_id=" & Val(Text4.Text) & " and return_date is null ", con, 3, 3
        If rs.RecordCount > 0 Then
            Label15.Caption = rs.RecordCount
            rs.Close
            'for taking membertypeid
            rs.Open "select mtype_id from member where memberid=" & Val(Text4.Text) & "", con, 3, 3
            temp = rs(0)
            rs.Close
            Select Case (temp)
                
                Case 1:
                        rs.Open "select * from student where memberid = " & Val(Text4.Text) & "", con, 3, 3
                        Text10.Text = "Student"
                        Text9.Text = rs(3)
                            Label9.Caption = "Course:"
                            c_id = rs(2)
                            Text11.Text = rs(5)
                            Text16.Text = rs(6)
                            rs.Close
                            rs.Open "select coursename from course where courseid=" & c_id & "", con, 3, 3
                            Text14.Text = rs(0)
                            rs.Close
                Case 2:
                        rs.Open "select * from teacher where memberid = " & Val(Text4.Text) & "", con, 3, 3
                        Text10.Text = "Teacher"
                        Text9.Text = rs(3)
                        Label9.Caption = "Department:"
                        d_id = rs(2)
                        Text11.Text = rs(6)
                        Text16.Text = rs(5)
                        rs.Close
                        rs.Open "select deptname from department where deptid=" & d_id & "", con, 3, 3
                        Text14.Text = rs(0)
                        rs.Close
            
            End Select
            
            rs.Open "select issueid as IssueID,book.bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price,issuedate as Issue_Date from book,issue where member_id=" & Val(Text4.Text) & " And  issue.bookid  = book.bookid  and return_date is null ", con, 3, 3
            'If rs.RecordCount > 0 Then
                Set DataGrid3.DataSource = rs
            'End If
        Else
            MsgBox "No Books are Issued ", vbCritical
        End If
Else
    MsgBox "Sorry,MemberID Field is not Fill", vbCritical
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text2.Text = Text2.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text3.Text = Text3.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text4.Text = Text4.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text5.Text = Text5.Text & Chr(KeyAscii)
End If
End Sub
