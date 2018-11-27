VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsearch 
   Caption         =   "Search Information"
   ClientHeight    =   10125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10125
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   14208
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "&Book"
      TabPicture(0)   =   "Form5.frx":25D36
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataGrid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Student"
      TabPicture(1)   =   "Form5.frx":25D52
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Teacher"
      TabPicture(2)   =   "Form5.frx":25D6E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Picture3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3015
         Left            =   -74520
         TabIndex        =   8
         Top             =   4080
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5318
         _Version        =   393216
         BackColor       =   16776960
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Select The Book For More Information"
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
      Begin VB.CommandButton Command12 
         Caption         =   "<<Back"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Picture         =   "Form5.frx":25D8A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7440
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "More Info.>>"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         Picture         =   "Form5.frx":2F144
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7440
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "More Info.>>"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63480
         Picture         =   "Form5.frx":384FE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7440
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "More Info.>>"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63600
         Picture         =   "Form5.frx":418B8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7440
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "<<Back"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         Picture         =   "Form5.frx":4AC72
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7440
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "<<Back"
         BeginProperty Font 
            Name            =   "Broadway"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         Picture         =   "Form5.frx":5402C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7440
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2655
         Left            =   -74400
         TabIndex        =   9
         Top             =   4560
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4683
         _Version        =   393216
         BackColor       =   16776960
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Select The Student For More Information"
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   2655
         Left            =   480
         TabIndex        =   10
         Top             =   4440
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4683
         _Version        =   393216
         BackColor       =   16776960
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Select The Teacher For More Information"
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
      Begin VB.PictureBox Picture1 
         Height          =   7695
         Left            =   -75000
         Picture         =   "Form5.frx":5D3E6
         ScaleHeight     =   7635
         ScaleWidth      =   13755
         TabIndex        =   11
         Top             =   360
         Width           =   13815
         Begin VB.Frame Frame1 
            Caption         =   "Search By Relevance"
            BeginProperty Font 
               Name            =   "Harlow Solid Italic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   480
            TabIndex        =   12
            Top             =   720
            Width           =   7095
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
               Height          =   375
               Left            =   2760
               TabIndex        =   20
               Top             =   1800
               Width           =   3855
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
               Height          =   375
               Left            =   2760
               TabIndex        =   19
               Top             =   1200
               Width           =   3855
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
               Height          =   375
               Left            =   2760
               TabIndex        =   18
               Top             =   600
               Width           =   3855
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Accession No.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Left            =   240
               Picture         =   "Form5.frx":72FB4
               TabIndex        =   15
               Top             =   360
               Width           =   2655
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Author:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   14
               Top             =   1800
               Width           =   1935
            End
            Begin VB.OptionButton Option11 
               Caption         =   "Title:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   240
               TabIndex        =   13
               Top             =   1200
               Width           =   1935
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   7695
         Left            =   -75000
         Picture         =   "Form5.frx":98CEA
         ScaleHeight     =   7635
         ScaleWidth      =   13755
         TabIndex        =   16
         Top             =   360
         Width           =   13815
         Begin VB.Frame Frame2 
            Caption         =   "Search By Relevance"
            BeginProperty Font 
               Name            =   "Harlow Solid Italic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3615
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   6855
            Begin VB.OptionButton Option8 
               Caption         =   "ScholarNo.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   480
               TabIndex        =   29
               Top             =   1200
               Width           =   2415
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   480
               TabIndex        =   28
               Top             =   2040
               Width           =   1335
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Course:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   480
               TabIndex        =   27
               Top             =   2880
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
               Caption         =   "MemberID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   480
               TabIndex        =   26
               Top             =   360
               Width           =   2295
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
               Left            =   3000
               TabIndex        =   25
               Top             =   2160
               Width           =   3375
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
               Left            =   3000
               TabIndex        =   24
               Top             =   1320
               Width           =   3375
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
               Left            =   3000
               TabIndex        =   23
               Top             =   480
               Width           =   3375
            End
            Begin VB.ComboBox Combo4 
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
               Left            =   3000
               TabIndex        =   22
               Top             =   2880
               Width           =   3375
            End
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   7695
         Left            =   0
         Picture         =   "Form5.frx":AE8B8
         ScaleHeight     =   7635
         ScaleWidth      =   13755
         TabIndex        =   17
         Top             =   360
         Width           =   13815
         Begin VB.Frame Frame3 
            Caption         =   "Search By Relevance"
            BeginProperty Font 
               Name            =   "Harlow Solid Italic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   480
            TabIndex        =   30
            Top             =   360
            Width           =   6975
            Begin VB.OptionButton Option5 
               Caption         =   "Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   360
               TabIndex        =   38
               Top             =   1920
               Width           =   1695
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Department:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   360
               TabIndex        =   37
               Top             =   2640
               Width           =   2535
            End
            Begin VB.OptionButton Option2 
               Caption         =   "FacultyID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   360
               TabIndex        =   36
               Top             =   1200
               Width           =   2295
            End
            Begin VB.OptionButton Option1 
               Caption         =   "MemberID:"
               DisabledPicture =   "Form5.frx":C4486
               DownPicture     =   "Form5.frx":DA054
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   360
               TabIndex        =   35
               Top             =   360
               Width           =   2175
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
               Left            =   3000
               TabIndex        =   34
               Top             =   2040
               Width           =   3615
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
               Left            =   3000
               TabIndex        =   33
               Top             =   1320
               Width           =   3615
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
               Left            =   3000
               TabIndex        =   32
               Top             =   480
               Width           =   3615
            End
            Begin VB.ComboBox Combo5 
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
               Left            =   3000
               TabIndex        =   31
               Top             =   2640
               Width           =   3615
            End
         End
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH INFORMATION"
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
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Combo4_Click()
Set rs = New ADODB.Recordset
If Combo4.ListIndex <> -1 Then
    If Option6.Value = True Then
             
            rs.Open "select memberid as MemberID,scholarno as ScholarNO,stuname as Name,coursename as Course,stusex as Gender,stucontact as ContactNO,stuaddress as Address from student,course where course.courseid = " & Combo4.ItemData(Combo4.ListIndex) & " and course.courseid=student.courseid", con, 3, 3
            If rs.RecordCount <> 0 Then
                Set DataGrid2.DataSource = rs
            End If
    End If
End If
End Sub

Private Sub Combo5_Click()
Set rs = New ADODB.Recordset
If Combo5.ListIndex <> -1 Then
    If Option4.Value = True Then
             
            rs.Open "select memberid as MemberID,tid as TeacherID,tname as Name,deptname as Department,tsex as Gender,taddress as Address,tcontact as ContactNO from teacher,department where teacher.deptid = " & Combo5.ItemData(Combo5.ListIndex) & " and department.deptid=teacher.deptid", con, 3, 3
            If rs.RecordCount <> 0 Then
                Set DataGrid3.DataSource = rs
            End If
    End If
End If
End Sub

Private Sub Command10_Click()
MDIForm1.Show
Unload Me
End Sub



Private Sub Command12_Click()
MDIForm1.Show
Unload Me
End Sub


Private Sub Command4_Click()
If DataGrid1.Row <> -1 Then
    frmbook.Show
    Me.Hide
Else
    MsgBox "Please Search The Book", vbCritical
End If

End Sub

Private Sub Command5_Click()
If DataGrid2.Row <> -1 Then
    frmstud.Show
    Me.Hide
Else
    MsgBox "Please Search The Student", vbCritical
End If
End Sub

Private Sub Command6_Click()
 If DataGrid3.Row <> -1 Then
    frmteach.Show
    Me.Hide
Else
    MsgBox "Please Search The Teacher", vbCritical
End If
End Sub



Private Sub Command9_Click()
MDIForm1.Show
Unload Me
End Sub


Private Sub DataGrid1_Click()
If DataGrid1.Row <> -1 Then
    i = DataGrid1.Row
    DataGrid1.RowBookmark (i)
End If
End Sub


Private Sub DataGrid2_Click()
If DataGrid2.Row <> -1 Then
    i = DataGrid2.Row
    DataGrid2.RowBookmark (i)
End If
End Sub

Private Sub DataGrid3_Click()
If DataGrid3.Row <> -1 Then
    i = DataGrid3.Row
    DataGrid3.RowBookmark (i)
End If
End Sub

Private Sub Form_Load()
con.Close
connect
'Adding values in stud course combo
rs.Open "select courseid,coursename from course", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo4.AddItem rs.Fields(1)
    Combo4.ItemData(i - 1) = rs("courseid")
    rs.MoveNext
Next i
rs.Close
'Adding values in teach dept combo
rs.Open "select deptid,deptname from department", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo5.AddItem rs.Fields(1)
    Combo5.ItemData(i - 1) = rs(0)
    rs.MoveNext
Next i
rs.Close
End Sub

Private Sub Option1_Click()
    Text9.Locked = True
    Text10.Locked = True
    Text8.Locked = False
    Combo5.Locked = True
End Sub

Private Sub Option10_Click()
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = False
End Sub

Private Sub Option11_Click()
    Text1.Locked = True
    Text2.Locked = False
    Text3.Locked = True
End Sub

Private Sub Option2_Click()
    Text9.Locked = False
    Text10.Locked = True
    Text8.Locked = True
    Combo5.Locked = True
End Sub

Private Sub Option3_Click()
    Text5.Locked = True
    Text6.Locked = True
    Text4.Locked = False
    Combo4.Locked = True
End Sub

Private Sub Option4_Click()
    Text9.Locked = True
    Text10.Locked = True
    Text8.Locked = True
    Combo5.Locked = False
End Sub

Private Sub Option5_Click()
    Text9.Locked = True
    Text10.Locked = False
    Text8.Locked = True
    Combo5.Locked = True
End Sub

Private Sub Option6_Click()
    Text5.Locked = True
    Text6.Locked = True
    Text4.Locked = True
    Combo4.Locked = False
End Sub

Private Sub Option7_Click()
    Text5.Locked = True
    Text6.Locked = False
    Text4.Locked = True
    Combo4.Locked = True
End Sub

Private Sub Option8_Click()
    Text5.Locked = False
    Text6.Locked = True
    Text4.Locked = True
    Combo4.Locked = True
End Sub

Private Sub Option9_Click()
    Text3.Locked = True
    Text2.Locked = True
    Text1.Locked = False
End Sub

Private Sub Text1_Change()
 Set rs = New ADODB.Recordset
 If Option9.Value = True Then

        rs.Open "select bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price from book where bookid like '" & "%" & Text1.Text & "%'", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid1.DataSource = rs
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text1.Text = Text1.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text10_Change()
Set rs = New ADODB.Recordset
If Option5.Value = True Then
        rs.Open "select memberid as MemberID,tid as TeacherID,tname as Name,deptname as Department,tsex as Gender,taddress as Address,tcontact as ContactNO from teacher,department where tname like '" & "%" & Text10.Text & "%' and department.deptid=teacher.deptid", con, 3, 3
        If rs.RecordCount <> 0 Then
             Set DataGrid3.DataSource = rs
        End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text10.Text = Text10.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text2_Change()
Set rs = New ADODB.Recordset
If Option11.Value = True Then

        rs.Open "select bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price from book where bookname like '" & "%" & Text2.Text & "%'", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid1.DataSource = rs
        End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text2.Text = Text2.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text3_Change()
Set rs = New ADODB.Recordset
If Option10.Value = True Then

        rs.Open "select bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price from book where author like '" & "%" & Text3.Text & "%'", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid1.DataSource = rs
        End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text3.Text = Text3.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text4_Change()
Set rs = New ADODB.Recordset
If Option3.Value = True Then
         
        rs.Open "select memberid as MemberID,scholarno as ScholarNO,stuname as Name,coursename as Course,stusex as Gender,stucontact as ContactNO,stuaddress as Address from student,course where memberid like '" & "%" & Text4.Text & "%' and course.courseid=student.courseid", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid2.DataSource = rs
        End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text4.Text = Text4.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text5_Change()
Set rs = New ADODB.Recordset
If Option8.Value = True Then
         
        rs.Open "select memberid as MemberID,scholarno as ScholarNO,stuname as Name,coursename as Course,stusex as Gender,stucontact as ContactNO,stuaddress as Address from student,course where scholarno like '" & "%" & Text5.Text & "%' and course.courseid=student.courseid", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid2.DataSource = rs
        End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text5.Text = Text5.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text6_Change()
Set rs = New ADODB.Recordset
If Option7.Value = True Then
         
        rs.Open "select memberid as MemberID,scholarno as ScholarNO,stuname as Name,coursename as Course,stusex as Gender,stucontact as ContactNO,stuaddress as Address from student,course where stuname like '" & "%" & Text6.Text & "%' and course.courseid=student.courseid", con, 3, 3
        If rs.RecordCount <> 0 Then
            Set DataGrid2.DataSource = rs
        End If
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text6.Text = Text6.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text8_Change()
Set rs = New ADODB.Recordset
If Option1.Value = True Then
         
        rs.Open "select memberid as MemberID,tid as TeacherID,tname as Name,deptname as Department,tsex as Gender,taddress as Address,tcontact as ContactNO from teacher,department where memberid like '" & "%" & Text8.Text & "%' and department.deptid=teacher.deptid", con, 3, 3
        If rs.RecordCount <> 0 Then
             Set DataGrid3.DataSource = rs
        End If
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text8.Text = Text8.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text9_Change()
Set rs = New ADODB.Recordset
If Option2.Value = True Then
         
        rs.Open "select memberid as MemberID,tid as TeacherID,tname as Name,deptname as Department,tsex as Gender,taddress as Address,tcontact as ContactNO from teacher,department where tid like '" & "%" & Text9.Text & "%' and department.deptid=teacher.deptid", con, 3, 3
        If rs.RecordCount <> 0 Then
             Set DataGrid3.DataSource = rs
        End If
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text9.Text = Text9.Text & Chr(KeyAscii)
End If
End Sub
