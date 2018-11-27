VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmadd 
   Caption         =   "Add_New_Record"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16920
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   16920
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   18360
      Picture         =   "Form6.frx":25D36
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10680
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   840
      TabIndex        =   8
      Top             =   720
      Width           =   19125
      _ExtentX        =   33734
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Student"
      TabPicture(0)   =   "Form6.frx":2F0F0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_delete_stu"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_addnew_stu"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmd_add_stu"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Teacher"
      TabPicture(1)   =   "Form6.frx":2F10C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmd_delete_teach"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmd_addnew_teach"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmd_add_teach"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text3"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Combo2"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text6"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Combo3"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Book"
      TabPicture(2)   =   "Form6.frx":2F128
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image3"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label19"
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(5)=   "Label23"
      Tab(2).Control(6)=   "Label24"
      Tab(2).Control(7)=   "cmd_delete_book"
      Tab(2).Control(8)=   "Frame3"
      Tab(2).Control(9)=   "cmd_addnew_book"
      Tab(2).Control(10)=   "Text12"
      Tab(2).Control(11)=   "Text13"
      Tab(2).Control(12)=   "Text14"
      Tab(2).Control(13)=   "Text15"
      Tab(2).Control(14)=   "cmd_add_book"
      Tab(2).Control(15)=   "Combo5"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Member Setting"
      TabPicture(3)   =   "Form6.frx":2F144
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image1"
      Tab(3).Control(1)=   "Label30"
      Tab(3).Control(2)=   "Label29"
      Tab(3).Control(3)=   "Label28"
      Tab(3).Control(4)=   "Label27"
      Tab(3).Control(5)=   "Label26"
      Tab(3).Control(6)=   "Label25"
      Tab(3).Control(7)=   "cmd_delete_setting"
      Tab(3).Control(8)=   "cmd_addnew_set"
      Tab(3).Control(9)=   "Text18"
      Tab(3).Control(10)=   "Text17"
      Tab(3).Control(11)=   "Text16"
      Tab(3).Control(12)=   "cmd_add_setting"
      Tab(3).Control(13)=   "cmb_mtype"
      Tab(3).Control(14)=   "cmb_rtype"
      Tab(3).Control(15)=   "Frame4"
      Tab(3).ControlCount=   16
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         Caption         =   "Search Existing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -65880
         TabIndex        =   75
         Top             =   1320
         Width           =   8775
         Begin VB.CommandButton cmd_setting_ok 
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3000
            Picture         =   "Form6.frx":2F160
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   2640
            Width           =   2055
         End
         Begin VB.ComboBox cmb_rtype2 
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
            Left            =   3000
            TabIndex        =   77
            Top             =   1800
            Width           =   3495
         End
         Begin VB.ComboBox cmb_mtype2 
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
            Left            =   3000
            TabIndex        =   76
            Top             =   1080
            Width           =   3495
         End
         Begin VB.PictureBox Picture1 
            Height          =   3495
            Left            =   0
            Picture         =   "Form6.frx":3851A
            ScaleHeight     =   3435
            ScaleWidth      =   8715
            TabIndex        =   79
            Top             =   360
            Width           =   8775
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Resource Type:"
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
               Left            =   600
               TabIndex        =   81
               Top             =   1440
               Width           =   2415
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Member Type:"
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
               Left            =   720
               TabIndex        =   80
               Top             =   720
               Width           =   2175
            End
         End
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   -70200
         TabIndex        =   48
         Top             =   2280
         Width           =   3255
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   -71400
         TabIndex        =   2
         Top             =   4140
         Width           =   3855
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   3600
         TabIndex        =   47
         Top             =   4500
         Width           =   3855
      End
      Begin VB.CommandButton cmd_add_book 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -70560
         Picture         =   "Form6.frx":418D4
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   8280
         Width           =   2175
      End
      Begin VB.TextBox Text15 
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
         Left            =   -70200
         TabIndex        =   45
         Top             =   6540
         Width           =   3255
      End
      Begin VB.TextBox Text14 
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
         Left            =   -70200
         TabIndex        =   44
         Top             =   5580
         Width           =   3255
      End
      Begin VB.TextBox Text13 
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
         Left            =   -70200
         TabIndex        =   43
         Top             =   4500
         Width           =   3255
      End
      Begin VB.TextBox Text12 
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
         Left            =   -70200
         TabIndex        =   42
         Top             =   3420
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3600
         TabIndex        =   41
         Top             =   7080
         Width           =   3855
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   3600
         TabIndex        =   40
         Top             =   5340
         Width           =   3855
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
         Left            =   3600
         TabIndex        =   39
         Top             =   6240
         Width           =   3855
      End
      Begin VB.TextBox Text2 
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
         Left            =   3600
         TabIndex        =   38
         Top             =   3540
         Width           =   3855
      End
      Begin VB.TextBox Text1 
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
         Left            =   -71400
         TabIndex        =   1
         Top             =   3300
         Width           =   3855
      End
      Begin VB.CommandButton cmd_add_teach 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4440
         Picture         =   "Form6.frx":44D2E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   8400
         Width           =   2295
      End
      Begin VB.TextBox Text4 
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
         Left            =   3600
         TabIndex        =   36
         Top             =   2580
         Width           =   3855
      End
      Begin VB.CommandButton cmd_add_stu 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -70320
         Picture         =   "Form6.frx":48188
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   8160
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -71400
         TabIndex        =   5
         Top             =   6540
         Width           =   3855
      End
      Begin VB.TextBox Text8 
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
         Left            =   -71400
         TabIndex        =   4
         Top             =   5700
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   -71400
         TabIndex        =   3
         Top             =   4860
         Width           =   3855
      End
      Begin VB.TextBox Text5 
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
         Left            =   -71400
         TabIndex        =   0
         Top             =   2460
         Width           =   3855
      End
      Begin VB.ComboBox cmb_rtype 
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
         Left            =   -71160
         TabIndex        =   34
         Top             =   2880
         Width           =   3495
      End
      Begin VB.ComboBox cmb_mtype 
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
         Left            =   -71160
         TabIndex        =   33
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton cmd_add_setting 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71640
         Picture         =   "Form6.frx":4B5E2
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   8040
         Width           =   2175
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71160
         TabIndex        =   31
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -71160
         TabIndex        =   30
         Top             =   5400
         Width           =   3495
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71160
         TabIndex        =   29
         Top             =   4560
         Width           =   3495
      End
      Begin VB.CommandButton cmd_addnew_stu 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -72600
         Picture         =   "Form6.frx":4EA3C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   8160
         Width           =   2055
      End
      Begin VB.CommandButton cmd_addnew_teach 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         Picture         =   "Form6.frx":51E96
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   8400
         Width           =   2055
      End
      Begin VB.CommandButton cmd_addnew_book 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -72840
         Picture         =   "Form6.frx":552F0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   8280
         Width           =   2055
      End
      Begin VB.CommandButton cmd_addnew_set 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -73920
         Picture         =   "Form6.frx":5874A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8040
         Width           =   2055
      End
      Begin VB.CommandButton cmd_delete_stu 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68160
         Picture         =   "Form6.frx":5BBA4
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   8160
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Search Existing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -66120
         TabIndex        =   20
         Top             =   1440
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3735
            Left            =   480
            TabIndex        =   22
            Top             =   1800
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   6588
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
            Caption         =   "Select The Student Record"
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
         Begin VB.TextBox Text10 
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
            Left            =   3000
            TabIndex        =   21
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Member Name:"
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
            TabIndex        =   23
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Image Image8 
            Height          =   7215
            Left            =   120
            Picture         =   "Form6.frx":5EFFE
            Top             =   360
            Width           =   11415
         End
      End
      Begin VB.CommandButton cmd_delete_teach 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6960
         Picture         =   "Form6.frx":683B8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Caption         =   "Search Existing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   9000
         TabIndex        =   15
         Top             =   1560
         Width           =   9855
         Begin VB.TextBox Text7 
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
            Left            =   3000
            TabIndex        =   16
            Top             =   1080
            Width           =   3135
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   3855
            Left            =   480
            TabIndex        =   17
            Top             =   1800
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   6800
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
            Caption         =   "Select the Teacher Record"
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
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Member Name:"
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
            TabIndex        =   18
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Image Image6 
            Height          =   7215
            Left            =   120
            Picture         =   "Form6.frx":6B812
            Top             =   360
            Width           =   11415
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C000&
         Caption         =   "Search Existing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -66000
         TabIndex        =   11
         Top             =   1800
         Width           =   9855
         Begin VB.TextBox Text11 
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
            Left            =   3000
            TabIndex        =   12
            Top             =   1080
            Width           =   3135
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   3975
            Left            =   480
            TabIndex        =   13
            Top             =   1800
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7011
            _Version        =   393216
            BackColor       =   16776960
            ColumnHeaders   =   -1  'True
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
            Caption         =   "Select the Book Record"
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
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Book Name:"
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
            Left            =   720
            TabIndex        =   14
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Image Image4 
            Height          =   7215
            Left            =   120
            Picture         =   "Form6.frx":74BCC
            Top             =   360
            Width           =   11415
         End
      End
      Begin VB.CommandButton cmd_delete_book 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68160
         Picture         =   "Form6.frx":7DF86
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8280
         Width           =   1935
      End
      Begin VB.CommandButton cmd_delete_setting 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69240
         Picture         =   "Form6.frx":813E0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8040
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Type:"
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
         Left            =   -73320
         TabIndex        =   74
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Below Resource Information"
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
         Left            =   -73680
         TabIndex        =   73
         Top             =   600
         Width           =   8775
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:-"
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
         Left            =   -73320
         TabIndex        =   72
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:-"
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
         Left            =   1440
         TabIndex        =   71
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Price(Rs.):"
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
         Left            =   -72600
         TabIndex        =   70
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition:"
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
         Left            =   -72240
         TabIndex        =   69
         Top             =   5580
         Width           =   2295
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         Height          =   615
         Left            =   -72240
         TabIndex        =   68
         Top             =   4500
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   -72000
         TabIndex        =   67
         Top             =   3420
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TeacherID:-"
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
         Left            =   960
         TabIndex        =   66
         Top             =   2580
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Department:-"
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
         Left            =   720
         TabIndex        =   65
         Top             =   5340
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   64
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ContactNo.:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   63
         Top             =   6240
         Width           =   2775
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
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         TabIndex        =   62
         Top             =   3540
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Below Teacher Information"
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
         Left            =   720
         TabIndex        =   61
         Top             =   780
         Width           =   8655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Below Student Information"
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
         Left            =   -73920
         TabIndex        =   60
         Top             =   840
         Width           =   8775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -73560
         TabIndex        =   59
         Top             =   6540
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ContactNo.:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -73920
         TabIndex        =   58
         Top             =   5700
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Course:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -73440
         TabIndex        =   57
         Top             =   4860
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:-"
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
         Left            =   -73200
         TabIndex        =   56
         Top             =   3300
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ScholarNo.:-"
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
         Left            =   -73920
         TabIndex        =   55
         Top             =   2460
         Width           =   2775
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Fine Amount:"
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
         Left            =   -73440
         TabIndex        =   54
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Duration:"
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
         Left            =   -73680
         TabIndex        =   53
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Quantity:"
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
         Left            =   -73680
         TabIndex        =   52
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Type:"
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
         Left            =   -73800
         TabIndex        =   51
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Member Type:"
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
         Left            =   -73560
         TabIndex        =   50
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Below Setting Information"
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
         Left            =   -73920
         TabIndex        =   49
         Top             =   600
         Width           =   7935
      End
      Begin VB.Image Image1 
         Height          =   14850
         Left            =   -75000
         Picture         =   "Form6.frx":8483A
         Top             =   360
         Width           =   20460
      End
      Begin VB.Image Image3 
         Height          =   14850
         Left            =   -75000
         Picture         =   "Form6.frx":9A408
         Top             =   360
         Width           =   20460
      End
      Begin VB.Image Image5 
         Height          =   14850
         Left            =   0
         Picture         =   "Form6.frx":AFFD6
         Top             =   360
         Width           =   20460
      End
      Begin VB.Image Image7 
         Height          =   14850
         Left            =   -75000
         Picture         =   "Form6.frx":C5BA4
         Top             =   360
         Width           =   20460
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW RECORD"
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
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cmd As ADODB.Command
Dim temp As Long 'for memberid
Dim save_update4 As Integer
Dim save_update3 As Integer
Dim save_update2 As Integer
Dim save_update1 As Integer
Dim check1 As Boolean 'for checking memberid at lostfocus event
Dim check2 As Boolean
Private Sub cmb_rtype_LostFocus()
Set rs = New ADODB.Recordset
If cmb_mtype.Text <> "" Then
    rs.Open "select * from membersetting", con, 3, 3
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        If rs(1) = cmb_rtype.ItemData(cmb_rtype.ListIndex) And rs(0) = cmb_mtype.ItemData(cmb_mtype.ListIndex) Then
            MsgBox "Member Settings Already Exists"
            cmb_mtype.SetFocus
        End If
        rs.MoveNext
    Next
    rs.Close
Else
    cmb_mtype.SetFocus
End If
End Sub

Private Sub cmd_add_setting_Click()
Set cmd = New ADODB.Command
cmd.CommandType = adCmdText
cmd.ActiveConnection = con
If cmb_mtype.Text <> "" And cmb_rtype.Text <> "" And Text16.Text <> "" And Text18.Text <> "" And Text17.Text <> "" Then
    If save_update4 = 1 Then
        cmd.CommandText = "insert into membersetting values(" & cmb_mtype.ItemData(cmb_mtype.ListIndex) & "," & cmb_rtype.ItemData(cmb_rtype.ListIndex) & "," & Text16.Text & "," & Text18.Text & "," & Text17.Text & ") "
        cmd.Execute
        save_update4 = 0
        MsgBox "Member Setting Data Saved Successfully", vbInformation
    ElseIf save_update4 = 2 Then
        cmd.CommandText = "update  membersetting set qty = " & Text16.Text & " , duration = " & Text18.Text & ",fine_amt = " & Text17.Text & " where mtype_id = " & cmb_mtype.ItemData(cmb_mtype.ListIndex) & " and rtype_id = " & cmb_rtype.ItemData(cmb_rtype.ListIndex) & ""
        cmd.Execute
        save_update4 = 0
        MsgBox "Member Setting Data Updated Successfully", vbInformation
    ElseIf save_update4 = 0 Then
        MsgBox "Click on Add New to add new setting", vbInformation
    End If
Else
    MsgBox "Please Fill all Information ", vbCritical
End If
End Sub


Private Sub cmd_add_stu_Click()
Set rs = New ADODB.Recordset
Set cmd = New ADODB.Command
cmd.CommandType = adCmdText
cmd.ActiveConnection = con
If Text5.Text <> "" And Text1.Text <> "" And Text8.Text <> "" And Text9.Text <> "" And Combo4.Text <> "" And Combo1.Text <> "" Then
    If save_update1 = 1 Then
        rs.Open "select max(memberid) from member", con, 3, 3
            temp = rs(0)
            temp = temp + 1
        rs.Close
        cmd.CommandText = "insert into member values(" & temp & "," & 1 & ")"
        cmd.Execute
        cmd.CommandText = "insert into student values(" & temp & "," & Val(Text5.Text) & "," & Combo1.ItemData(Combo1.ListIndex) & " ,' " & Text1.Text & " ','" & Combo4.Text & "'," & Text8.Text & ",'" & Text9.Text & "') "
        cmd.Execute
        save_update1 = 0
        MsgBox "Student Data Saved Successfully and MemberID is " & temp, vbInformation
    ElseIf save_update1 = 2 Then
        cmd.CommandText = "update  student set stuname = '" & Text1.Text & "' , scholarno = " & Text5.Text & ",stusex = '" & Combo4.Text & "',stuaddress = '" & Text9.Text & "',stucontact=" & Text8.Text & " where memberid = " & DataGrid1.Columns(0) & " "
        cmd.Execute
        If Combo1.ListIndex <> -1 Then
            cmd.CommandText = "update  student set courseid =" & Combo1.ItemData(Combo1.ListIndex) & " where memberid = " & DataGrid1.Columns(0) & ""
            cmd.Execute
        End If
        save_update1 = 0
        MsgBox "Student Data Updated Successfully", vbInformation
    ElseIf save_update1 = 0 Then
        MsgBox "Click on Add New to add New Student", vbInformation
    End If
Else
    MsgBox "Please Fill all Information ", vbCritical
End If
End Sub

Private Sub cmd_add_teach_Click()
Set cmd = New ADODB.Command
cmd.CommandType = adCmdText
cmd.ActiveConnection = con
If Text4.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text6.Text <> "" And Combo3.Text <> "" And Combo2.Text <> "" Then
    If save_update2 = 1 Then
        rs.Open "select max(memberid) from member", con, 3, 3
            temp = rs(0)
            temp = temp + 1
        rs.Close
        cmd.CommandText = "insert into member values(" & temp & "," & 2 & ")"
        cmd.Execute
        cmd.CommandText = "insert into teacher values(" & Val(Text4.Text) & "," & temp & "," & Combo2.ItemData(Combo2.ListIndex) & " ,' " & Text2.Text & " ','" & Combo3.Text & "','" & Text6.Text & "'," & Text3.Text & ") "
        cmd.Execute
        save_update2 = 0
        MsgBox "Teacher Data Saved Successfully and MemberID is " & temp, vbInformation
    ElseIf save_update2 = 2 Then
        cmd.CommandText = "update teacher set tname = '" & Text2.Text & "' , tid = " & Text4.Text & ",tsex = '" & Combo3.Text & "',taddress = '" & Text6.Text & "',tcontact=" & Text3.Text & " where memberid = " & DataGrid2.Columns(0) & " "
        cmd.Execute
        If Combo2.ListIndex <> -1 Then
            cmd.CommandText = "update  student set deptid =" & Combo2.ItemData(Combo2.ListIndex) & " where memberid = " & DataGrid2.Columns(0) & ""
            cmd.Execute
        End If
        
        save_update2 = 0
        MsgBox "Teacher Data Updated Successfully", vbInformation
    ElseIf save_update2 = 0 Then
        MsgBox "Click on Add New to add New Teacher", vbInformation
    End If
Else
    MsgBox "Please Fill all Information ", vbCritical
End If
End Sub

Private Sub cmd_add_book_Click()
Set cmd = New ADODB.Command
cmd.CommandType = adCmdText
cmd.ActiveConnection = con
If Text12.Text <> "" And Text13.Text <> "" And Text15.Text <> "" And Combo5.Text <> "" Then
    If save_update3 = 1 Then
        rs.Open "select max(bookid) from book", con, 3, 3
            temp = rs(0)
            temp = temp + 1
        rs.Close
        cmd.CommandText = "insert into book values(" & temp & ",'" & Text12.Text & " ',' " & Text13.Text & " '," & Val(Text14.Text) & "," & Val(Text15.Text) & "," & Combo5.ItemData(Combo5.ListIndex) & ", " & False & ") "
        cmd.Execute
        save_update3 = 0
        MsgBox "Book Data Saved Successfully and BookID is " & temp, vbInformation
    ElseIf save_update3 = 2 Then
        cmd.CommandText = "update  book set bookname = '" & Text12.Text & "' , author = '" & Text13.Text & "',edition = " & Text14.Text & ",price = " & Text15.Text & " where bookid = " & DataGrid3.Columns(0) & " "
        cmd.Execute
        save_update3 = 0
        MsgBox "Book Data Updated Successfully", vbInformation
    ElseIf save_update3 = 0 Then
        MsgBox "Click on Add New to add new Book", vbInformation
    End If
Else
    MsgBox "Please Fill all Information ", vbCritical
End If

End Sub

Private Sub cmd_addnew_book_Click()

save_update3 = 1
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Combo5.Text = ""
Combo5.SetFocus
Set DataGrid3.DataSource = Nothing
End Sub

Private Sub cmd_addnew_set_Click()
save_update4 = 1
cmb_rtype2.Text = ""
cmb_mtype2.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
cmb_mtype.Text = ""
cmb_rtype.Text = ""
cmb_mtype.SetFocus
End Sub

Private Sub cmd_addnew_teach_Click()
save_update2 = 1
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Combo3.Text = ""
Combo2.Text = ""
Text4.SetFocus
Set DataGrid2.DataSource = Nothing
End Sub



Private Sub cmd_addnew_stu_Click()
save_update1 = 1
Text5.Text = ""
Text1.Text = ""
Text8.Text = ""
Text9.Text = ""
Combo4.Text = ""
Combo1.Text = ""
Text5.SetFocus
Set DataGrid1.DataSource = Nothing
End Sub

Private Sub cmd_search_stu_Click()
frmsearch.Show
Me.Hide
End Sub

Private Sub cmd_delete_book_Click()

    If save_update3 <> 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.ActiveConnection = con
        If Text12.Text <> "" And Text13.Text <> "" And Text14.Text <> "" And Text15.Text <> "" And Combo5.Text <> "" Then
            sure = MsgBox("Are You Sure Want To Delete The Book Record", 1)
            If sure = 1 Then
                rs.Open "select issueflag from book where bookid=" & DataGrid3.Columns(0) & "", con, 3, 3
                Flag = rs(0)    'storing flagvalue
                rs.Close
                If Flag = False Then
                    cmd.CommandText = "delete from book where bookid=" & DataGrid3.Columns(0) & ""
                    cmd.Execute
                    MsgBox "Record Deleted Successfully", vbInformation
                    Text12.Text = ""
                    Text13.Text = ""
                    Text14.Text = ""
                    Text15.Text = ""
                    Combo5.Text = ""
                    Set DataGrid3.DataSource = Nothing
                Else
                    MsgBox "Book is Issued to Someone,It Cannot Be deleted ", vbCritical
                End If
            Else
                MsgBox "Book Record Not Deleted", vbInformation
            End If
        Else
            MsgBox "Please Select the Book Record", vbCritical
        End If
    Else
        MsgBox "Please Select the Book Record", vbCritical
    End If

End Sub

Private Sub cmd_delete_setting_Click()

    If save_update4 <> 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.ActiveConnection = con
        If cmb_mtype.Text <> "" And cmb_rtype.Text <> "" And Text16.Text <> "" And Text18.Text <> "" And Text17.Text <> "" Then
            sure = MsgBox("Are You Sure Want To Delete The Member Setting Record", 1)
            If sure = 1 Then
                cmd.CommandText = "delete from membersetting where mtype_id=" & cmb_mtype.ItemData(cmb_mtype.ListIndex) & " and rtype_id=" & cmb_rtype.ItemData(cmb_rtype.ListIndex) & ""
                cmd.Execute
                MsgBox "Record Deleted Successfully", vbInformation
                cmb_rtype.Text = ""
                cmb_mtype.Text = ""
                Text16.Text = ""
                Text17.Text = ""
                Text18.Text = ""
            Else
                MsgBox "Member Setting Record Not Deleted", vbInformation
            End If
        Else
            MsgBox "Please Select the Member Setting", vbCritical
        End If
    Else
        MsgBox "Please Select the Member Setting Record", vbCritical
    End If

End Sub

Private Sub cmd_delete_stu_Click()

    If save_update1 <> 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.ActiveConnection = con
        If Text5.Text <> "" And Text1.Text <> "" And Text8.Text <> "" And Text9.Text <> "" And Combo4.Text <> "" And Combo1.Text <> "" Then
            sure = MsgBox("Are You Sure Want To Delete The Student Record", 1)
            If sure = 1 Then
                cmd.CommandText = "delete from member where  memberid=" & DataGrid1.Columns(0) & ""
                cmd.Execute
                cmd.CommandText = "delete from student where memberid=" & DataGrid1.Columns(0) & ""
                cmd.Execute
                MsgBox "Record Deleted Successfully", vbInformation
                Text5.Text = ""
                Text1.Text = ""
                Text8.Text = ""
                Text9.Text = ""
                Combo4.Text = ""
                Combo1.Text = ""
                Text10.Text = ""
                Set DataGrid1.DataSource = Nothing
            Else
                MsgBox "Student Record Not Deleted", vbInformation
            End If
        Else
            MsgBox "Please Select the Student Record", vbCritical
        End If
    Else
        MsgBox "Please Select the Student Record", vbCritical
    End If

End Sub

Private Sub cmd_delete_teach_Click()

    If save_update2 <> 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.ActiveConnection = con
        If Text4.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text6.Text <> "" And Combo3.Text <> "" And Combo2.Text <> "" Then
            sure = MsgBox("Are You Sure Want To Delete The Teacher Record", 1)
            If sure = 1 Then
                cmd.CommandText = "delete from member where  memberid=" & DataGrid2.Columns(0) & ""
                cmd.Execute
                cmd.CommandText = "delete from teacher where memberid=" & DataGrid2.Columns(0) & ""
                cmd.Execute
                MsgBox "Record Deleted Successfully", vbInformation
                Text4.Text = ""
                Text2.Text = ""
                Text3.Text = ""
                Text6.Text = ""
                Combo3.Text = ""
                Combo2.Text = ""
                Set DataGrid2.DataSource = Nothing
            Else
                MsgBox "Teacher Record Not Deleted", vbInformation
            End If
        Else
            MsgBox "Please Select the Teacher Record", vbCritical
        End If
    Else
        MsgBox "Please Select the Teacher Record", vbCritical
    End If

End Sub

Private Sub cmd_setting_ok_Click()
If cmb_rtype2.ListIndex <> -1 And cmb_mtype2.ListIndex <> -1 Then
      rs.Open "select *  from  membersetting where mtype_id=" & cmb_mtype.ItemData(cmb_mtype2.ListIndex) & " and rtype_id=" & cmb_rtype2.ItemData(cmb_rtype2.ListIndex), con, 3, 3
      If rs.RecordCount > 0 Then
        cmb_mtype.ListIndex = cmb_mtype2.ListIndex
        cmb_rtype.ListIndex = cmb_rtype2.ListIndex
        Text16.Text = rs(2)
        Text18.Text = rs(3)
        Text17.Text = rs(4)
        cmb_mtype.Locked = True
        cmb_rtype.Locked = True
        save_update4 = 2
      Else
        MsgBox "Sorry, there is no data for this member and resource type", vbCritical
      End If
      rs.Close
Else
  MsgBox "Please select member and resource type", vbCritical
End If
End Sub


Private Sub Command1_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub DataGrid1_Click()
If DataGrid1.Row <> -1 Then
    Set rs = New ADODB.Recordset
    i = DataGrid1.Row
    DataGrid1.RowBookmark (i)
    rs.Open "select * from student where memberid=" & DataGrid1.Columns(0) & "", con, 3, 3
    If rs.RecordCount > 0 Then
            Text5.Text = rs(1)
            Text1.Text = rs(3)
            Text8.Text = rs(5)
            Text9.Text = rs(6)
            Combo4.Text = rs(4)
            temp = rs("courseid")
            save_update1 = 2
    End If
    rs.Close
    rs.Open "select coursename from course where courseid=" & temp & "", con, 3, 3
    If rs.RecordCount > 0 Then
            Combo1.Text = rs(0)
    End If
    rs.Close
    check1 = True
End If
End Sub

Private Sub DataGrid2_Click()
If DataGrid2.Row <> -1 Then
    Set rs = New ADODB.Recordset
    i = DataGrid2.Row
    DataGrid2.RowBookmark (i)
    rs.Open "select * from teacher where memberid=" & DataGrid2.Columns(0) & "", con, 3, 3
    If rs.RecordCount > 0 Then
            Text4.Text = rs(0)
            Text2.Text = rs(3)
            Text3.Text = rs(6)
            Text6.Text = rs(5)
            Combo3.Text = rs(4)
            temp = rs("deptid")
            save_update2 = 2
    
    End If
    rs.Close
    rs.Open "select deptname from department where deptid=" & temp & "", con, 3, 3
    If rs.RecordCount > 0 Then
                Combo2.Text = rs(0)
                check2 = True
    End If
    rs.Close
End If
End Sub

Private Sub DataGrid3_Click()
If DataGrid3.Row <> -1 Then
    Set rs = New ADODB.Recordset
    i = DataGrid3.Row
    DataGrid3.RowBookmark (i)
    rs.Open "select * from book where bookid=" & DataGrid3.Columns(0) & "", con, 3, 3
    If rs.RecordCount > 0 Then
            Text12.Text = rs(1)
            Text13.Text = rs(2)
            Text14.Text = rs(3)
            Text15.Text = rs(4)
            temp = rs(5)
            save_update3 = 2
    End If
    rs.Close
    rs.Open "select r_name from resource where rtype_id=" & temp & "", con, 3, 3
    If rs.RecordCount > 0 Then
            Combo5.Text = rs(0)
    End If
    rs.Close
End If
End Sub

Private Sub Form_Load()
con.Close
connect
rs.Open "select rtype_id,r_name from resource", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    cmb_rtype.AddItem rs("r_name")
    cmb_rtype.ItemData(i - 1) = rs("rtype_id")
    cmb_rtype2.AddItem rs("r_name")
    cmb_rtype2.ItemData(i - 1) = rs("rtype_id")
    rs.MoveNext
Next i
rs.Close

rs.Open "select mtype_id,mtype_name from membertype", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    cmb_mtype.AddItem rs("mtype_name")
    cmb_mtype.ItemData(i - 1) = rs("mtype_id")
    cmb_mtype2.AddItem rs("mtype_name")
    cmb_mtype2.ItemData(i - 1) = rs("mtype_id")
    rs.MoveNext
Next i
rs.Close

rs.Open "select rtype_id,r_name from resource", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo5.AddItem rs("r_name")
    Combo5.ItemData(i - 1) = rs("rtype_id")
    rs.MoveNext
Next i
rs.Close

rs.Open "select deptid, deptname from department", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo2.AddItem rs("deptname")
    Combo2.ItemData(i - 1) = rs("deptid")
    rs.MoveNext
Next i
rs.Close

rs.Open "select courseid,coursename from course", con, 3, 3
rs.MoveFirst
For i = 1 To rs.RecordCount
    Combo1.AddItem rs.Fields(1)
    Combo1.ItemData(i - 1) = rs("courseid")
    rs.MoveNext
Next i
rs.Close

Combo4.AddItem "male"
Combo4.AddItem "female"

Combo3.AddItem "male"
Combo3.AddItem "female"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text5.Text = Text5.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text10_Change()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select memberid as MemberID,scholarno as ScholarNO,stuname as Name,coursename as Course,stusex as Gender,stucontact as ContactNO,stuaddress as Address from student,course where stuname like '" & "%" & Text10.Text & "%' and course.courseid=student.courseid", con, 3, 3
If rs.RecordCount > 0 Then
    Set DataGrid1.DataSource = rs
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text10.Text = Text10.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text11_Change()

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    rs.Open "select bookid as BookID,bookname as Book_Name,author as Author,edition as Edition,price as Price from book where bookname like '" & "%" & Text11.Text & "%'", con, 3, 3
    If rs.RecordCount > 0 Then
        Set DataGrid3.DataSource = rs
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text11.Text = Text11.Text & Chr(KeyAscii)
End If
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

Private Sub Text16_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text16.Text = Text16.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text17.Text = Text17.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text18.Text = Text18.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
    KeyAscii = 0
    Text2.Text = Text2.Text & Chr(KeyAscii)
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(Text3.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            Text3.Text = Text3.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        Text3.Text = Text3.Text & Chr(KeyAscii)
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text4.Text = Text4.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text4_LostFocus()
Set rs = New ADODB.Recordset
If check2 = False Then
If Text4.Text <> "" Then
        rs.Open "select * from teacher", con, 3, 3
        rs.MoveFirst
        For j = 0 To rs.RecordCount - 1
            If rs(0) = Val(Text4.Text) Then
                MsgBox ("Please Give Another TeacherID")
            End If
            rs.MoveNext
        Next
        rs.Close
End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = False And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text5.Text = Text5.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text5_LostFocus()
Set rs = New ADODB.Recordset
   If check1 = False Then
   If Text5.Text <> "" Then
        rs.Open "select * from student", con, 3, 3
        rs.MoveFirst
            For j = 0 To rs.RecordCount - 1
                If rs(1) = Val(Text5.Text) Then
                    MsgBox ("Please Give Another ScholarNO. ")
                    Text5.SetFocus
                End If
                rs.MoveNext
            Next
        rs.Close
    End If
    End If
End Sub

Private Sub Text7_change()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    rs.Open "select memberid as MemberID,tid as TeacherID,tname as Name,deptname as Department,tsex as Gender,taddress as Address,tcontact as ContactNO from teacher,department where tname like '" & "%" & Text7.Text & "%' and department.deptid=teacher.deptid", con, 3, 3
    If rs.RecordCount > 0 Then
        Set DataGrid2.DataSource = rs
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And Not KeyAscii = 8 Then
    KeyAscii = 0
    Text7.Text = Text7.Text & Chr(KeyAscii)
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(Text8.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            Text5.Text = Text5.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        Text5.Text = Text5.Text & Chr(KeyAscii)
    End If
End If
End Sub
