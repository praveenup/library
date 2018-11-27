VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16080
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1905
      ButtonWidth     =   1640
      ButtonHeight    =   1746
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Student"
            Key             =   "student"
            Description     =   "Student Master"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Teacher"
            Key             =   "teacher"
            Description     =   "Teacher Master"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Book"
            Key             =   "book"
            Description     =   "Book Master"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Member"
            Key             =   "member"
            Description     =   "Member settings"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Issue"
            Key             =   "issue"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Deposit"
            Key             =   "deposit"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&earch"
            Key             =   "search"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&tatus"
            Key             =   "status"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Report"
            Key             =   "report"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "E&xit"
            Key             =   "exit"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   8160
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   44
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":5EC59
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6040B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6199D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":62F97
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":64399
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":65653
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":672C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":685D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6A2A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6BBF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6D245
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnumaster 
      Caption         =   "&Master"
      Begin VB.Menu mnustu 
         Caption         =   "&Student"
      End
      Begin VB.Menu mnuteach 
         Caption         =   "&Teacher"
      End
      Begin VB.Menu mnuresource 
         Caption         =   "&Resource"
      End
      Begin VB.Menu mnusetting 
         Caption         =   "&Member Setting"
      End
   End
   Begin VB.Menu tran 
      Caption         =   "&Transaction"
      Index           =   3
      Begin VB.Menu bissue 
         Caption         =   "&Issue"
      End
      Begin VB.Menu bdeposit 
         Caption         =   "&Deposit"
      End
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "&Information"
      Begin VB.Menu mnuteacher 
         Caption         =   "&Teacher"
      End
      Begin VB.Menu mnustudent 
         Caption         =   "&Student"
      End
      Begin VB.Menu mnubook 
         Caption         =   "&Book"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Status"
      Index           =   2
      Begin VB.Menu Mstatus 
         Caption         =   "&MemberStatus"
      End
      Begin VB.Menu Bstatus 
         Caption         =   "&BookStatus"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnufine 
         Caption         =   "&Fine "
      End
   End
   Begin VB.Menu mdiexit 
      Caption         =   "&Exit"
      Index           =   4
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bdeposit_Click()
frmtrans.SSTab1.Tab = 1
End Sub

Private Sub bissue_Click()
frmtrans.SSTab1.Tab = 0
End Sub

Private Sub Bstatus_Click()
frmstatus.SSTab1.Tab = 1
End Sub






Private Sub mdiexit_Click(Index As Integer)
con.Close
frmlogin.Show
End Sub

Private Sub mnubook_Click()
frmsearch.SSTab1.Tab = 0
End Sub


Private Sub mnufine_Click()
DataReport2.Show
End Sub

Private Sub mnuissueboks_Click()
DataReport1.Show
End Sub

Private Sub mnuresource_Click()
frmadd.SSTab1.Tab = 2
End Sub

Private Sub mnusetting_Click()
frmadd.SSTab1.Tab = 3
End Sub

Private Sub mnustu_Click()
frmadd.SSTab1.Tab = 0
End Sub

Private Sub mnustudent_Click()
frmsearch.SSTab1.Tab = 1
End Sub

Private Sub mnuteach_Click()
frmadd.SSTab1.Tab = 1
End Sub

Private Sub mnuteacher_Click()
frmsearch.SSTab1.Tab = 2
End Sub

Private Sub Mstatus_Click()
frmstatus.SSTab1.Tab = 0
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
    Case "student"
        mnustu_Click
    Case "book"
        mnuresource_Click
    Case "issue"
        bissue_Click
    Case "deposit"
        bdeposit_Click
    Case "teacher"
        mnuteach_Click
    Case "member"
        mnusetting_Click
    Case "search"
        mnubook_Click
    Case "status"
        Mstatus_Click
    Case "report"
        mnufine_Click

End Select
End Sub


