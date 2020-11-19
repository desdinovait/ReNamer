VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1155
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General "
      TabPicture(0)   =   "About.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Picture1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "License "
      TabPicture(1)   =   "About.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   240
         Picture         =   "About.frx":0044
         ScaleHeight     =   570
         ScaleWidth      =   540
         TabIndex        =   3
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail to: duff_stormsoftx@mail.com for bugs, infos and new versions"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label Label8 
         Caption         =   "for Windows 95/98/NT/2000/ME/XP/03"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "No warrenty :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   11
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":0D0E
         Height          =   735
         Left            =   -74670
         TabIndex        =   10
         Top             =   1605
         Width           =   5535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":0D98
         Height          =   855
         Left            =   -74640
         TabIndex        =   9
         Top             =   840
         Width           =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   840
         X2              =   5835
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   840
         X2              =   5805
         Y1              =   2055
         Y2              =   2040
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Files ReNamer"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Special thanks to :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   825
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tonoli Mauro - "
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Created by: Ferla Daniele - 2000/2005"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "for concept, test and other ideas"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   2520
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4845
      TabIndex        =   0
      Top             =   4395
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1890
      TabIndex        =   13
      Top             =   765
      Width           =   1515
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   120
      Picture         =   "About.frx":0E3B
      Top             =   120
      Width           =   6060
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "ReNamer " + Versione + " - About"
Label7.Caption = "ReNamer ver " + Versione + " - Freeware"
Label9.Caption = "Version " + Versione
End Sub

