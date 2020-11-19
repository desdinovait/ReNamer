VERSION 5.00
Begin VB.Form CreateDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ReNamer"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "CreateDir.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Create Directory : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "\"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox NewDir 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "NewDir"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Current disk :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Current directory :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "CreateDir.frx":000C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "New directory name :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "CreateDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
   
Private Sub Command2_Click()
Dim newDirs As String
'Questa funzione risolve il problema del Dir1 riguardo lo slash
FILEx = Split(Text1.Text, "\", -1)
If FILEx(1) = "" Then
    newDirs = Text1.Text + newDir.Text + "\"
End If

If FILEx(1) <> "" Then
    newDirs = Text1.Text + "\" + newDir.Text
End If

MkDir newDirs
Form1.Dir2.Path = newDirs
Form1.Drive2.Drive = Drive1.Drive

'Refresh
Form1.Drive1.Refresh
Form1.Drive2.Refresh
Form1.Dir1.Refresh
Form1.Dir2.Refresh
Unload Me
End Sub

Private Sub Dir1_Change()
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo 10
Dir1.Path = Drive1.Drive
GoTo 20
10
InsertDisk.Show 1
20
End Sub

Private Sub Form_Load()
Me.Caption = "ReNamer " + Versione + " - Create Directory"
Me.Drive1.Drive = Form1.Drive2.Drive
Me.Dir1.Path = Form1.Dir2.Path
newDir.Text = Form1.Matrix.Text
Text1.Text = Dir1.Path
End Sub
