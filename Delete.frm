VERSION 5.00
Begin VB.Form Delete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ReNamer"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "Delete.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Bar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Left            =   -240
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Bord 
      Height          =   135
      Left            =   -120
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Delete.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure do you want to delete all origin files?"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim deleteFinale As Long
Dim deleteS As Long
Dim Scor
Dim killName As String

Bar.Visible = True
Bar.Width = 15
Scor = 0
Scor = Bord.Width \ Form1.Grid.Rows

deleteFinale = Val(Form1.Number.Caption)
On Error Resume Next
For deleteS = 1 To deleteFinale
    killName = Form1.Grid.TextMatrix(deleteS, 0) + Form1.Grid.TextMatrix(deleteS, 1)
    Kill killName
    'Scorrimento barra
    Bar.Width = Bar.Width + Scor
    If Bar.Width > Bord.Width Then Bar.Width = Bord.Width
    If Bar.Width = Bord.Width Then Bar.Visible = False: Bord.Visible = False
Next deleteS

Bar.Visible = False
Unload Me

If Form1.Number.Caption > 0 Then Form1.Command1.Enabled = True: Form1.Up.Enabled = True: Form1.Down.Enabled = True: Form1.Command7.Enabled = True: Form1.Prop.Enabled = True Else Form1.Command1.Enabled = False: Form1.Up.Enabled = False: Form1.Down.Enabled = False: Form1.Command7.Enabled = False: Form1.Prop.Enabled = False

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "ReNamer " + Versione + " - Delete Files"
End Sub
