VERSION 5.00
Begin VB.Form Info 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ReNamer"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Info : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Modify filename :"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Modif 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Matrix :"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Initial number :"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Index :"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Position :"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Destination directory :"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Files number :"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Typee 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label FilesNumber 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Matrix 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label TotalIndex 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label InitialNumber 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4485
         TabIndex        =   6
         Top             =   1440
         Width           =   1530
      End
      Begin VB.Label Destination 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Info.frx":000C
         Top             =   360
         Width           =   480
      End
      Begin VB.Label examp 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "First filename example :"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
StartProcessFile
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Prio
Me.Caption = "ReNamer " + Versione + " - Info"
If Form1.Check1.Value = 1 Then Modif.Caption = "Yes"
If Form1.Check1.Value = 0 Then Modif.Caption = "No"

If Form1.Option1.Value = True Then Typee.Caption = "Before"
If Form1.Option2.Value = True Then Typee.Caption = "After"
FilesNumber.Caption = Form1.Number.Caption
Matrix.Caption = Form1.Matrix.Text
TotalIndex.Caption = Form1.InsertZero.Text
InitialNumber.Caption = Form1.Numero.Text
Destination.Caption = Form1.DestDir.Text
examp.Caption = Form1.Esempio.Text
End Sub

Private Sub TotalNumber_Click()

End Sub
