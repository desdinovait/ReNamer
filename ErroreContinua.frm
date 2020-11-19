VERSION 5.00
Begin VB.Form ErroreContinua 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ReNamer"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "ErroreContinua.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select another file and try again."
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unable to determinate the matrix and index"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "of the selected file."
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "ErroreContinua.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "ErroreContinua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "ReNamer " + Versione + " - Error"
End Sub
