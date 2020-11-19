VERSION 5.00
Begin VB.Form InsertDisk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Disk"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "InsertDisk.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Please insert a disk into drive."
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "InsertDisk.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "InsertDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

