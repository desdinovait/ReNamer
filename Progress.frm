VERSION 5.00
Begin VB.Form Progress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Total Progress - Please wait..."
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Label curNumProg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "\"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   870
         TabIndex        =   2
         Top             =   270
         Width           =   4995
      End
      Begin VB.Label curFileProg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "\"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   5295
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Progress.frx":0000
         Top             =   285
         Width           =   480
      End
      Begin VB.Shape Barra 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   720
         Top             =   240
         Width           =   15
      End
      Begin VB.Shape Bordo 
         Height          =   255
         Left            =   720
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

