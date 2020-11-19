VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3105
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4905
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton Option2 
      Caption         =   "Random"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "In order"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   2520
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "Next Tip"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      Picture         =   "frmTip.frx":000C
      ScaleHeight     =   2835
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   735
         Left            =   0
         Top             =   2280
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   600
         X2              =   3480
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   0
         Picture         =   "frmTip.frx":0316
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   780
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Select a tip at random.
    If Option2.Value = True Then
     CurrentTip = Int((Tips.Count * Rnd) + 1)
    End If
    
    ' Or, you could cycle through the Tips in order

    If Option1.Value = True Then
     CurrentTip = CurrentTip + 1
     If Tips.Count < CurrentTip Then
         CurrentTip = 1
     End If
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub Form_Load()
Dim ShowAtStartup As Long
Dim tipo As Integer

Me.Caption = "ReNamer " + Versione + " - Tip of the Day"

On Error GoTo 10
tipo = GetSetting("ReNamer", "Prev", "Type")
If tipo = "0" Then Option1.Value = 1
If tipo = "1" Then Option2.Value = 1
   
10
    ' See if we should be shown at startup
 '   ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
 '   If ShowAtStartup = 0 Then
 '       Unload Me
 '       Exit Sub
 '   End If
        
    ' Set the checkbox, this will force the value to be written back out to the registry
   ' Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\TIPOFDAY.TXT") = False Then
        lblTipText.Caption = "Tips File not found..."
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Option1.Value = True Then SaveSetting "ReNamer", "Prev", "Type", "0"
If Option2.Value = True Then SaveSetting "ReNamer", "Prev", "Type", "1"


End Sub
