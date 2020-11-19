VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReNamer"
   ClientHeight    =   8610
   ClientLeft      =   1710
   ClientTop       =   2055
   ClientWidth     =   11850
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Remove All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origin directory :"
      Height          =   4155
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Command6 
         Caption         =   "="
         Height          =   315
         Left            =   4800
         TabIndex        =   6
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Execute"
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "All >"
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   3720
         Width           =   735
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   3405
         Hidden          =   -1  'True
         Left            =   2520
         MultiSelect     =   2  'Extended
         System          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Prop 
         Caption         =   "Properties"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         TabIndex        =   11
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton Down 
         Caption         =   "Down"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         TabIndex        =   10
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton Up 
         Caption         =   "Up"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         TabIndex        =   9
         Top             =   1740
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10440
         TabIndex        =   12
         Top             =   3180
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Select >"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   3720
         Width           =   735
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Delete files after rename"
         Height          =   735
         Left            =   10425
         TabIndex        =   8
         Top             =   975
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3780
         Left            =   5280
         TabIndex        =   7
         Top             =   255
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   6668
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483648
         AllowBigSelection=   0   'False
         HighLight       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "<|<"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N° files:"
         Height          =   255
         Left            =   10440
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Number 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   10440
         TabIndex        =   44
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   8280
      TabIndex        =   37
      Top             =   4335
      Width           =   3495
      Begin VB.CheckBox Check2 
         Caption         =   "Images Preview :"
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   0
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.Image Immagine 
         Height          =   690
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Click to open default images viewer"
         Top             =   675
         Width           =   855
      End
      Begin VB.Label ImmagineDim 
         Caption         =   "\"
         Height          =   255
         Left            =   1080
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label ImmagineDimLab 
         Caption         =   "Dimensions:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image ImmagineRif 
         Height          =   2415
         Left            =   120
         Top             =   675
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About..."
      Height          =   375
      Left            =   9480
      TabIndex        =   33
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      TabIndex        =   34
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8280
      TabIndex        =   32
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10680
      TabIndex        =   35
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination directory :"
      Height          =   4185
      Left            =   120
      TabIndex        =   36
      Top             =   4335
      Width           =   8055
      Begin VB.CheckBox Check1 
         Caption         =   "Modify filename"
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   600
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "After"
         Height          =   255
         Left            =   6600
         TabIndex        =   29
         Top             =   2520
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Before"
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Auto-Matrix"
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "="
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton Last 
         Caption         =   "Last File"
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Su 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7155
         TabIndex        =   26
         Top             =   1575
         Width           =   240
      End
      Begin VB.CommandButton Giu 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6915
         TabIndex        =   25
         Top             =   1575
         Width           =   240
      End
      Begin VB.TextBox DestDir 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "\"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Esempio 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "\"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox Numero 
         Height          =   285
         Left            =   6600
         MaxLength       =   9
         TabIndex        =   27
         Text            =   "0"
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox InsertZero 
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "2"
         Top             =   1560
         Width           =   300
      End
      Begin VB.TextBox Matrix 
         Height          =   285
         Left            =   5280
         TabIndex        =   23
         Text            =   "Matrix"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.FileListBox File2 
         Height          =   3405
         Hidden          =   -1  'True
         Left            =   2520
         System          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   2655
      End
      Begin VB.DirListBox Dir2 
         Height          =   3465
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Execute"
         Height          =   315
         Left            =   4080
         TabIndex        =   19
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Position :"
         Height          =   255
         Left            =   5280
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label LabelIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "Total index :"
         Height          =   255
         Left            =   5280
         TabIndex        =   41
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label LabelExample 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "First filename example :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5280
         TabIndex        =   40
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label LabelNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial number :"
         Height          =   255
         Left            =   5280
         TabIndex        =   39
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label LabelName 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrix's name :"
         Height          =   255
         Left            =   5280
         TabIndex        =   38
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Menu MenuOrigin 
      Caption         =   "Origin"
      Visible         =   0   'False
      Begin VB.Menu MenuExecute 
         Caption         =   "Execute"
      End
      Begin VB.Menu MenuSelect 
         Caption         =   "Select"
      End
      Begin VB.Menu MenuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu MenuEqual 
         Caption         =   "Equal Directory"
      End
      Begin VB.Menu MenuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu MenuDestination 
      Caption         =   "Destination"
      Visible         =   0   'False
      Begin VB.Menu MenuExecuteDest 
         Caption         =   "Execute"
      End
      Begin VB.Menu MenuEqualDEst 
         Caption         =   "Equal Directory"
      End
      Begin VB.Menu MenuAuto 
         Caption         =   "Auto-Matrix"
      End
      Begin VB.Menu MenuLastFile 
         Caption         =   "Last File"
      End
      Begin VB.Menu MenuRefreshDest 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nome
Dim n As Integer
Dim a As Variant
Dim W As Variant
Dim X As Variant
Dim Num
Dim DEST
Dim File
Dim VediFile
Dim currentRow As Long


Private Sub LoadRapImage(Name As String)
Dim X As Picture
Dim origWidth As Long
Dim origHeight As Long
Dim origRatio As Double
Dim maxDim As Long
Dim isWidth As Boolean

On Error GoTo 200
Immagine.Picture = Nothing
'Imposta l'immagine nel controllo con dimensioni in rapporto
ImmagineDim.Caption = "\"
ImmagineDim.Visible = False
ImmagineDimLab.Visible = False

Set X = LoadPicture(Name)

origWidth = ScaleX(X.Width, vbHimetric, vbPixels)
origHeight = ScaleY(X.Height, vbHimetric, vbPixels)
If origWidth <= origHeight Then
    origRatio = origHeight / origWidth
    maxDim = origHeight
    isWidth = False
Else
    origRatio = origWidth / origHeight
    maxDim = origWidth
    isWidth = True
End If
If isWidth = True Then
    Immagine.Width = ImmagineRif.Width
    Immagine.Height = ImmagineRif.Width / origRatio
Else
    Immagine.Width = ImmagineRif.Height / origRatio
    Immagine.Height = ImmagineRif.Height
End If

Immagine.Left = ImmagineRif.Left + ((ImmagineRif.Width - Immagine.Width) / 2)
Immagine.Top = ImmagineRif.Top + ((ImmagineRif.Height - Immagine.Height) / 2)
If (Immagine.Left < ImmagineRif.Left) Then Immagine.Left = ImmagineRif.Left
If (Immagine.Top < ImmagineRif.Top) Then Immagine.Top = ImmagineRif.Top

Immagine.Picture = LoadPicture(Name)
ImmagineDim.Caption = origWidth & " x "
ImmagineDim.Caption = ImmagineDim.Caption & origHeight
ImmagineDim.Visible = True
ImmagineDimLab.Visible = True

200
End Sub



Private Sub Check1_Click()
If Check1.Value = 1 Then Option1.Enabled = True: Option2.Enabled = True
If Check1.Value = 0 Then Option1.Enabled = False: Option2.Enabled = False
Prev
End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 0 Then
    LoadRapImage ("")
    Immagine.Enabled = False
End If

If Check2.Value = 1 Then
    LoadRapImage (VediFile)
    Immagine.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Info.Show 1
End Sub

Private Sub Command10_Click()
Dir1.Path = Dir2.Path
Drive1.Drive = Drive2.Drive
End Sub

Private Sub Command11_Click()
CreateDir.Show 1
End Sub

Private Sub Command12_Click()
'SELEZIONA TUTTI I FILE DELLA DIRECTORY CORRENTE

Dim Sel As Long
'Questa funzione risolve il problema del Dir1 riguardo lo slash
FILEx = Split(Form1.Dir1.Path, "\", -1)
If FILEx(1) = "" Then
    For Sel = 0 To File1.ListCount - 1
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
         Number.Caption = Number.Caption + 1
    Next Sel
End If

If FILEx(1) <> "" Then
   For Sel = 0 To File1.ListCount - 1
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path + "\"
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
         Number.Caption = Number.Caption + 1
   Next Sel
End If

'Altro
If Number.Caption > 0 Then Command1.Enabled = True: Up.Enabled = True: Down.Enabled = True: Command7.Enabled = True: Command9.Enabled = True: Prop.Enabled = True Else Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False

'Print Lista.ListCount
CalculateStart
End Sub

Private Sub Command13_Click()
Dim DEST3
Dim Destdir3
Dim Avvia
'Questa funzione risolve il problema del Dir1 riguardo lo slash
DEST3 = Split(Form1.Dir2.Path, "\", -1)
If DEST3(1) = "" Then
    Destdir3 = Dir1.Path
End If
If DEST3(1) <> "" Then
    Destdir3 = Dir2.Path + "\"
End If

Avvia = Destdir3 + File2.fileName
Dim res&
res& = ShellExecute(hwnd, "open", Avvia, vbNullString, vbNullString, SW_SHOW)
If res < 32 Then
    MsgBox "Unable to open selected file"
End If
End Sub

Private Sub Command4_Click()
Drive1.Refresh
Drive2.Refresh
Dir1.Refresh
Dir2.Refresh
File1.Refresh
File2.Refresh
Grid.Refresh
End Sub

Private Sub Command5_Click()
Dim Sel As Long

'Questa funzione risolve il problema del Dir1 riguardo lo slash
FILEx = Split(Form1.Dir1.Path, "\", -1)
If FILEx(1) = "" Then
    For Sel = 0 To File1.ListCount - 1
     If File1.Selected(Sel) = True Then
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
         Number.Caption = Number.Caption + 1
     End If
    Next Sel
End If

If FILEx(1) <> "" Then
   For Sel = 0 To File1.ListCount - 1
    If File1.Selected(Sel) = True Then
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path + "\"
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
         Number.Caption = Number.Caption + 1
    End If
   Next Sel
End If

'Altro
If Number.Caption > 0 Then Command1.Enabled = True: Up.Enabled = True: Down.Enabled = True: Command7.Enabled = True: Command9.Enabled = True: Prop.Enabled = True Else Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False

'Print Lista.ListCount
CalculateStart
End Sub

Private Sub Prev()
Dim Joined As String
Dim Numero1 As String
Numero1 = Left(String(Val(InsertZero.Text), "0"), Len(String(Val(InsertZero.Text), "0")) - Len(Numero.Text)) & Numero.Text
If Check1.Value = 1 Then
    If Option1.Value = True Then Esempio.Text = Numero1 & Matrix.Text & ".ext"
    If Option2.Value = True Then Esempio.Text = Matrix.Text & Numero1 & ".ext"
Else
    Esempio.Text = Numero1 & Matrix.Text & "nomefile.ext"
End If
End Sub

Private Sub CalculateContinue(Stringa As String)
Dim StringaSP
Dim sottoStringa As String
Dim nomeMatrice As String
Dim numeroMatrice As String
Dim numeroMatrice2 As String

On Error GoTo 10

StringaSP = Split(Stringa, ".", -1)
Dim Indice As Integer

'Scompone la stringa
For i = 0 To Len(StringaSP(0))
    sottoStringa = Right(StringaSP(0), i)
    If IsNumeric(Left(sottoStringa, 1)) Then numeroMatrice = numeroMatrice + (Left(sottoStringa, 1)): Indice = Indice + 1
Next i
    
'Ricostruisce la stringa
For i = 0 To Len(numeroMatrice)
    sottoStringaMatrice = Right(numeroMatrice, i)
    If IsNumeric(Left(sottoStringaMatrice, 1)) Then
     numeroMatrice2 = numeroMatrice2 + (Left(sottoStringaMatrice, 1))
    End If
Next i

nomeMatrice = Left(StringaSP(0), Len(StringaSP(0)) - Len(numeroMatrice2))

'Riempie i campi
Matrix.Text = nomeMatrice
InsertZero = Indice
Numero.Text = numeroMatrice2 + 1
If InsertZero < 2 Then GoTo 10
GoTo 20

10
InsertZero.Text = 2
Numero.Text = 0
Matrix.Text = "Matrix"
ErroreContinua.Show 1

20
End Sub

Private Sub Command6_Click()
Dir2.Path = Dir1.Path
Drive2.Drive = Drive1.Drive
End Sub

Private Sub Command8_Click()
Dim DEST2
Dim Destdir2
Dim Avvia
'Questa funzione risolve il problema del Dir1 riguardo lo slash
DEST2 = Split(Form1.Dir1.Path, "\", -1)
If DEST2(1) = "" Then
    Destdir2 = Dir1.Path
End If
If DEST2(1) <> "" Then
    Destdir2 = Dir1.Path + "\"
End If

Avvia = Destdir2 + File1.fileName
Dim res&
'res& = ShellExecute(hwnd, "open", Avvia, vbNullString, vbNullString, SW_SHOW)
res& = ShellExecute(hwnd, "open", Avvia, vbNullString, Form1.Dir1.Path, SW_SHOW)
'Me.Caption = res
If res < 32 Then
    MsgBox "Unable to open selected file"
End If

End Sub

Private Sub Command9_Click()
Dim total As Long
Dim s As Long
total = Grid.Rows
For s = 1 To total - 2
    Grid.RemoveItem (Grid.Rows)
Next s
Grid.TextMatrix(1, 0) = ""
Grid.TextMatrix(1, 1) = ""
currentRow = 0
Number.Caption = 0

Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False
End Sub

Private Sub DestDir_Change()
Me.Caption = "ReNamer " + Versione + " - " + DestDir.Text
End Sub

Private Sub File1_DblClick()
'Questa funzione risolve il problema del Dir1 riguardo lo slash
FILEx = Split(Form1.Dir1.Path, "\", -1)
If FILEx(1) = "" Then
    For Sel = 0 To File1.ListCount - 1
     If File1.Selected(Sel) = True Then
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
     End If
    Next Sel
End If

If FILEx(1) <> "" Then
   For Sel = 0 To File1.ListCount - 1
    If File1.Selected(Sel) = True Then
         currentRow = currentRow + 1
         If currentRow > 1 Then Grid.Rows = Grid.Rows + 1
         Grid.TextMatrix(Grid.Rows - 1, 0) = Dir1.Path + "\"
         Grid.TextMatrix(Grid.Rows - 1, 1) = File1.List(Sel)
    End If
   Next Sel
End If

'Altro
Number.Caption = Grid.Rows - 1
If Number.Caption > 0 Then Command1.Enabled = True: Up.Enabled = True: Down.Enabled = True: Command7.Enabled = True: Command9.Enabled = True: Prop.Enabled = True Else Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False

CalculateStart

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu MenuOrigin
End Sub

Private Sub File2_Click()
Dim Master As String
On Error GoTo 10

'Questa funzione risolve il problema del Dir1 riguardo lo slash
If Check2.Value = 1 Then
    File = Split(Form1.Dir2.Path, "\", -1)
    If File(1) = "" Then
        Master = Dir2.Path + File2.fileName
    End If
    If File(1) <> "" Then
        Master = Dir2.Path + "\" + File2.fileName
    End If
    LoadRapImage (Master)
End If
GoTo 20

10

20
Dim DEST2
Dim Destdir2
'Questa funzione risolve il problema del Dir1 riguardo lo slash
DEST2 = Split(Form1.Dir2.Path, "\", -1)
If DEST2(1) = "" Then
    Destdir2 = Dir2.Path
End If
If DEST2(1) <> "" Then
    Destdir2 = Dir2.Path + "\"
End If
VediFile = Destdir2 + File2.fileName

If Check6.Value = 1 Then CalculateContinue (File2.fileName)
End Sub

Private Sub File2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu MenuDestination
End Sub

Private Sub File2_PathChange()
If File2.ListCount <= 0 Then Last.Enabled = False
If File2.ListCount > 0 Then Last.Enabled = True
End Sub

Private Sub Grid_DblClick()
If currentRow > 1 Then
    Grid.RemoveItem (Grid.RowSel)
End If
If currentRow = 1 Then
    Grid.TextMatrix(1, 0) = ""
    Grid.TextMatrix(1, 1) = ""
    Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False
End If
If currentRow >= 1 Then
    currentRow = currentRow - 1
    Number.Caption = Number.Caption - 1
End If
End Sub

Private Sub Grid_RowColChange()
Dim a As String
Dim b As String
On Error GoTo 10

'Questa funzione risolve il problema del Dir1 riguardo lo slash
a = Grid.TextMatrix(Grid.RowSel, 0)
b = Grid.TextMatrix(Grid.RowSel, 1)
If Check2.Value = 1 Then
    LoadRapImage (a + b)
End If
10

VediFile = a + b
End Sub

Private Sub Last_Click()
CalculateContinue (File2.List(File2.ListCount - 1))
End Sub

Private Sub MenuAuto_Click()
CalculateContinue (File2.fileName)
End Sub

Private Sub MenuEqual_Click()
Command6_Click
End Sub

Private Sub MenuEqualDEst_Click()
Command10_Click
End Sub

Private Sub MenuExecute_Click()
Command8_Click
End Sub

Private Sub MenuExecuteDest_Click()
Command13_Click
End Sub

Private Sub MenuLastFile_Click()
CalculateContinue (File2.List(File2.ListCount - 1))
End Sub

Private Sub MenuRefresh_Click()
File1.Refresh
Dir1.Refresh
Drive1.Refresh
End Sub

Private Sub MenuRefreshDest_Click()
File2.Refresh
Dir2.Refresh
Drive2.Refresh
End Sub

Private Sub MenuSelect_Click()
Command5_Click
End Sub
Private Sub MenuSelectAll_Click()
Command12_Click
End Sub

Private Sub Option1_Click()
Prev
End Sub

Private Sub Option2_Click()
Prev
End Sub

Private Sub Prop_Click()
Dim part1 As String
Dim part2 As String
part1 = Grid.TextMatrix(Grid.RowSel, 0)
part2 = Grid.TextMatrix(Grid.RowSel, 1)
s$ = part1 + part2

NomeSelezionato = s
frmView.Initialize s$
frmView.Show 1
10
End Sub

Private Sub Command7_Click()
On Error Resume Next
If currentRow > 1 Then
    Grid.RemoveItem (Grid.RowSel)
End If
If currentRow = 1 Then
    Grid.TextMatrix(1, 0) = ""
    Grid.TextMatrix(1, 1) = ""
    Command1.Enabled = False: Up.Enabled = False: Down.Enabled = False: Command7.Enabled = False: Command9.Enabled = False: Prop.Enabled = False
End If
If currentRow >= 1 Then
    currentRow = currentRow - 1
    Number.Caption = Number.Caption - 1
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo 10
Dir1.Path = Drive1.Drive
GoTo 20
10
InsertDisk.Show 1
20
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive2_Change()
On Error GoTo 10
Dir2.Path = Drive2.Drive
GoTo 20
10
InsertDisk.Show 1
20
End Sub
Private Sub Dir2_Change()
File2.Path = Dir2.Path
'Questa funzione risolve il problema del Dir1 riguardo lo slash
DEST = Split(Form1.Dir2.Path, "\", -1)
If DEST(1) = "" Then
    DestDir.Text = Dir2.Path
End If
If DEST(1) <> "" Then
    DestDir.Text = Dir2.Path + "\"
End If
End Sub

Private Sub Command2_Click()
About.Show 1
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub

Private Sub File1_Click()
Dim Master As String
On Error GoTo 10


'Questa funzione risolve il problema del Dir1 riguardo lo slash
If Check2.Value = 1 Then
    File = Split(Form1.Dir1.Path, "\", -1)
    If File(1) = "" Then
        Master = Dir1.Path + File1.fileName
    End If
    If File(1) <> "" Then
        Master = Dir1.Path + "\" + File1.fileName
    End If
    LoadRapImage (Master)
End If
GoTo 20

10

20
Dim DEST2
Dim Destdir2
'Questa funzione risolve il problema del Dir1 riguardo lo slash
DEST2 = Split(Form1.Dir1.Path, "\", -1)
If DEST2(1) = "" Then
    Destdir2 = Dir1.Path
End If
If DEST2(1) <> "" Then
    Destdir2 = Dir1.Path + "\"
End If
VediFile = Destdir2 + File1.fileName

End Sub


Private Sub Form_Load()
Dim DEST
Dim prevDrive1
Dim prevDrive2
Dim prevDir1
Dim prevDir2
Dim prevText1
Dim prevTip
Dim prevPreview
Dim prevRef
Dim prevInfo
Dim prevMatrix
Dim prevDelete
Dim prevGif
Dim prevBefore

Grid.ColWidth(0) = 1400
Grid.ColWidth(1) = 5500
Grid.TextMatrix(0, 0) = "Directory"
Grid.TextMatrix(0, 1) = "Filename"

'Caricamento Directory precedenti
On Error GoTo 11
prevDrive1 = GetSetting("ReNamer", "Prev", "Drive1")
prevDrive2 = GetSetting("ReNamer", "Prev", "Drive2")
prevDir1 = GetSetting("ReNamer", "Prev", "Dir1")
prevDir2 = GetSetting("ReNamer", "Prev", "Dir2")
prevPreview = GetSetting("ReNamer", "Prev", "Preview")
prevMatrix = GetSetting("ReNamer", "Prev", "Matrix")
prevDelete = GetSetting("ReNamer", "Prev", "Delete")
prevBefore = GetSetting("ReNamer", "Prev", "Before")
prevModify = GetSetting("ReNamer", "Prev", "Modify")

If prevPreview = "0" Then Check2.Value = 0: Immagine.Enabled = False
If prevPreview = "1" Then Check2.Value = 1: Immagine.Enabled = True
If prevMatrix = "0" Then Check6.Value = 0
If prevMatrix = "1" Then Check6.Value = 1
If prevDelete = "0" Then Check7.Value = 0
If prevDelete = "1" Then Check7.Value = 1
If prevBefore = "1" Then Option1.Value = True
If prevBefore = "0" Then Option2.Value = True
If prevModify = "0" Then Check1.Value = 0
If prevModify = "1" Then Check1.Value = 1

Drive1.Drive = prevDrive1
Drive2.Drive = prevDrive2
Dir1.Path = prevDir1
Dir2.Path = prevDir2

11
Prev
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting "ReNamer", "Prev", "Drive1", Drive1.Drive
SaveSetting "ReNamer", "Prev", "Dir1", Dir1.Path
SaveSetting "ReNamer", "Prev", "Drive2", Drive2.Drive
SaveSetting "ReNamer", "Prev", "Dir2", Dir2.Path

If Check2.Value = 1 Then SaveSetting "ReNamer", "Prev", "Preview", "1"
If Check2.Value = 0 Then SaveSetting "ReNamer", "Prev", "Preview", "0"
If Check6.Value = 1 Then SaveSetting "ReNamer", "Prev", "Matrix", "1"
If Check6.Value = 0 Then SaveSetting "ReNamer", "Prev", "Matrix", "0"
If Check7.Value = 1 Then SaveSetting "ReNamer", "Prev", "Delete", "1"
If Check7.Value = 0 Then SaveSetting "ReNamer", "Prev", "Delete", "0"
If Check1.Value = 1 Then SaveSetting "ReNamer", "Prev", "Modify", "1"
If Check1.Value = 0 Then SaveSetting "ReNamer", "Prev", "Modify", "0"

If Option1.Value = True Then SaveSetting "ReNamer", "Prev", "Before", "1"
If Option2.Value = True Then SaveSetting "ReNamer", "Prev", "Before", "0"

End Sub


Private Sub Immagine_DblClick()
Dim Avvia

Avvia = VediFile
'Print VediFile
Dim res&
res& = ShellExecute(hwnd, "open", Avvia, vbNullString, vbNullString, SW_SHOW)
If res < 32 Then
    MsgBox "Unable to open selected image"
End If

End Sub

Private Sub Matrix_Change()
Prev
End Sub

Private Sub Numero_Change()
On Error GoTo 10
GoTo 20
10
Numero.Text = 0
20
If Len(Numero.Text) > InsertZero.Text Then InsertZero.Text = InsertZero.Text + 1
If Numero.Text > 0 And Numero.Text <= 999999999 Then GoTo 30
30
Prev
End Sub

Private Sub Su_Click()
InsertZero.Text = InsertZero.Text + 1
If InsertZero.Text > 9 Then InsertZero.Text = 9
Prev
End Sub
Private Sub Giu_Click()
InsertZero.Text = InsertZero.Text - 1
If InsertZero.Text < Len(Numero.Text) Then InsertZero.Text = InsertZero.Text + 1
If InsertZero.Text < 2 Then InsertZero.Text = 2
Prev
End Sub

Private Sub Down_Click()
Dim Sopra1
Dim Sopra2
Dim Sotto1
Dim Sotto2
If Grid.Rows > 1 And Grid.RowSel < Grid.Rows - 1 Then
    Sopra1 = Grid.TextMatrix(Grid.RowSel, 0)
    Sopra2 = Grid.TextMatrix(Grid.RowSel, 1)
    Sotto1 = Grid.TextMatrix(Grid.RowSel + 1, 0)
    Sotto2 = Grid.TextMatrix(Grid.RowSel + 1, 1)
    Grid.TextMatrix(Grid.RowSel, 0) = Sotto1
    Grid.TextMatrix(Grid.RowSel, 1) = Sotto2
    Grid.TextMatrix(Grid.RowSel + 1, 0) = Sopra1
    Grid.TextMatrix(Grid.RowSel + 1, 1) = Sopra2
    Grid.RowSel = Grid.RowSel + 1
End If
Grid.Refresh
End Sub

Private Sub Up_Click()
Dim Sopra1
Dim Sopra2
Dim Sotto1
Dim Sotto2
If Grid.RowSel > 1 And Grid.Rows > 1 Then
    Sopra1 = Grid.TextMatrix(Grid.RowSel, 0)
    Sopra2 = Grid.TextMatrix(Grid.RowSel, 1)
    Sotto1 = Grid.TextMatrix(Grid.RowSel - 1, 0)
    Sotto2 = Grid.TextMatrix(Grid.RowSel - 1, 1)
    Grid.TextMatrix(Grid.RowSel, 0) = Sotto1
    Grid.TextMatrix(Grid.RowSel, 1) = Sotto2
    Grid.TextMatrix(Grid.RowSel - 1, 0) = Sopra1
    Grid.TextMatrix(Grid.RowSel - 1, 1) = Sopra2
    Grid.RowSel = Grid.RowSel - 1
End If
Grid.Refresh
End Sub


Private Sub CalculateStart()
If Number.Caption > 0 Then Command1.Enabled = True
If Number.Caption > 0 Then Command1.Enabled = True
If Number.Caption = 0 Then Command1.Enabled = False
If Number.Caption = 0 Then Command1.Enabled = False
End Sub
