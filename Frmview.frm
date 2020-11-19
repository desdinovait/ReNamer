VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Properties"
   ClientHeight    =   4170
   ClientLeft      =   1545
   ClientTop       =   1740
   ClientWidth     =   5370
   Icon            =   "Frmview.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "File Properties : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame1 
         Caption         =   "Attributes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   840
         TabIndex        =   10
         Top             =   2160
         Width           =   4215
         Begin VB.CheckBox chkReadOnly 
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox chkHidden 
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox chkSystem 
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox chkArchive 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox chkCompressed 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "Hidden"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "System"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Archive"
            Height          =   255
            Left            =   2160
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Read-Only"
            Height          =   255
            Left            =   480
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Compressed"
            Height          =   255
            Left            =   2160
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "General Info :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         Begin VB.Label lblFileSize 
            Caption         =   "Not found"
            Height          =   255
            Left            =   1200
            TabIndex        =   9
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "File Size :"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblFileName 
            Caption         =   "Not found"
            Height          =   735
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "File Path :"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Modify :"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Not found"
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   1320
            Width           =   2895
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Frmview.frx":000C
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblType 
      Caption         =   "Document - Non cancellare"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 1997 by Desaware Inc. All Rights Reserved

Dim hFile&, TaskID&
'**********************************
'**  Function Declarations:

#If Win32 Then

Private Declare Function GetTimeFormat& Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long)
Private Declare Function GetDateFormat& Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long)
Private Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Private Declare Function CreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long)
Private Declare Function GetFileInformationByHandle& Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION)
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Private Declare Function GetBinaryType& Lib "kernel32" (ByVal szFileName As String, fType As Long)
Private Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal myString As String, ByVal nCount As Long)

#End If 'WIN32


'**********************************
'**  Type Definitions:

#If Win32 Then

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type
#End If 'WIN32 Types

'**********************************
'**  Constant Definitions:

#If Win32 Then
Private Const SCS_32BIT_BINARY& = 0
Private Const SCS_DOS_BINARY& = 1
Private Const SCS_OS216_BINARY& = 5
Private Const SCS_PIF_BINARY& = 3
Private Const SCS_POSIX_BINARY& = 4
Private Const SCS_WOW_BINARY& = 2
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
#End If 'WIN32

Public Sub Initialize(fileName$)
On Error GoTo 10
    Dim dl&, myTime As SYSTEMTIME
    Dim s$, fType&
    Dim myFileInfo As BY_HANDLE_FILE_INFORMATION
    Dim a As Variant
    lblFileName = fileName$
    hFile& = CreateFile&(fileName$, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    
    dl& = GetBinaryType&(fileName$, fType&)
    
    If fType& And SCS_32BIT_BINARY Then
        lblType = "Win32 Application"
    ElseIf fType& And SCS_DOS_BINARY Then
        lblType = "DOS Application"
    ElseIf fType& And SCS_OS216_BINARY Then
        lblType = "OS/2 16-Bit Application"
    ElseIf fType& And SCS_PIF_BINARY Then
        lblType = "DOS PIF Application"
    ElseIf fType& And SCS_WOW_BINARY Then
        lblType = "Win16 Application"
    Else
        lblType = "Document"
    End If
    
    dl& = GetFileInformationByHandle&(hFile&, myFileInfo)
    With myFileInfo
        If .dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE Then chkArchive = 1
        If .dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED Then chkCompressed = 1
        If .dwFileAttributes And FILE_ATTRIBUTE_HIDDEN Then chkHidden = 1
        If .dwFileAttributes And FILE_ATTRIBUTE_READONLY Then chkReadOnly = 1
        If .dwFileAttributes And FILE_ATTRIBUTE_SYSTEM Then chkSystem = 1
        
       ' dl& = FileTimeToSystemTime(.ftCreationTime, myTime)
       ' s$ = String(255, 0)
       ' dl& = GetTimeFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblCreated = Left(s$, dl&)
       ' s$ = String(255, 0)
       ' dl& = GetDateFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblCreated = Left(s$, dl& - 1) & " " & lblCreated
        
       ' dl& = FileTimeToSystemTime(.ftLastAccessTime, myTime)
       ' s$ = String(255, 0)
       ' dl& = GetTimeFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblAccessed = Left$(s$, dl&)
       ' s$ = String(255, 0)
       ' dl& = GetDateFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblAccessed = Left(s$, dl& - 1) & " " & lblAccessed
        
       ' dl& = FileTimeToSystemTime(.ftLastWriteTime, myTime)
       ' s$ = String(255, 0)
       ' dl& = GetTimeFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblModified = Left$(s$, dl&)
       ' s$ = String(255, 0)
       ' dl& = GetDateFormat&(&H800, 0, myTime, 0, s$, 254)
       ' lblModified = Left(s$, dl& - 1) & " " & lblModified
                
        lblFileSize = CStr(.nFileSizeLow) & " bytes"
    End With
    
Label6.Caption = FileDateTime(NomeSelezionato)
    dl& = CloseHandle&(hFile&)
GoTo 20
10 Unload Me
20
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "ReNamer " + Versione + " - File Properties"
End Sub
