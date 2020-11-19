Attribute VB_Name = "ModuloSistema"
Option Explicit
Option Compare Text

'*** CHIAMATA API X LA DIRECTORY DI WINDOWS + ALTRO
'
' Global Constants
'

Global Const gstrSEP_DIR$ = "\"                         ' Directory separator character
Global Const gstrSEP_REGKEY$ = "\"                      ' Registration key separator character.
Global Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Global Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Global Const gstrSEP_PROGID = "."
Global Const gstrSEP_FILE$ = "|"                        ' Use the character for delimiting filename lists because it is not a valid character in a filename.
Global Const gstrSEP_LIST = "|"
Global Const gstrSEP_URL$ = "://"                       ' Separator that follows HPPT in URL address
Global Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Global Const gstrUNC$ = "\\"                            'UNC specifier \\
Global Const gstrCOLON$ = ":"
Global Const gstrSwitchPrefix1 = "-"
Global Const gstrSwitchPrefix2 = "/"
Global Const gstrCOMMA$ = ","
Global Const gstrDECIMAL$ = "."
Global Const gstrQUOTE$ = """"
Global Const gstrCCOMMENT$ = "//"                       ' Comment specifier used in C, etc.
Global Const gstrASSIGN$ = "="
Global Const gstrINI_PROTOCOL = "Protocol"
Global Const gstrREMOTEAUTO = "RA"
Global Const gstrDCOM = "DCOM"

Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Global Const gintMAX_PATH_LEN% = 260                    ' Maximum allowed path length including path, filename,
                                                        ' and command line arguments for NT (Intel) and Win95.
Global Const gintMAX_GROUPNAME_LEN% = 30                ' Maximum length that we allow for an NT 3.51 group name.
Global Const gintMIN_BUTTONWIDTH% = 1200
Global Const gsngBUTTON_BORDER! = 1.4

Global Const intDRIVE_REMOVABLE% = 2                    'Constants for GetDriveType
Global Const intDRIVE_FIXED% = 3
Global Const intDRIVE_REMOTE% = 4
Global Const intDRIVE_CDROM% = 5
Global Const intDRIVE_RAMDISK% = 6

Global Const gintNOVERINFO% = 32767                     'flag indicating no version info

'File names
Global Const gstrFILE_SETUP$ = "SETUP.LST"              'Name of setup information file
Global Const gstrTEMP_DIR$ = "TEMP"
Global Const gstrTMP_DIR$ = "TMP"

'Share type macros for files
Global Const mstrPRIVATEFILE = ""
Global Const mstrSHAREDFILE = "$(Shared)"

'INI File keys
Global Const gstrINI_SETUP$ = "Setup"
Global Const gstrINI_BOOT$ = "Bootstrap"
Global Const gstrINI_APPNAME$ = "Title"
Global Const gstrINI_APPDIR$ = "DefaultDir"
Global Const gstrINI_APPEXE$ = "AppExe"
Global Const gstrINI_APPTOUNINSTALL = "AppToUninstall"
Global Const gstrINI_APPPATH$ = "AppPath"
Global Const gstrINI_FORCEUSEDEFDEST = "ForceUseDefDir"
Global Const gstrINI_DEFGROUP$ = "DefProgramGroup"
Global Const gstrINI_CABNAME$ = "CabFile"

Global Const gstrEXT_DEP$ = "DEP"

'Setup information file macros
Global Const gstrAPPDEST$ = "$(AppPath)"
Global Const gstrWINDEST$ = "$(WinPath)"
Global Const gstrWINSYSDEST$ = "$(WinSysPath)"
Global Const gstrWINSYSDESTSYSFILE$ = "$(WinSysPathSysFile)"
Global Const gstrPROGRAMFILES$ = "$(ProgramFiles)"
Global Const gstrCOMMONFILES$ = "$(CommonFiles)"
Global Const gstrCOMMONFILESSYS$ = "$(CommonFilesSys)"
Global Const gstrDAODEST$ = "$(MSDAOPath)"
Global Const gstrDONOTINSTALL$ = "$(DoNotInstall)"

'Mouse Pointer Constants
Global Const gintMOUSE_DEFAULT% = 0

'MsgError() Constants
Global Const MSGERR_ERROR = 1
Global Const MSGERR_WARNING = 2

'Shell Constants
Global Const NORMAL_PRIORITY_CLASS      As Long = &H20&
Global Const INFINITE                   As Long = -1&

Global Const STATUS_WAIT_0              As Long = &H0
Global Const STATUS_ABANDONED_WAIT_0    As Long = &H80
Global Const STATUS_USER_APC            As Long = &HC0
Global Const STATUS_TIMEOUT             As Long = &H102
Global Const STATUS_PENDING             As Long = &H103

Global Const WAIT_FAILED                As Long = &HFFFFFFFF
Global Const WAIT_OBJECT_0              As Long = STATUS_WAIT_0
Global Const WAIT_TIMEOUT               As Long = STATUS_TIMEOUT

Global Const WAIT_ABANDONED             As Long = STATUS_ABANDONED_WAIT_0
Global Const WAIT_ABANDONED_0           As Long = STATUS_ABANDONED_WAIT_0

Global Const WAIT_IO_COMPLETION         As Long = STATUS_USER_APC
Global Const STILL_ACTIVE               As Long = STATUS_PENDING




Public Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # vblf & vblf DWord
    nLSHi As Integer                                    'High word of ver # vblf & vblf DWord
    strPad3(1 To 16) As Byte                            'Skip some of VERINFO struct (16 bytes)
    FileOS As Long                                      'Information about the OS this file is targeted for.
    strPad4(1 To 16) As Byte                            'Pad out the resto of VERINFO struct (16 bytes)
End Type

Type PROTOCOL
    strName As String
    strFriendlyName As String
End Type

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Global Const OF_EXIST& = &H4000&
Global Const OF_SEARCH& = &H400&
Global Const HFILE_ERROR% = -1

'
' Global variables used for silent and SMS installation
'
Public gfSilent As Boolean                              ' Whether or not we are doing a silent install
Public gstrSilentLog As String                          ' filename for output during silent install.
Public gfSMS As Boolean                                 ' Whether or not we are doing an SMS silent install
Public gstrMIFFile As String                            ' status output file for SMS
Public gfSMSStatus As Boolean                           ' status of SMS installation
Public gstrSMSDescription As String                     ' description string written to MIF file for SMS installation
Public gfNoUserInput As Boolean                         ' True if either gfSMS or gfSilent is True
Public gfDontLogSMS As Boolean                          ' Prevents MsgFunc from being logged to SMS (e.g., for confirmation messasges)
Public ImgX As Integer
Public ImgY As Integer
Global Const MAX_SMS_DESCRIP = 255                      ' SMS does not allow description strings longer than 255 chars.
'
'List of available protocols
'
Global gProtocol() As PROTOCOL
Global gcProtocols As Integer
'
' AXDist.exe and wint351.exe needed.  These are self extracting exes
' that install other files not installed by setup1.
'
Public gfAXDist As Boolean
Global Const gstrFILE_AXDIST = "AXDIST.EXE"
Public gstrAXDISTInstallPath As String
Public gfAXDistChecked As Boolean
Public gfMDag As Boolean
Global Const gstrFILE_MDAG = "mdac_typ.exe"
Global Const gstrFILE_MDAGARGS = " /q:a /c:""setup.exe /q0"""
Public gstrMDagInstallPath As String
Public gfWINt351 As Boolean
Global Const gstrFILE_WINT351 = "WINt351.EXE"
Public gstrWINt351InstallPath As String
Public gfWINt351Checked As Boolean
'
'API/DLL Declarations for 32 bit SetupToolkit
'
Declare Function SetTime Lib "vb6stkit.dll" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Integer
Declare Function DLLSelfRegister Lib "vb6stkit.dll" (ByVal lpDllName As String) As Integer
Declare Function RegisterTLB Lib "vb6stkit.dll" (ByVal lpTLBName As String) As Integer
Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function OSfRemoveShellLink Lib "vb6stkit.dll" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function InputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFilename As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Declare Function GetTempFilename32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFilename As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Global Const LB_FINDSTRINGEXACT = &H1A2
Global Const LB_ERR = (-1)

Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long



'----------------------------------------------------------
' FUNCTION: GetWinPlatform
' Get the current windows platform.
' ---------------------------------------------------------
Public Function GetWinPlatform() As Long
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
'-----------------------------------------------------------
' SUB: AddURLDirSep
' Add a trailing URL path separator (forward slash) to the
' end of a URL unless one (or a back slash) already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddURLDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = Trim(strPathName) & gstrSEP_URLDIR
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    ' If the string is quoted, remove the quotes.
    '
    strPathName = strUnQuoteString(strPathName)
    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: DirExists
'
' Determines whether the specified directory name exists.
' This function is used (for example) to determine whether
' an installation floppy is in the drive by passing in
' something like 'A:\'.
'
' IN: [strDirName] - name of directory to check for
'
' Returns: True if the directory exists, False otherwise
'-----------------------------------------------------------
'
Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    AddDirSep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetDriveType
' Determine whether a disk is fixed, removable, etc. by
' calling Windows GetDriveType()
'-----------------------------------------------------------
'
Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrSEP_DRIVE & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
End Function


'-----------------------------------------------------------
' FUNCTION: ResolveResString
' Reads resource and replaces given macros with given values
'
' Example, given a resource number 14:
'    "Could not read '|1' in drive |2"
'   The call
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   would return the string
'     "Could not read 'TXTFILE.TXT' in drive A:"
'
' IN: [resID] - resource identifier
'     [varReplacements] - pairs of macro/replacement value
'-----------------------------------------------------------
'
Public Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function
'-----------------------------------------------------------
' SUB: GetLicInfoFromVBL
' Parses a VBL file name and extracts the license key for
' the registry and license information.
'
' IN: [strVBLFile] - must be a valid VBL.
'
' OUT: [strLicKey] - registry key to write license info to.
'                    This key will be added to
'                    HKEY_CLASSES_ROOT\Licenses.  It is a
'                    guid.
' OUT: [strLicVal] - license information.  Usually in the
'                    form of a string of cryptic characters.
'-----------------------------------------------------------
'
Public Sub GetLicInfoFromVBL(strVBLFile As String, strLicKey As String, strLicVal As String)
    Dim fn As Integer
    Const strREGEDIT = "REGEDIT"
    Const strLICKEYBASE = "HKEY_CLASSES_ROOT\Licenses\"
    Dim strTemp As String
    Dim posEqual As Integer
    Dim fLicFound As Boolean
    
    fn = FreeFile
    Open strVBLFile For Input Access Read Lock Read Write As #fn
    '
    ' Read through the file until we find a line that starts with strLICKEYBASE
    '
    fLicFound = False
    Do While Not EOF(fn)
        Line Input #fn, strTemp
        strTemp = Trim(strTemp)
        If Left$(strTemp, Len(strLICKEYBASE)) = strLICKEYBASE Then
            '
            ' We've got the line we want.
            '
            fLicFound = True
            Exit Do
        End If
    Loop

    Close fn
    
    If fLicFound Then
        '
        ' Parse the data on this line to split out the
        ' key and the license info.  The line should be
        ' the form of:
        ' "HKEY_CLASSES_ROOT\Licenses\<lickey> = <licval>"
        '
        posEqual = InStr(strTemp, gstrASSIGN)
        If posEqual > 0 Then
            strLicKey = Mid$(Trim(Left$(strTemp, posEqual - 1)), Len(strLICKEYBASE) + 1)
            strLicVal = Trim(Mid$(strTemp, posEqual + 1))
        End If
    Else
        strLicKey = vbNullString
        strLicVal = vbNullString
    End If
End Sub

 
 '-----------------------------------------------------------
 ' FUNCTION GetShortPathName
 '
 ' Retrieve the short pathname version of a path possibly
 '   containing long subdirectory and/or file names
 '-----------------------------------------------------------
 '
 Function GetShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String
     Dim lResult As Long

     On Error GoTo 0
     strShortPath = String(cchBuffer, Chr$(0))
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         'Error 53 ' File not found
         GetShortPathName = vbNullString
     Else
         GetShortPathName = StripTerminator(strShortPath)
     End If
 End Function
 
'-----------------------------------------------------------
' FUNCTION: GetTempFilename
' Get a temporary filename for a specified drive and
' filename prefix
' PARAMETERS:
'   strDestPath - Location where temporary file will be created.  If this
'                 is an empty string, then the location specified by the
'                 tmp or temp environment variable is used.
'   lpPrefixString - First three characters of this string will be part of
'                    temporary file name returned.
'   wUnique - Set to 0 to create unique filename.  Can also set to integer,
'             in which case temp file name is returned with that integer
'             as part of the name.
'   lpTempFilename - Temporary file name is returned as this variable.
' RETURN:
'   True if function succeeds; false otherwise
'-----------------------------------------------------------
'
Function GetTempFilename(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
    If strDestPath = vbNullString Then
        '
        ' No destination was specified, use the temp directory.
        '
        strDestPath = String(gintMAX_PATH_LEN, vbNullChar)
        If GetTempPath(gintMAX_PATH_LEN, strDestPath) = 0 Then
            GetTempFilename = False
            Exit Function
        End If
    End If
    lpTempFilename = String(gintMAX_PATH_LEN, vbNullChar)
    GetTempFilename = GetTempFilename32(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
'-----------------------------------------------------------
' FUNCTION: GetDefMsgBoxButton
' Decode the flags passed to the MsgBox function to
' determine what the default button is.  Use this
' for silent installs.
'
' IN: [intFlags] - Flags passed to MsgBox
'
' Returns: VB defined number for button
'               vbOK        1   OK button pressed.
'               vbCancel    2   Cancel button pressed.
'               vbAbort     3   Abort button pressed.
'               vbRetry     4   Retry button pressed.
'               vbIgnore    5   Ignore button pressed.
'               vbYes       6   Yes button pressed.
'               vbNo        7   No button pressed.
'-----------------------------------------------------------
'
Function GetDefMsgBoxButton(intFlags) As Integer
    '
    ' First determine the ordinal of the default
    ' button on the message box.
    '
    Dim intButtonNum As Integer
    Dim intDefButton As Integer
    
    If (intFlags And vbDefaultButton2) = vbDefaultButton2 Then
        intButtonNum = 2
    ElseIf (intFlags And vbDefaultButton3) = vbDefaultButton3 Then
        intButtonNum = 3
    Else
        intButtonNum = 1
    End If
    '
    ' Now determine the type of message box we are dealing
    ' with and return the default button.
    '
    If (intFlags And vbRetryCancel) = vbRetryCancel Then
        intDefButton = IIf(intButtonNum = 1, vbRetry, vbCancel)
    ElseIf (intFlags And vbYesNoCancel) = vbYesNoCancel Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbYes
            Case 2
                intDefButton = vbNo
            Case 3
                intDefButton = vbCancel
            'End Case
        End Select
    ElseIf (intFlags And vbOKCancel) = vbOKCancel Then
        intDefButton = IIf(intButtonNum = 1, vbOK, vbCancel)
    ElseIf (intFlags And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbAbort
            Case 2
                intDefButton = vbRetry
            Case 3
                intDefButton = vbIgnore
            'End Case
        End Select
    ElseIf (intFlags And vbYesNo) = vbYesNo Then
        intDefButton = IIf(intButtonNum = 1, vbYes, vbNo)
    Else
        intDefButton = vbOK
    End If
    
    GetDefMsgBoxButton = intDefButton
    
End Function

'-----------------------------------------------------------
' FUNCTION: GetUNCShareName
'
' Given a UNC names, returns the leftmost portion of the
' directory representing the machine name and share name.
' E.g., given "\\SCHWEIZ\PUBLIC\APPS\LISTING.TXT", returns
' the string "\\SCHWEIZ\PUBLIC"
'
' Returns a string representing the machine and share name
'   if the path is a valid pathname, else returns NULL
'-----------------------------------------------------------
'
Function GetUNCShareName(ByVal strFN As String) As Variant
    GetUNCShareName = Null
    If IsUNCName(strFN) Then
        Dim iFirstSeparator As Integer
        iFirstSeparator = InStr(3, strFN, gstrSEP_DIR)
        If iFirstSeparator > 0 Then
            Dim iSecondSeparator As Integer
            iSecondSeparator = InStr(iFirstSeparator + 1, strFN, gstrSEP_DIR)
            If iSecondSeparator > 0 Then
                GetUNCShareName = Left$(strFN, iSecondSeparator - 1)
            Else
                GetUNCShareName = strFN
            End If
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetWindowsSysDir
'
' Calls the windows API to get the windows\SYSTEM directory
' and ensures that a trailing dir separator is present
'
' Returns: The windows\SYSTEM directory
'-----------------------------------------------------------
'
Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = vbNullString
    End If
End Function
'-----------------------------------------------------------
' SUB: TreatAsWin95
'
' Returns True iff either we're running under Windows 95
' or we are treating this version of NT as if it were
' Windows 95 for registry and application loggin and
' removal purposes.
'-----------------------------------------------------------
'
Function TreatAsWin95() As Boolean
    If IsWindows95() Then
        TreatAsWin95 = True
    ElseIf NTWithShell() Then
        TreatAsWin95 = True
    Else
        TreatAsWin95 = False
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: NTWithShell
'
' Returns true if the system is on a machine running
' NT4.0 or greater.
'-----------------------------------------------------------
'
Function NTWithShell() As Boolean

    If Not IsWindowsNT() Then
        NTWithShell = False
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0 or higher?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion >= NT4MajorVersion) Then
        NTWithShell = False
    Else
        NTWithShell = True
    End If
    
End Function
'-----------------------------------------------------------
' FUNCTION: IsDepFile
'
' Returns true if the file passed to this routine is a
' dependency (*.dep) file.  We make this determination
' by verifying that the extension is .dep and that it
' contains version information.
'-----------------------------------------------------------
'
Function fIsDepFile(strFilename As String) As Boolean
    Const strEXT_DEP = "DEP"
    
    fIsDepFile = False
    
    If UCase(Extension(strFilename)) = strEXT_DEP Then
        If GetFileVersion(strFilename) <> vbNullString Then
            fIsDepFile = True
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsWin32
'
' Returns true if this program is running under Win32 (i.e.
'   any 32-bit operating system)
'-----------------------------------------------------------
'
Function IsWin32() As Boolean
    IsWin32 = (IsWindows95() Or IsWindowsNT())
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindows95
'
' Returns true if this program is running under Windows 95
'   or successor
'-----------------------------------------------------------
'
Function IsWindows95() As Boolean
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT
'
' Returns true if this program is running under Windows NT
'-----------------------------------------------------------
'
Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT4WithoutSP2
'
' Determines if the user is running under Windows NT 4.0
' but without Service Pack 2 (SP2).  If running under any
' other platform, returns False.
'
' IN: [none]
'
' Returns: True if and only if running under Windows NT 4.0
' without at least Service Pack 2 installed.
'-----------------------------------------------------------
'
Function IsWindowsNT4WithoutSP2() As Boolean
    IsWindowsNT4WithoutSP2 = False
    
    If Not IsWindowsNT() Then
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion <> NT4MajorVersion) Or (osvi.dwMinorVersion <> NT4MinorVersion) Then
        'No.  Return False.
        Exit Function
    End If
    
    'If no service pack is installed, or if Service Pack 1 is
    'installed, then return True.
    Const strSP1 = "SERVICE PACK 1"
    If strCSDVersion = "" Then
        IsWindowsNT4WithoutSP2 = True 'No service pack installed
    ElseIf strCSDVersion = strSP1 Then
        IsWindowsNT4WithoutSP2 = True 'Only SP1 installed
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsUNCName
'
' Determines whether the pathname specified is a UNC name.
' UNC (Universal Naming Convention) names are typically
' used to specify machine resources, such as remote network
' shares, named pipes, etc.  An example of a UNC name is
' "\\SERVER\SHARE\FILENAME.EXT".
'
' IN: [strPathName] - pathname to check
'
' Returns: True if pathname is a UNC name, False otherwise
'-----------------------------------------------------------
'
Function IsUNCName(ByVal strPathName As String) As Integer
    Const strUNCNAME$ = "\\//\"        'so can check for \\, //, \/, /\

    IsUNCName = ((InStr(strUNCNAME, Left$(strPathName, 2)) > 0) And _
                 (Len(strPathName) > 1))
End Function
'-----------------------------------------------------------
' FUNCTION: LogSilentMsg
'
' If this is a silent install, this routine writes
' a message to the gstrSilentLog file.
'
' IN: [strMsg] - The message
'
' Normally, this routine is called inlieu of displaying
' a MsgBox and strMsg is the same message that would
' have appeared in the MsgBox

'-----------------------------------------------------------
'
Sub LogSilentMsg(strMsg As String)
    If Not gfSilent Then Exit Sub
    
    Dim fn As Integer
    
    On Error Resume Next
    
    fn = FreeFile
    
    Open gstrSilentLog For Append As fn
    Print #fn, strMsg
    Close fn
    Exit Sub
End Sub
'-----------------------------------------------------------
' FUNCTION: LogSMSMsg
'
' If this is a SMS install, this routine appends
' a message to the gstrSMSDescription string.  This
' string will later be written to the SMS status
' file (*.MIF) when the installation completes (success
' or failure).
'
' Note that if gfSMS = False, not message will be logged.
' Therefore, to prevent some messages from being logged
' (e.g., confirmation only messages), temporarily set
' gfSMS = False.
'
' IN: [strMsg] - The message
'
' Normally, this routine is called inlieu of displaying
' a MsgBox and strMsg is the same message that would
' have appeared in the MsgBox
'-----------------------------------------------------------
'
Sub LogSMSMsg(strMsg As String)
    If Not gfSMS Then Exit Sub
    '
    ' Append the message.  Note that the total
    ' length cannot be more than 255 characters, so
    ' truncate anything after that.
    '
    gstrSMSDescription = Left(gstrSMSDescription & strMsg, MAX_SMS_DESCRIP)
End Sub

'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
#If LOGGING Then
                NewAction gstrKEY_CREATEDIR, """" & strPath & """"
#End If
                MkDir strPath
#If LOGGING Then
                If Err Then
                    LogError ResolveResString(resMAKEDIR) & " " & strPath
                    AbortAction
                    GoTo Done
                Else
                    CommitAction
                End If
#End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: MsgError
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [intLogType] (optional) - The type of logfile entry to make.
'                   By default, creates an error entry.  Use
'                   the MsgWarning() function to create a warning.
'                   Valid types as MSGERR_ERROR and MSGERR_WARNING
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgError(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String, Optional ByVal intLogType As Variant) As Integer
    Dim iRet As Integer
    
    iRet = MsgFunc(strMsg, intFlags, strCaption)
    MsgError = iRet
#If LOGGING Then
    ' We need to log this error and decode the user's response.
    Dim strID As String
    Dim strLogMsg As String

    Select Case iRet
        Case vbOK
            strID = ResolveResString(resLOG_vbok)
        Case vbCancel
            strID = ResolveResString(resLOG_vbCancel)
        Case vbAbort
            strID = ResolveResString(resLOG_vbabort)
        Case vbRetry
            strID = ResolveResString(resLOG_vbretry)
        Case vbIgnore
            strID = ResolveResString(resLOG_vbignore)
        Case vbYes
            strID = ResolveResString(resLOG_vbyes)
        Case vbNo
            strID = ResolveResString(resLOG_vbno)
        Case Else
            strID = ResolveResString(resLOG_IDUNKNOWN)
        'End Case
    End Select

    strLogMsg = strMsg & vbLf & "(" & ResolveResString(resLOG_USERRESPONDEDWITH, "|1", strID) & ")"
    If IsMissing(intLogType) Then
        intLogType = MSGERR_ERROR
    End If
    On Error Resume Next
    Select Case intLogType
        Case MSGERR_WARNING
            LogWarning strLogMsg
        Case MSGERR_ERROR
            LogError strLogMsg
        Case Else
            LogError strLogMsg
        'End Case
    End Select
#End If
End Function

'-----------------------------------------------------------
' FUNCTION: MsgFunc
'
' Forces mouse pointer to default and calls VB's MsgBox
' function.  See also MsgError.
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgFunc(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    Dim intOldPointer As Integer
  
    intOldPointer = Screen.MousePointer
    If gfNoUserInput Then
        MsgFunc = GetDefMsgBoxButton(intFlags)
        If gfSilent = True Then
            LogSilentMsg strMsg
        End If
        If gfSMS = True Then
            LogSMSMsg strMsg
            gfDontLogSMS = False
        End If
    Else
        Screen.MousePointer = gintMOUSE_DEFAULT
        MsgFunc = MsgBox(strMsg, intFlags, strCaption)
        Screen.MousePointer = intOldPointer
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: MsgWarning
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgWarning(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    MsgWarning = MsgError(strMsg, intFlags, strCaption, MSGERR_WARNING)
End Function

'-----------------------------------------------------------
' SUB: SetMousePtr
'
' Provides a way to set the mouse pointer only when the
' pointer state changes.  For every HOURGLASS call, there
' should be a corresponding DEFAULT call.  Other types of
' mouse pointers are set explicitly.
'
' IN: [intMousePtr] - type of mouse pointer desired
'-----------------------------------------------------------
'
Sub SetMousePtr(intMousePtr As Integer)
    Static intPtrState As Integer

    Select Case intMousePtr
        Case vbHourglass
            intPtrState = intPtrState + 1
        Case gintMOUSE_DEFAULT
            intPtrState = intPtrState - 1
            If intPtrState < 0 Then
                intPtrState = 0
            End If
        Case Else
            Screen.MousePointer = intMousePtr
            Exit Sub
        'End Case
    End Select

    Screen.MousePointer = IIf(intPtrState > 0, vbHourglass, gintMOUSE_DEFAULT)
End Sub

'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVersion
'
' Returns the internal file version number for the specified
' file.  This can be different than the 'display' version
' number shown in the File Manager File Properties dialog.
' It is the same number as shown in the VB5 SetupWizard's
' File Details screen.  This is the number used by the
' Windows VerInstallFile API when comparing file versions.
'
' IN: [strFilename] - the file whose version # is desired
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote ActiveX component support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
'
' Returns: The Version number string if found, otherwise
'          vbnullstring
'-----------------------------------------------------------
'
Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
    Dim sVerInfo As VERINFO
    Dim strVer As String

    On Error GoTo GFVError

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If
    
    '
    'Get the file version into a VERINFO struct, and then assemble a version string
    'from the appropriate elements.
    '
    If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
        strVer = Format$(sVerInfo.nMSHi) & gstrDECIMAL & Format$(sVerInfo.nMSLo) & gstrDECIMAL
        strVer = strVer & Format$(sVerInfo.nLSHi) & gstrDECIMAL & Format$(sVerInfo.nLSLo)
        GetFileVersion = strVer
    Else
        GetFileVersion = vbNullString
    End If
    
    Exit Function
    
GFVError:
    GetFileVersion = vbNullString
    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVerStruct
'
' Gets the file version information into a VERINFO TYPE
' variable
'
' IN: [strFilename] - name of file to get version info for
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote ActiveX component support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strFIXEDFILEINFO$ = "\"

    Dim lVerSize As Long
    Dim lVerHandle As Long
    Dim lpBufPtr As Long
    Dim byteVerData() As Byte
    Dim fFoundVer As Boolean

    GetFileVerStruct = False
    fFoundVer = False

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If
    
    If fIsRemoteServerSupportFile Then
        GetFileVerStruct = GetRemoteSupportFileVerStruct(strFilename, sVerInfo)
        fFoundVer = True
    Else
        '
        'Get the size of the file version info, allocate a buffer for it, and get the
        'version info.  Next, we query the Fixed file info portion, where the internal
        'file version used by the Windows VerInstallFile API is kept.  We then copy
        'the fixed file info into a VERINFO structure.
        '
        lVerSize = GetFileVersionInfoSize(strFilename, lVerHandle)
        If lVerSize > 0 Then
            ReDim byteVerData(lVerSize)
            If GetFileVersionInfo(strFilename, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
                If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
                    CopyMemory sVerInfo, lpBufPtr, lVerSize
                    fFoundVer = True
                    GetFileVerStruct = True
                End If
            End If
        End If
    End If
    
    If Not fFoundVer Then
        '
        ' We were unsuccessful in finding the version info from the file.
        ' One possibility is that this is a dependency file.
        '
        If UCase(Extension(strFilename)) = gstrEXT_DEP Then
            GetFileVerStruct = GetDepFileVerStruct(strFilename, sVerInfo)
        End If
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: GetDepFileVerStruct
'
' Gets the file version information from a dependency
' file (*.dep).  Such files do not have a Windows version
' stamp, but they do have an internal version stamp that
' we can look for.
'
' IN: [strFilename] - name of dep file to get version info for
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
Function GetDepFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    GetDepFileVerStruct = False
    
    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            GetDepFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetDepFileVerStruct = False
End Function

'-----------------------------------------------------------
' FUNCTION: GetRemoteSupportFileVerStruct
'
' Gets the file version information of a remote ActiveX component
' support file into a VERINFO TYPE variable (Enterprise
' Edition only).  Such files do not have a Windows version
' stamp, but they do have an internal version stamp that
' we can look for.
'
' IN: [strFilename] - name of file to get version info for
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFilename For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            'Convert the format 1.2.3 from the .VBR into
            '1.2.0.3, which is really want we want
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0
            
            GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetRemoteSupportFileVerStruct = False
End Function
'-----------------------------------------------------------
' FUNCTION: GetWindowsDir
'
' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory
'-----------------------------------------------------------
'
Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = strBuf
    Else
        GetWindowsDir = vbNullString
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: ExtractFilenameItem
'
' Extracts a quoted or unquoted filename from a string.
'
' IN: [str] - string to parse for a filename.
'     [intAnchor] - index in str at which the filename begins.
'             The filename continues to the end of the string
'             or up to the next comma in the string, or, if
'             the filename is enclosed in quotes, until the
'             next double quote.
' OUT: Returns the filename, without quotes.
'      [intAnchor] is set to the comma, or else one character
'             past the end of the string
'      [fErr] is set to True if a parsing error is discovered
'
'-----------------------------------------------------------
'
Function strExtractFilenameItem(ByVal str As String, intAnchor As Integer, fErr As Boolean) As String
    While Mid$(str, intAnchor, 1) = " "
        intAnchor = intAnchor + 1
    Wend
    
    Dim iEndFilenamePos As Integer
    Dim strFilename As String
    If Mid$(str, intAnchor, 1) = """" Then
        ' Filename is surrounded by quotes
        iEndFilenamePos = InStr(intAnchor + 1, str, """") ' Find matching quote
        If iEndFilenamePos > 0 Then
            strFilename = Mid$(str, intAnchor + 1, iEndFilenamePos - 1 - intAnchor)
            intAnchor = iEndFilenamePos + 1
            While Mid$(str, intAnchor, 1) = " "
                intAnchor = intAnchor + 1
            Wend
            If (Mid$(str, intAnchor, 1) <> gstrCOMMA) And (Mid$(str, intAnchor, 1) <> "") Then
                fErr = True
                Exit Function
            End If
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' Filename continues until next comma or end of string
        Dim iCommaPos As Integer
        
        iCommaPos = InStr(intAnchor, str, gstrCOMMA)
        If iCommaPos = 0 Then
            iCommaPos = Len(str) + 1
        End If
        iEndFilenamePos = iCommaPos
        
        strFilename = Mid$(str, intAnchor, iEndFilenamePos - intAnchor)
        intAnchor = iCommaPos
    End If
    
    strFilename = Trim$(strFilename)
    If strFilename = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameItem = strFilename
End Function

'-----------------------------------------------------------
' FUNCTION: Extension
'
' Extracts the extension portion of a file/path name
'
' IN: [strFilename] - file/path to get the extension of
'
' Returns: The extension if one exists, else vbnullstring
'-----------------------------------------------------------
'
Function Extension(ByVal strFilename As String) As String
    Dim intPos As Integer

    Extension = vbNullString

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_EXT
                Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Exit Do
            'End Case
        End Select

        intPos = intPos - 1
    Loop
End Function

'-----------------------------------------------------------
' SUB: PackVerInfo
'
' Parses a file version number string of the form
' x[.x[.x[.x]]] and assigns the extracted numbers to the
' appropriate elements of a VERINFO type variable.
' Examples of valid version strings are '3.11.0.102',
' '3.11', '3', etc.
'
' IN: [strVersion] - version number string
'
' OUT: [sVerInfo] - VERINFO type variable whose elements
'                   are assigned the appropriate numbers
'                   from the version number string
'-----------------------------------------------------------
'
Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
    Dim intOffset As Integer
    Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub

Public Function strQuoteString(strUnQuotedString As String, Optional vForce As Variant, Optional vTrim As Variant)
'
' This routine adds quotation marks around an unquoted string, by default.  If the string is already quoted
' it returns without making any changes unless vForce is set to True (vForce defaults to False) except that white
' space before and after the quotes will be removed unless vTrim is False.  If the string contains leading or
' trailing white space it is trimmed unless vTrim is set to False (vTrim defaults to True).
'
    Dim strQuotedString As String
    
    If IsMissing(vForce) Then
        vForce = False
    End If
    If IsMissing(vTrim) Then
        vTrim = True
    End If
    
    strQuotedString = strUnQuotedString
    '
    ' Trim the string if necessary
    '
    If vTrim = True Then
        strQuotedString = Trim(strQuotedString)
    End If
    '
    ' See if the string is already quoted
    '
    If vForce = False Then
        If (Left(strQuotedString, 1) = gstrQUOTE) And (Right(strQuotedString, 1) = gstrQUOTE) Then
            '
            ' String is already quoted.  We are done.
            '
            GoTo DoneQuoteString
        End If
    End If
    '
    ' Add the quotes
    '
    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
DoneQuoteString:
    strQuoteString = strQuotedString
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function
Public Function fCheckFNLength(strFilename As String) As Boolean
'
' This routine verifies that the length of the filename strFilename is valid.
' Under NT (Intel) and Win95 it can be up to 259 (gintMAX_PATH_LEN-1) characters
' long.  This length must include the drive, path, filename, commandline
' arguments and quotes (if the string is quoted).
'
    fCheckFNLength = (Len(strFilename) < gintMAX_PATH_LEN)
End Function
Public Function intGetNextFldOffset(ByVal intAnchor As Integer, strList As String, strDelimit As String, Optional CompareType As Variant) As Integer
'
' This routine reads from a strDelimit separated list, strList, and locates the next
' item in the list following intAnchor.  Basically it finds the next
' occurance of strDelimit that is not inside quotes.  If strDelimit is not
' found the routine returns 0.  Note intAnchor must be outside of quotes
' or this routine will return incorrect results.
'
' strDelimit is typically a comma.
'
' If there is an error this routine returns -1.
'
    Dim intQuote As Integer
    Dim intDelimit As Integer
    
    Const CompareBinary = 0
    Const CompareText = 1

    If IsMissing(CompareType) Then
        CompareType = CompareText
    End If
    
    If intAnchor = 0 Then intAnchor = 1
    
    intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
    intDelimit = InStr(intAnchor, strList, strDelimit, CompareType)
    
    If (intQuote > intDelimit) Or (intQuote = 0) Then
        '
        ' The next delimiter is not within quotes.  Therefore,
        ' we have found what we are looking for.  Note that the
        ' case where there are no delimiters is also handled here.
        '
        GoTo DoneGetNextFldOffset
    ElseIf intQuote < intDelimit Then
        '
        ' A quote appeared before the next delimiter.  This
        ' means we might be inside quotes.  We still need to check
        ' if the closing quote comes after the delmiter or not.
        '
        intAnchor = intQuote + 1
        intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
        If (intQuote > intDelimit) Then
            '
            ' The delimiter was inside quotes.  Therefore, ignore it.
            ' The next delimiter after the closing quote must be outside
            ' of quotes or else we have a corrupt file.
            '
            intAnchor = intQuote + 1
            intDelimit = InStr(intAnchor, strList, strDelimit, CompareType)
            '
            ' Sanity check.  Make sure there is not another quote before
            ' the delimiter we just found.
            '
            If intDelimit > 0 Then
                intQuote = InStr(intAnchor, strList, gstrQUOTE, CompareType)
                If (intQuote > 0) And (intQuote < intDelimit) Then
                    '
                    ' Something is wrong.  We've encountered a stray
                    ' quote.  Means the string is probably corrupt.
                    '
                    intDelimit = -1 ' Error
                End If
            End If
        End If
    End If
DoneGetNextFldOffset:
    intGetNextFldOffset = intDelimit
End Function


Public Function StringFromBuffer(Buffer As String) As String
    Dim nPos As Long

    nPos = InStr(Buffer, Chr$(0))
    If nPos > 0 Then
        StringFromBuffer = Left$(Buffer, nPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function

''==============================================================================
''Code flow routines:

Public Function SyncShell(CommandLine As String, Optional Timeout As Long, _
    Optional WaitForInputIdle As Boolean, Optional Hide As Boolean = False) As Boolean

    Dim hProcess As Long

    Const STARTF_USESHOWWINDOW As Long = &H1
    Const SW_HIDE As Long = 0
    
    Dim ret As Long
    Dim nMilliseconds As Long

    If Timeout > 0 Then
        nMilliseconds = Timeout
    Else
        nMilliseconds = INFINITE
    End If

    hProcess = StartProcess(CommandLine, Hide)

    If WaitForInputIdle Then
        'Wait for the shelled application to finish setting up its UI:
        ret = InputIdle(hProcess, nMilliseconds)
    Else
        'Wait for the shelled application to terminate:
        ret = WaitForSingleObject(hProcess, nMilliseconds)
    End If

    CloseHandle hProcess

    'Return True if the application finished. Otherwise it timed out or erred.
    SyncShell = (ret = WAIT_OBJECT_0)
End Function

Public Function StartProcess(CommandLine As String, Optional Hide As Boolean = False) As Long
    Const STARTF_USESHOWWINDOW As Long = &H1
    Const SW_HIDE As Long = 0
    
    Dim proc As PROCESS_INFORMATION
    Dim Start As STARTUPINFO

    'Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    If Hide Then
        Start.dwFlags = STARTF_USESHOWWINDOW
        Start.wShowWindow = SW_HIDE
    End If
    'Start the shelled application:
    CreateProcessA 0&, CommandLine, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc

    StartProcess = proc.hProcess
End Function



