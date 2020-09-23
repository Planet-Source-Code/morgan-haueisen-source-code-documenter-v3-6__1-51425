VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//*************************************/
'//     Author: Morgan Haueisen        */
'//             morganh@hartcom.net    */
'//     Copyright (c) 1996-2004        */
'//*************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

Private Type SHFILEOPSTRUCT
   hWnd                  As Long
   wFunc                 As Long
   pFrom                 As String
   pTo                   As String
   fFlags                As Integer
   fAnyOperationsAborted As Boolean
   hNameMappings         As Long
   lpszProgressTitle     As String '//  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
      (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_MOVE               As Long = &H1
Private Const FO_COPY               As Long = &H2
Private Const FO_DELETE             As Long = &H3
Private Const FO_RENAME             As Long = &H4

Private Const FOF_MULTIDESTFILES    As Long = &H1
Private Const FOF_SILENT            As Long = &H4 '//  don't create progress/report
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION    As Long = &H10 '//  Don't prompt the user.
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20 '//  Fill in SHFILEOPSTRUCT.hNameMappings
Private Const FOF_ALLOWUNDO         As Long = &H40
Private Const FOF_FILESONLY         As Long = &H80 '//  on *.*, do only files
Private Const FOF_SIMPLEPROGRESS    As Long = &H100 '//  means don't show names of Files
Private Const FOF_NOCONFIRMMKDIR    As Long = &H200 '//  don't confirm making any needed dirs

'// This API returns The Attrbiutes of the File
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
      (ByVal lpFileName As String) As Long

Private mlngNoConfirmation As Long

'// Examples:
'// VBGetOpenFileName(sFile, , True, , , , "All Icon Files (*.EXE;*.DLL;*.ICO)|*.EXE;*.DLL;*.ICO|Icon" & " Files
'  (*.ICO)|*.ICO|Executables (*.EXE;*.DLL)|*.EXE;*.DLL|All Files (*.*)|*.*", 1, , , "ICO", Me.hWnd)

'// VBGetSaveFileName(sFile, , , "ICO Files (*.ICO)|*.ICO|All Files (*.*)|*.*", , , , "ICO", Me.hWnd)

Private Const C_lngMAX_PATH        As Long = 260&
Private Const C_lngMAX_FILE        As Long = 260&

Private Type OPENFILENAME
   lStructSize                    As Long    '// Filled with UDT size
   hWndOwner                      As Long    '// Tied to Owner
   hInstance                      As Long    '// Ignored (used only by templates)
   lpstrFilter                    As String '// Tied to Filter
   lpstrCustomFilter              As String '// Ignored (exercise for reader)
   nMaxCustFilter                 As Long    '// Ignored (exercise for reader)
   nFilterIndex                   As Long    '// Tied to FilterIndex
   lpstrFile                      As String '// Tied to FileName
   nMaxFile                       As Long    '// Handled internally
   lpstrFileTitle                 As String '// Tied to FileTitle
   nMaxFileTitle                  As Long    '// Handled internally
   lpstrInitialDir                As String '// Tied to InitDir
   lpstrTitle                     As String '// Tied to DlgTitle
   Flags                          As Long    '// Tied to Flags
   nFileOffset                    As Integer '// Ignored (exercise for reader)
   nFileExtension                 As Integer '// Ignored (exercise for reader)
   lpstrDefExt                    As String '// Tied to DefaultExt
   lCustData                      As Long    '// Ignored (needed for hooks)
   lpfnHook                       As Long    '// Ignored (good luck with hooks)
   lpTemplateName                 As Long    '// Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFilename Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFilename Lib "COMDLG32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private mlngApiReturn     As Long
Private mlngExtendedError As Long

Public Enum enuOpenFile
   OFN_ReadOnly = &H1
   OFN_OverWritePrompt = &H2
   OFN_HideReadOnly = &H4
   OFN_NOCHANGEDIR = &H8
   OFN_SHOWHELP = &H10
   OFN_ENABLEHOOK = &H20
   OFN_ENABLETEMPLATE = &H40
   OFN_ENABLETEMPLATEHANDLE = &H80
   OFN_NOVALIDATE = &H100
   OFN_ALLOWMULTISELECT = &H200
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_PATHMUSTEXIST = &H800
   OFN_FileMustExist = &H1000
   OFN_CREATEPROMPT = &H2000
   OFN_SHAREAWARE = &H4000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOLONGNAMES = &H40000
   OFN_EXPLORER = &H80000
   OFN_NODEREFERENCELINKS = &H100000
   OFN_LONGNAMES = &H200000
End Enum

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
      (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
      (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Const BIF_STATUSTEXT        As Long = &H4
Private Const BIF_RETURNONLYFSDIRS  As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_NEWDIALOGSTYLE    As Long = &H40
Private Const BIF_USENEWUI          As Long = BIF_NEWDIALOGSTYLE

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
      (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'// Required for changing a File's Date & Time
Private Type FILETIME
   dwLowDateTime    As Long
   dwHighDateTime   As Long
End Type
Private Type SYSTEMTIME
   wYear            As Integer
   wMonth           As Integer
   wDayOfWeek       As Integer
   wDay             As Integer
   wHour            As Integer
   wMinute          As Integer
   wSecond          As Integer
   wMilliseconds    As Integer
End Type
Private Const GENERIC_WRITE    As Long = &H40000000
Private Const OPEN_EXISTING    As Long = 3
Private Const FILE_SHARE_READ  As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
      (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
      ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
      ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
      ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" _
      (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
      lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" _
      (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" _
      (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Sub Class_Initialize()
   
   mlngNoConfirmation = 0
   
End Sub

Public Function CopyDir(ByVal vstrSPath As String, _
                        ByVal vstrDPath As String, _
                        Optional ByVal vblnNoShow As Boolean = False, _
                        Optional ByVal vblnNoConformation As Boolean = False) As Boolean
   
  Dim lngTemp As Long
   
   lngTemp = mlngNoConfirmation
   
   If vblnNoConformation Then
      mlngNoConfirmation = FOF_NOCONFIRMATION
    Else
      mlngNoConfirmation = 0
   End If
   
   If right$(vstrSPath, 1) = "\" Then vstrSPath = left$(vstrSPath, Len(vstrSPath) - 1)
   If right$(vstrDPath, 1) = "\" Then vstrDPath = left$(vstrDPath, Len(vstrDPath) - 1)
   Call CreateDir(vstrDPath)
   
   If LCase$(Me.RetOnlyPath(vstrSPath, False)) = LCase$(vstrDPath) Then
      CopyDir = False
    Else
      CopyDir = CopyMoveSingle(vstrSPath, vstrDPath, FO_COPY, False, vblnNoShow)
   End If
   
   mlngNoConfirmation = lngTemp
   
End Function

Public Function CopyFile(ByVal vstrSPathFile As String, _
                         ByVal vstrDPathFile As String, _
                         Optional ByVal vblnOnlyNewer As Boolean = False, _
                         Optional ByVal vblnNoShow As Boolean = True) As Boolean
   
   If LCase$(vstrSPathFile) = LCase$(vstrDPathFile) Then
      CopyFile = False
    Else
      CopyFile = CopyMoveSingle(vstrSPathFile, vstrDPathFile, FO_COPY, vblnOnlyNewer, vblnNoShow)
   End If
   
End Function

Private Function CopyMoveMulti(ByVal vstrSPath As String, _
                               ByVal vstrDPath As String, _
                               ByRef vntFileName() As String, _
                               ByVal vlngAction As Long, _
                               Optional ByVal vblnOnlyNewer As Boolean = False, _
                               Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
  Dim udtOptions    As SHFILEOPSTRUCT
  Dim lngI          As Long
  Dim strSFileNames As String
  Dim strDFileNames As String
   
   On Local Error GoTo Function_Error
   
   If right$(vstrSPath, 1) <> "\" Then vstrSPath = vstrSPath & "\"
   If right$(vstrDPath, 1) <> "\" Then vstrDPath = vstrDPath & "\"
   
   For lngI = LBound(vntFileName) To UBound(vntFileName)
      If vntFileName(lngI) > vbNullString Then
         If vblnOnlyNewer Then
            If NotTheSame(vstrSPath & vntFileName(lngI), vstrDPath & vntFileName(lngI)) Then
               strSFileNames = strSFileNames & vstrSPath & vntFileName(lngI) & vbNullChar
            End If
          Else
            strSFileNames = strSFileNames & vstrSPath & vntFileName(lngI) & vbNullChar
         End If
      End If
   Next lngI
   
   If strSFileNames > vbNullString Then
      
      strSFileNames = strSFileNames & vbNullChar
      strDFileNames = left$(vstrDPath, Len(vstrDPath) - 1) & vbNullChar
      
      With udtOptions
         .hWnd = 0
         .wFunc = vlngAction
         .pFrom = strSFileNames
         .pTo = strDFileNames
         .fFlags = FOF_ALLOWUNDO
      End With
      
      If vblnNoShow Then udtOptions.fFlags = udtOptions.fFlags Or FOF_SILENT
      If vblnOnlyNewer Then
         udtOptions.fFlags = udtOptions.fFlags Or FOF_NOCONFIRMATION
       Else
         udtOptions.fFlags = udtOptions.fFlags Or mlngNoConfirmation
      End If
      
      CopyMoveMulti = Not CBool(SHFileOperation(udtOptions))
   End If
   
Exit_Function:
   On Local Error GoTo 0
   
   Exit Function
   
Function_Error:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "CopyMoveMulti"
   Err.Clear
   Resume Exit_Function
   
End Function

Private Function CopyMoveSingle(ByVal vstrSPathFile As String, _
                                ByVal vstrDPathFile As String, _
                                ByVal vlngAction As Long, _
                                Optional ByVal vblnOnlyNewer As Boolean = False, _
                                Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
  Dim udtOptions As SHFILEOPSTRUCT
   
   If vblnOnlyNewer Then
      If Not NotTheSame(vstrSPathFile, vstrDPathFile) Then Exit Function
   End If
   
   On Error GoTo Function_Error
   
   With udtOptions
      .hWnd = 0
      .wFunc = vlngAction
      .pFrom = vstrSPathFile & vbNullChar
      .pTo = vstrDPathFile & vbNullChar
      .fFlags = FOF_ALLOWUNDO
   End With
   
   If vblnNoShow Then udtOptions.fFlags = udtOptions.fFlags Or FOF_SILENT
   If vblnOnlyNewer Then
      udtOptions.fFlags = udtOptions.fFlags Or FOF_NOCONFIRMATION
    Else
      udtOptions.fFlags = udtOptions.fFlags Or mlngNoConfirmation
   End If
   
   CopyMoveSingle = Not CBool(SHFileOperation(udtOptions))
   DoEvents
   
Exit_Function:
   On Error GoTo 0
   
   Exit Function
   
Function_Error:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "CopyMoveSingle"
   Err.Clear
   Resume Exit_Function
   
End Function

Public Function CopyMultiFiles(ByVal vstrSPath As String, _
                               ByVal vstrDPath As String, _
                               ByRef vntFileName() As String, _
                               Optional ByVal vblnOnlyNewer As Boolean = False, _
                               Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
   If LCase$(vstrSPath) = LCase$(vstrDPath) Then
      CopyMultiFiles = False
    Else
      CopyMultiFiles = CopyMoveMulti(vstrSPath, vstrDPath, vntFileName(), FO_COPY, vblnOnlyNewer, vblnNoShow)
   End If
   
End Function

Public Function CreateDir(ByVal vstrPath As String) As Boolean
   
   '//**********************************************************
   '// Check to see if the strDirectory exists and make if required
   '//**********************************************************
   
  Dim lngStart     As Long
  Dim lngPos       As Long
  Dim strDirectory As String
   
   On Error GoTo errCreation
   
   vstrPath = Trim$(vstrPath)
   '// if null string why bother....
   If vstrPath = vbNullString Then Err.Raise vbObjectError + 1
   
   lngPos = 0
   If right$(vstrPath, 1) = "\" Then vstrPath = left$(vstrPath, Len(vstrPath) - 1)
   
TryAgain:
   lngStart = lngPos + 1
   lngPos = InStr(lngStart, vstrPath, "\")
   
   If lngPos > 0 Then
      '// not at the last strDirectory in the path string...
      strDirectory = strDirectory & Mid$(vstrPath, lngStart, lngPos - lngStart) & "\"
      If InStr(1, Mid$(vstrPath, lngStart, lngPos - lngStart), ":") = 0 And LenB(Dir(strDirectory, vbDirectory)) = 0 Then
         MkDir Mid$(strDirectory, 1, Len(strDirectory) - 1)
      End If
      GoTo TryAgain
    Else
      '// the last or only strDirectory in the path string
      strDirectory = strDirectory & Mid$(vstrPath, lngStart, Len(vstrPath) - lngStart + 1)
      MkDir Mid$(strDirectory, 1, Len(strDirectory))
      strDirectory = vbNullString
   End If
   
   '// success return true
   On Error GoTo 0
   CreateDir = True
   
   Exit Function
   
   '// if it gets here, an exception was thrown
   '// propogate the error to the calling function
errCreation:
   Err.Clear
   CreateDir = False
   
End Function

Public Sub DeleteDir(ByVal vstrFileName As String, _
                     Optional ByVal vblnNoShow As Boolean = True)
   
   On Error GoTo Err_Proc
   
   vstrFileName = Trim$(vstrFileName)
   '// if null string why bother....
   If vstrFileName = vbNullString Then
      Err.Raise vbObjectError + 1
      Exit Sub
   End If
   
   If right$(vstrFileName, 1) = "\" Then vstrFileName = left$(vstrFileName, Len(vstrFileName) - 1)
   Call DeleteFile(vstrFileName, vblnNoShow)
   
Exit_Here:
   
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "DeleteDir"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub DeleteFile(ByVal vstrFileName As String, _
                      Optional ByVal vblnNoShow As Boolean = True)
   
  Dim udtOptions As SHFILEOPSTRUCT
   
   On Local Error GoTo Sub_Error
   
   vstrFileName = Trim$(vstrFileName)
   '// if null string why bother....
   If vstrFileName = vbNullString Then Err.Raise vbObjectError + 1
   
   If vstrFileName > vbNullString Then
      If right$(vstrFileName, 1) = "\" Then vstrFileName = left$(vstrFileName, Len(vstrFileName) - 1)
      With udtOptions
         .hWnd = 0
         .wFunc = FO_DELETE
         .pFrom = vstrFileName
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      End With
      If vblnNoShow Then udtOptions.fFlags = udtOptions.fFlags Or FOF_SILENT
      Call SHFileOperation(udtOptions)
      DoEvents
   End If
   
Exit_Function:
   On Local Error GoTo 0
   
   Exit Sub
   
Sub_Error:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "DeleteFile"
   Err.Clear
   Resume Exit_Function
   
End Sub

Public Sub DeleteMultiFiles(ByRef vntFileName() As String)
   
  Dim udtOptions As SHFILEOPSTRUCT
  Dim sFileNames As String
  Dim lngI       As Long
   
   On Local Error GoTo Sub_Error
   
   For lngI = LBound(vntFileName) To UBound(vntFileName)
      If vntFileName(lngI) > vbNullString Then
         sFileNames = sFileNames & vntFileName(lngI) & vbNullChar
      End If
   Next lngI
   If sFileNames > vbNullString Then
      sFileNames = sFileNames & vbNullChar
      With udtOptions
         .hWnd = 0
         .wFunc = FO_DELETE
         .pFrom = sFileNames
         .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
      End With
      Call SHFileOperation(udtOptions)
   End If
   
Exit_Sub:
   On Local Error GoTo 0
   
   Exit Sub
   
Sub_Error:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "DeleteMultiFiles"
   Err.Clear
   Resume Exit_Sub
   
End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)
   
  Dim strTemp As String
  Dim lngFN   As Long
   
   '// Purpose: Error handling - On Error
   
   '// Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If Len(vstrErrNumber) > 0 Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
      If Len(vstrErrDescription) > 0 Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If Len(vstrModuleName) > 0 Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If Len(vstrProcName) > 0 Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If
   
   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, Environ$("username"), Environ$("computername")
   Close #lngFN
   
End Sub

Public Function FileAttributesGet(ByVal vstrFileName As String) As VbFileAttribute
   
   FileAttributesGet = GetAttr(vstrFileName)
   
End Function

Public Sub FileAttributesSet(ByRef sFileName As String, Attributes As VbFileAttribute)
   
   SetAttr sFileName, Attributes
   
End Sub

Public Function FileDateTimeGet(ByVal vstrFileName As String) As Date
   
   FileDateTimeGet = FileDateTime(vstrFileName)
   
End Function

Public Sub FileDateTimeSet(ByRef vstrFileName As String, _
                           Optional ByVal vdtmDateTime As Date = 0)
   
  Dim lngHandle     As Long
  Dim udtFileTime   As FILETIME
  Dim udtLocalTime  As FILETIME
  Dim udtSystemTime As SYSTEMTIME
   
   On Error GoTo Err_Proc
   
   If vdtmDateTime = 0 Then vdtmDateTime = Now
   
   udtSystemTime.wYear = Year(vdtmDateTime)
   udtSystemTime.wMonth = Month(vdtmDateTime)
   udtSystemTime.wDay = Day(vdtmDateTime)
   udtSystemTime.wDayOfWeek = Weekday(vdtmDateTime) - 1
   udtSystemTime.wHour = Hour(vdtmDateTime)
   udtSystemTime.wMinute = Minute(vdtmDateTime)
   udtSystemTime.wSecond = Second(vdtmDateTime)
   udtSystemTime.wMilliseconds = 0
   
   '// Convert system time to local time
   SystemTimeToFileTime udtSystemTime, udtLocalTime
   '// Convert local time to GMT
   LocalFileTimeToFileTime udtLocalTime, udtFileTime
   '// Open the file to get the filehandle
   lngHandle = CreateFile(vstrFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
   '// Change date/time property of the file
   SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
   '// Close the handle
   CloseHandle lngHandle
   
Err_Proc:
   On Error GoTo 0
   
End Sub

Public Function FileLenBytesGet(ByVal vstrFileName As String) As Long
   
   FileLenBytesGet = FileLen(vstrFileName)
   
End Function

Public Function FileLenKB(ByVal sFileName As String) As Single
   
   FileLenKB = FileLen(sFileName) / 1000
   
End Function

Private Function FilterLookup(ByVal vstrFilters As String, _
                              ByVal vlngCur As Long) As String
   
  Dim lngStart As Long
  Dim lngEnd   As Long
  Dim strTemp  As String
   
   On Error GoTo Err_Proc
   
   lngStart = 1
   If vstrFilters = vbNullString Then Exit Function
   Do
      '// Cut out both parts marked by null character
      lngEnd = InStr(lngStart, vstrFilters, vbNullChar)
      If lngEnd = 0 Then Exit Function
      lngEnd = InStr(lngEnd + 1, vstrFilters, vbNullChar)
      If lngEnd Then
         strTemp = Mid$(vstrFilters, lngStart, lngEnd - lngStart)
       Else
         strTemp = Mid$(vstrFilters, lngStart)
      End If
      lngStart = lngEnd + 1
      If vlngCur = 1 Then
         FilterLookup = strTemp
         Exit Function
      End If
      vlngCur = vlngCur - 1
   Loop While vlngCur
   
Exit_Here:
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "FilterLookup"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function FolderExists(ByVal vstrFolderName As String) As Boolean
   
   On Error GoTo Err_Proc
   
   '// Does the Folder exist?
   If right$(vstrFolderName, 1) = "\" Then vstrFolderName = left$(vstrFolderName, Len(vstrFolderName) - 1)
   If LenB(Dir(vstrFolderName, vbDirectory)) > 0 Then
      FolderExists = True
    Else
      FolderExists = False
   End If
   
Exit_Here:
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "FolderExists"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function GetExtensionName(ByVal vstrFileName As String) As String
   
  Dim lngI As Long
   
   lngI = InStrRev(vstrFileName, ".")
   GetExtensionName = LCase$(Mid$(vstrFileName, lngI + 1))
   
End Function

Public Function MoveDir(ByVal vstrSPath As String, _
                        ByVal vstrDPath As String, _
                        Optional ByVal vblnNoShow As Boolean = False, _
                        Optional ByVal vblnNoConformation As Boolean = False) As Boolean
   
  Dim lngTemp As Long
   
   lngTemp = mlngNoConfirmation
   
   If vblnNoConformation Then
      mlngNoConfirmation = FOF_NOCONFIRMATION
    Else
      mlngNoConfirmation = 0
   End If
   
   If right$(vstrSPath, 1) = "\" Then vstrSPath = left$(vstrSPath, Len(vstrSPath) - 1)
   If right$(vstrDPath, 1) = "\" Then vstrDPath = left$(vstrDPath, Len(vstrDPath) - 1)
   Call CreateDir(vstrDPath)
   
   If LCase$(Me.RetOnlyPath(vstrSPath, False)) = LCase$(vstrDPath) Then
      MoveDir = False
    Else
      MoveDir = CopyMoveSingle(vstrSPath, vstrDPath, FO_MOVE, False, vblnNoShow)
   End If
   
   mlngNoConfirmation = lngTemp
   
End Function

Public Function MoveFile(ByVal vstrSPathFile As String, _
                         ByVal vstrDPathFile As String, _
                         Optional ByVal vblnOnlyNewer As Boolean = False, _
                         Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
   If vblnOnlyNewer And Not NotTheSame(vstrSPathFile, vstrDPathFile) Then Exit Function
   
   If LCase$(vstrSPathFile) = LCase$(vstrDPathFile) Then
      MoveFile = False
    Else
      MoveFile = CopyMoveSingle(vstrSPathFile, vstrDPathFile, FO_MOVE, vblnOnlyNewer, vblnNoShow)
   End If
   
End Function

Public Function MoveMultiFiles(ByVal vstrSPath As String, _
                               ByVal vstrDPath As String, _
                               ByRef vntFileName() As String, _
                               Optional ByVal vblnOnlyNewer As Boolean = False, _
                               Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
   If LCase$(vstrSPath) = LCase$(vstrDPath) Then
      MoveMultiFiles = False
    Else
      MoveMultiFiles = CopyMoveMulti(vstrSPath, vstrDPath, vntFileName(), FO_MOVE, vblnOnlyNewer, vblnNoShow)
   End If
   
End Function

Public Property Get NoConfirmation() As Boolean
   
   NoConfirmation = CBool(mlngNoConfirmation)
   
End Property

Public Property Let NoConfirmation(ByVal vNewValue As Boolean)
   
   If vNewValue Then
      mlngNoConfirmation = FOF_NOCONFIRMATION
    Else
      mlngNoConfirmation = 0
   End If
   
End Property

Private Function NotTheSame(ByVal vstrSPathFileName As String, _
                            ByVal vstrDPathFileName As String) As Boolean
   
   On Error GoTo Err_Proc
   
   '    If LenB(Dir$(vstrSPathFileName)) = 0 Then
   '        NotTheSame = False
   '    Else
   '        If LenB(Dir$(vstrDPathFileName)) = 0 Then
   '            NotTheSame = True
   '        Else
   NotTheSame = CBool(FileDateTime(vstrSPathFileName) > FileDateTime(vstrDPathFileName))
   '        End If
   '    End If
   
   Exit Function
   
Err_Proc:
   NotTheSame = True
   
End Function

Public Function RenameFile(ByVal vstrSPathFile As String, _
                           ByVal vstrDPathFile As String, _
                           Optional ByVal vblnNoShow As Boolean = False) As Boolean
   
  Dim udtOptions As SHFILEOPSTRUCT
   
   On Error GoTo Function_Error
   
   With udtOptions
      .hWnd = 0
      .wFunc = FO_RENAME
      .pFrom = vstrSPathFile & vbNullChar
      .pTo = vstrDPathFile & vbNullChar
      .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
   End With
   If vblnNoShow Then udtOptions.fFlags = udtOptions.fFlags Or FOF_SILENT
   RenameFile = Not CBool(SHFileOperation(udtOptions))
   
Exit_Function:
   On Error GoTo 0
   
   Exit Function
   
Function_Error:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "RenameFile"
   Err.Clear
   Resume Exit_Function
   
End Function

Public Function RetFileList(ByVal vstrPathFilter As String, _
                            ByRef rstrFileList() As String, _
                            Optional vblnIncludePath As Boolean = False) As Boolean
   
  Dim lngI             As Long
  Dim strNewFileName   As String
  Dim strRootPath      As String
  Dim blnFileFound     As Boolean
   
   On Error GoTo Err_Proc
   
   If right$(vstrPathFilter, 1) = "\" Then vstrPathFilter = vstrPathFilter & "*.*"
   
   blnFileFound = False
   lngI = 0
   ReDim rstrFileList(lngI) As String
   strRootPath = Me.RetOnlyPath(vstrPathFilter)
   
   strNewFileName = Dir(vstrPathFilter)
   If LenB(strNewFileName) Then
      Do
         ReDim Preserve rstrFileList(lngI) As String
         If vblnIncludePath Then
            rstrFileList(lngI) = LCase$(strRootPath & strNewFileName)
          Else
            rstrFileList(lngI) = LCase$(strNewFileName)
         End If
         lngI = lngI + 1
         blnFileFound = True
         strNewFileName = Dir
      Loop Until LenB(strNewFileName) = 0
   End If
   
Exit_Here:
   RetFileList = blnFileFound
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "RetFileList"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function RetOnlyFilename(ByVal vstrPathFileName As String) As String
   
   RetOnlyFilename = right$(vstrPathFileName, (Len(vstrPathFileName) - InStrRev(vstrPathFileName, "\", , vbTextCompare)))
   
End Function

Public Function RetOnlyPath(ByVal vstrPathFileName As String, _
                            Optional ByVal vblnReturnSlash As Boolean = True) As String
   
   If right$(vstrPathFileName, 1) = "\" Then
      vstrPathFileName = left$(vstrPathFileName, Len(vstrPathFileName) - 1)
   End If
   
   If vblnReturnSlash Then
      RetOnlyPath = left$(vstrPathFileName, InStrRev(vstrPathFileName, "\", , vbTextCompare))
    Else
      RetOnlyPath = left$(vstrPathFileName, InStrRev(vstrPathFileName, "\", , vbTextCompare) - 1)
   End If
   
End Function

Public Function RetShortPathName(ByVal vstrLongPathName As String) As String
   
  Dim strShortPath  As String
  Dim lngPathLength As Long
   
   '// Returns short (DOS) pathname of the passed long pathname
   strShortPath = Space$(250)
   lngPathLength = GetShortPathName(vstrLongPathName, strShortPath, Len(strShortPath))
   If lngPathLength Then
      RetShortPathName = LCase$(left$(strShortPath, lngPathLength))
   End If
   
End Function

Public Function RetSubDirList(ByVal vstrRootPath As String, _
                              ByRef rstrSubDirList() As String) As Boolean
   
  Dim lngI                          As Long
  Dim strNewFileName                As String
  Dim blnDirFound                   As Boolean
  Const FILE_ATTRIBUTE_DIRECTORY    As Long = &H10
   
   On Error GoTo Err_Proc
   
   lngI = 0
   ReDim rstrSubDirList(lngI) As String
   If right$(vstrRootPath, 1) <> "\" Then vstrRootPath = vstrRootPath & "\"
   
   strNewFileName = Dir(vstrRootPath & "*.*", vbDirectory)
   If strNewFileName > vbNullString Then
      Do
         If left$(strNewFileName, 1) <> "." Then
            If GetFileAttributes(vstrRootPath & strNewFileName) = FILE_ATTRIBUTE_DIRECTORY Then
               ReDim Preserve rstrSubDirList(lngI) As String
               rstrSubDirList(lngI) = strNewFileName
               lngI = lngI + 1
               blnDirFound = True
            End If
         End If
         strNewFileName = Dir
         If strNewFileName = vbNullString Then Exit Do
      Loop
   End If
   RetSubDirList = blnDirFound
   
Exit_Here:
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "RetSubDirList"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function RetTempPath(Optional ByVal vblnReturnSlash As Boolean = True) As String
   
   If vblnReturnSlash Then
      RetTempPath = Environ$("TEMP") & "\"
    Else
      RetTempPath = Environ$("TEMP")
   End If
   
End Function

Public Function StripQuotes(ByVal vstrTextString As String) As String
   
   StripQuotes = Replace(vstrTextString, Chr$(34), vbNullString)
   
End Function

Private Function StrZToStr(ByVal vstrTextString As String) As String
   
   StrZToStr = left$(vstrTextString, lstrlen(vstrTextString))
   
End Function

Public Function VBGetDirectory(Optional ByRef rOwnerForm As Form, _
                               Optional ByVal vstrTitle As String = "Select" & " a Directory", _
                               Optional vblnShowCreateNewFolder As Boolean = True) As String
   
   '// Opens a Browse Folders Dialog Box that displays the
   '// directories in your computer
   
  Dim lngDList      As Long '// Declare Varibles
  Dim strBuffer     As String
  Dim lng_hWnd      As Long
  Dim udtBrowseInfo As BrowseInfo
   
   On Error Resume Next
   
   lng_hWnd = rOwnerForm.lng_hWnd
   
   With udtBrowseInfo
      .hWndOwner = lng_hWnd
      .lpszTitle = lstrcat(vstrTitle, vbNullChar)
      If vblnShowCreateNewFolder Then
         .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_NEWDIALOGSTYLE
       Else
         .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT
      End If
   End With
   
   lngDList = SHBrowseForFolder(udtBrowseInfo)
   If (lngDList) Then
      strBuffer = Space$(C_lngMAX_PATH)
      SHGetPathFromIDList lngDList, strBuffer
      strBuffer = left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
      VBGetDirectory = strBuffer
   End If
   
End Function

Public Function VBGetOpenFileName(ByRef rstrPathFileName As String, _
                                  Optional ByRef rstrFileName As String, _
                                  Optional ByVal vblnFileMustExist As Boolean = True, _
                                  Optional ByVal vblnMultiSelect As Boolean = False, _
                                  Optional ByVal vblnReadOnly As Boolean = False, _
                                  Optional ByVal vblnHideReadOnly As Boolean = False, _
                                  Optional ByVal rstrFilter As String = "All (*.*)| *.*", _
                                  Optional ByVal rlngFilterIndex As Long = 1, _
                                  Optional ByVal vstrInitDir As String, _
                                  Optional ByVal vstrDlgTitle As String, _
                                  Optional ByVal vstrDefaultExt As String, _
                                  Optional ByVal vlngOwner_hWnd As Long = -1, _
                                  Optional ByVal venuFlags As enuOpenFile = 0) As Boolean
   
  Dim lngCH      As String
  Dim lngI       As Long
  Dim udtOpFile  As OPENFILENAME
  Dim strTemp    As String
  Dim lngMax     As Long
   
   On Error GoTo Err_Proc
   
   mlngApiReturn = 0
   mlngExtendedError = 0
   
   With udtOpFile
      .lStructSize = Len(udtOpFile)
      
      '// Add in specific Flags and strip out non-VB Flags
      
      .Flags = (-vblnFileMustExist * OFN_FileMustExist) Or (-vblnMultiSelect * OFN_ALLOWMULTISELECT) Or (-vblnReadOnly * OFN_ReadOnly) Or (-vblnHideReadOnly * OFN_HideReadOnly) Or (venuFlags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
      
      '// Owner can take handle of owning window
      If vlngOwner_hWnd <> -1 Then .hWndOwner = vlngOwner_hWnd
      '// vstrInitDir can take initial directory string
      .lpstrInitialDir = vstrInitDir
      '// vstrDefaultExt can take default extension
      .lpstrDefExt = vstrDefaultExt
      '// vstrDlgTitle can take dialog box title
      .lpstrTitle = vstrDlgTitle
      
      '// To make Windows-style filter, replace | and : with nulls
      For lngI = 1 To Len(rstrFilter)
         lngCH = Mid$(rstrFilter, lngI, 1)
         If lngCH = "|" Or lngCH = ":" Then
            strTemp = strTemp & vbNullChar
          Else
            strTemp = strTemp & lngCH
         End If
      Next lngI
      '// Put double null at end
      strTemp = strTemp & vbNullChar & vbNullChar
      .lpstrFilter = strTemp
      .nFilterIndex = rlngFilterIndex
      
      '// Pad file and file title buffers to maximum path
      lngMax = C_lngMAX_PATH
      If (.Flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
         lngMax = 8192
      End If
      strTemp = rstrPathFileName & String$(lngMax - Len(rstrPathFileName), 0)
      .lpstrFile = strTemp
      .nMaxFile = lngMax
      strTemp = rstrFileName & String$(lngMax - Len(rstrFileName), 0)
      .lpstrFileTitle = strTemp
      .nMaxFileTitle = lngMax
      '// All other fields set to zero
      
      mlngApiReturn = GetOpenFilename(udtOpFile)
      Select Case mlngApiReturn
       Case 1
         '// Success
         VBGetOpenFileName = True
         If (.Flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
            rstrFileName = vbNullString
            lngMax = InStr(.lpstrFile, Chr$(0) & Chr$(0))
            If (lngMax = 0) Then
               rstrPathFileName = StrZToStr(.lpstrFile)
             Else
               rstrPathFileName = left$(.lpstrFile, lngMax - 1)
            End If
          Else
            rstrPathFileName = StrZToStr(.lpstrFile)
            rstrFileName = StrZToStr(.lpstrFileTitle)
         End If
         'venuFlags = .Flags
         '// Return the filter index
         rlngFilterIndex = .nFilterIndex
         '// Look up the filter the user selected and return that
         rstrFilter = FilterLookup(.lpstrFilter, rlngFilterIndex)
         If (.Flags And OFN_ReadOnly) Then vblnReadOnly = True
       Case 0
         '// Cancelled
         VBGetOpenFileName = False
         rstrPathFileName = vbNullString
         rstrFileName = vbNullString
         'venuFlags = 0
         rlngFilterIndex = -1
         rstrFilter = vbNullString
       Case Else
         '// Extended error
         mlngExtendedError = CommDlgExtendedError()
         VBGetOpenFileName = False
         rstrPathFileName = vbNullString
         rstrFileName = vbNullString
         'venuFlags = 0
         rlngFilterIndex = -1
         rstrFilter = vbNullString
      End Select
   End With
   
Exit_Here:
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "VBGetOpenFileName"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function VBGetSaveFileName(ByRef rstrPathFileName As String, _
                                  Optional ByRef rstrFileName As String, _
                                  Optional ByVal vblnOverWritePrompt As Boolean = True, _
                                  Optional ByVal rstrFilter As String = "All" & " (*.*)| *.*", _
                                  Optional ByVal rlngFilterIndex As Long = 1, _
                                  Optional ByVal vstrInitDir As String, _
                                  Optional ByVal vstrDlgTitle As String, _
                                  Optional ByVal vstrDefaultExt As String, _
                                  Optional ByVal vlngOwner_hWnd As Long = -1, _
                                  Optional ByVal venuFlags As enuOpenFile) As Boolean
   
  Dim udtOpFile As OPENFILENAME
  Dim strTemp   As String
  Dim strCH     As String
  Dim lngI      As Long
   
   On Error GoTo Err_Proc
   
   mlngApiReturn = 0
   mlngExtendedError = 0
   
   With udtOpFile
      .lStructSize = Len(udtOpFile)
      
      '// Add in specific Flags and strip out non-VB Flags
      .Flags = (-vblnOverWritePrompt * OFN_OverWritePrompt) Or OFN_HideReadOnly Or (venuFlags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
      
      '// Owner can take handle of owning window
      If vlngOwner_hWnd <> -1 Then .hWndOwner = vlngOwner_hWnd
      '// vstrInitDir can take initial directory string
      .lpstrInitialDir = vstrInitDir
      '// vstrDefaultExt can take default extension
      .lpstrDefExt = vstrDefaultExt
      '// vstrDlgTitle can take dialog box title
      .lpstrTitle = vstrDlgTitle
      
      '// Make new filter with bars (|) replacing nulls and double null at end
      For lngI = 1 To Len(rstrFilter)
         strCH = Mid$(rstrFilter, lngI, 1)
         If strCH = "|" Or strCH = ":" Then
            strTemp = strTemp & vbNullChar
          Else
            strTemp = strTemp & strCH
         End If
      Next lngI
      
      '// Put double null at end
      strTemp = strTemp & vbNullChar & vbNullChar
      .lpstrFilter = strTemp
      .nFilterIndex = rlngFilterIndex
      
      '// Pad file and file title buffers to maximum path
      strTemp = rstrPathFileName & String$(C_lngMAX_PATH - Len(rstrPathFileName), 0)
      .lpstrFile = strTemp
      .nMaxFile = C_lngMAX_PATH
      strTemp = rstrFileName & String$(C_lngMAX_FILE - Len(rstrFileName), 0)
      .lpstrFileTitle = strTemp
      .nMaxFileTitle = C_lngMAX_FILE
      
      '// All other fields zero
      mlngApiReturn = GetSaveFilename(udtOpFile)
      
      Select Case mlngApiReturn
       Case 1
         VBGetSaveFileName = True
         rstrPathFileName = StrZToStr(.lpstrFile)
         rstrFileName = StrZToStr(.lpstrFileTitle)
         'venuFlags = .Flags
         '// Return the filter index
         rlngFilterIndex = .nFilterIndex
         '// Look up the filter the user selected and return that
         rstrFilter = FilterLookup(.lpstrFilter, rlngFilterIndex)
       Case 0
         '// Cancelled:
         VBGetSaveFileName = False
         rstrPathFileName = vbNullString
         rstrFileName = vbNullString
         'venuFlags = 0
         rlngFilterIndex = 0
         rstrFilter = vbNullString
       Case Else
         '// Extended error:
         VBGetSaveFileName = False
         mlngExtendedError = CommDlgExtendedError()
         rstrPathFileName = vbNullString
         rstrFileName = vbNullString
         'venuFlags = 0
         rlngFilterIndex = 0
         rstrFilter = vbNullString
      End Select
   End With
   
Exit_Here:
   
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsFileUtilities", "VBGetSaveFileName"
   Err.Clear
   Resume Exit_Here
   
End Function

