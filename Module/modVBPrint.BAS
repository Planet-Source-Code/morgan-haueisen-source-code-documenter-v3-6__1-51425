Attribute VB_Name = "modVBPrint"
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
Option Compare Text

Public Const C_UNKOWN As String = "Unknown/Unregistered"

Public gcPrint              As clsMultiPgPreview

Public gastrFileNameList()  As String
Public gastrFilePathList()  As String
Public galngFileSizeList()  As Long
Public gastrProjectGroups() As String

Public gblnQuitCommand      As Boolean
Public gstrRevisionDate     As String
Public glngLinesOfCode      As Long
Public glngLinesOfCom       As Long
Public glngLinesOfBlank     As Long
Public gstrProjectName      As String
Public gbln_IsVBG           As Boolean

Private Type ProjectInfoType
   IconForm                As String
   Startup                 As String
   HelpFile                As String
   Title                   As String
   ExeName32               As String
   Command32               As String
   PName                   As String
   HelpContextID           As String
   CompatibleMode          As String
   MajorVer                As String
   MinorVer                As String
   RevisionVer             As String
   AutoIncrementVer        As String
   ServerSupportFiles      As String
   CompilationType         As String
   OptimizationType        As String
   FavorPentiumPro         As String
   CodeViewDebugInfo       As String
   NoAliasing              As String
   BoundsCheck             As String
   OverflowCheck           As String
   FlPointCheck            As String
   FDIVCheck               As String
   UnroundedFP             As String
   StartMode               As String
   Unattended              As String
   Retained                As String
   ThreadPerObject         As String
   MaxNumberOfThreads      As String
   VersionComments         As String
   VersionCompanyName      As String
   VersionFileDescription  As String
   VersionLegalCopyright   As String
   VersionLegalTrademarks  As String
   VersionProductName      As String
End Type
Public gaudtProjectInfo() As ProjectInfoType

Public Function FixedPath(ByVal vPath As String) As String
  
  Dim strTemp As String
  Dim lngC    As Long
  Dim lngI    As Long
   
   '// Changes the path name from "..\..\MyPath" to a literal path string
   Do
      lngC = lngC + 1
      vPath = Mid$(vPath, 4)
   Loop Until left$(vPath, 1) <> "."
   
   strTemp = left$(gstrProjectName, InStrRev(gstrProjectName, "\", , vbTextCompare) - 1)
   For lngI = 1 To lngC
      strTemp = left$(strTemp, InStrRev(strTemp, "\", , vbTextCompare) - 1)
   Next lngI
   
   FixedPath = strTemp & "\" & vPath
      
End Function

Public Function GetTypeLibInfo(ByVal vClassID As String, _
                                 Optional ByVal vGetName As Boolean = False) As String
  
  Dim strTemp  As String
  Dim lngI     As Long
  Dim cReg     As clsRegistry

   Set cReg = New clsRegistry
   
   '// TypeLib's path = "HKEY_CLASSES_ROOT\TypeLib\{ClassID}\{version}\0\Win32"
   '// TypeLib's name = "HKEY_CLASSES_ROOT\TypeLib\{ClassID}\{version}"
   
   If vGetName Then
      lngI = InStrRev(vClassID, "\")
      vClassID = left$(vClassID, lngI - 1)
      strTemp = "TypeLib\" & vClassID
   Else
      strTemp = "TypeLib\" & vClassID & "\Win32"
   End If

   
   strTemp = cReg.QueryValue(HKEY_CLASSES_ROOT, strTemp, "")

   If Len(strTemp) Then
      GetTypeLibInfo = left$(strTemp, Len(strTemp) - 1)
   Else
      GetTypeLibInfo = C_UNKOWN
   End If
   
   Set cReg = Nothing
   
End Function

Public Sub Main()
   
   If App.Comments = "exe" Then
      Call ManifestWrite
   End If
   
   frmMain.Show
   
End Sub

