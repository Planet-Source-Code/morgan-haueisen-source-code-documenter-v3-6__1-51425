VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   BorderStyle     =   0  'None
   Caption         =   "View File"
   ClientHeight    =   7920
   ClientLeft      =   1380
   ClientTop       =   1785
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "frmView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   45
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1410
      TabIndex        =   2
      Top             =   45
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2805
      TabIndex        =   1
      Top             =   45
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   7140
      Left            =   0
      TabIndex        =   0
      Top             =   585
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   12594
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuSave 
      Caption         =   "mnuSave"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveRTF 
         Caption         =   "Save As RTF"
      End
      Begin VB.Menu mnuSaveHTML 
         Caption         =   "Save as HTML"
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//*************************************/
'//     Author: Morgan Haueisen        */
'//             morganh@hartcom.net    */
'//     Copyright (c) 1996-2004        */
'//*************************************/
Option Explicit

Public mblnSaveRTF       As Boolean
Public mblnSaveHTML      As Boolean

Private mstrFileSaveName As String
Private mcObjEditor      As clsEditor
Private mstrFileHeader   As String

Private mcFile           As clsFileUtilities

'// Used for Manifest files (Win XP)
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub cmdClose_Click()
   
   Unload Me
   
End Sub

Private Sub cmdPrint_Click()
   
   On Error GoTo Err_Proc
   
   If frmMsgBox.SMessageModal("Ok to print the file?", vbQuestion + vbYesNo) = vbYes Then
      Printer.ColorMode = vbPRCMColor
      Printer.Print vbNullString;
      Printer.Print Me.txtSource.Text
      Printer.EndDoc
   End If
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "cmdPrint_Click"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub cmdSave_Click()
   
   PopupMenu mnuSave, , cmdSave.Left, cmdSave.Top + cmdSave.Height
   
End Sub

Private Sub Form_Activate()
   
   Screen.MousePointer = vbDefault
   Me.WindowState = 2
   
End Sub

Private Sub Form_Initialize()
   
   '// Used for Manifest files (Win XP)
   Call InitCommonControls
   
   Set mcFile = New clsFileUtilities
   
End Sub

Private Sub Form_Load()
   
   On Error GoTo Err_Proc
   
   Screen.MousePointer = vbHourglass
   DoEvents
   Me.Move 0, 0, Screen.Width, Screen.Height
   
   Set mcObjEditor = New clsEditor
   
   '// set editor objects
   mcObjEditor.SetEditorObjects txtSource
   InitWords mcObjEditor
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "Form_Load"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub Form_Resize()
   
   txtSource.Width = Me.ScaleWidth
   txtSource.Height = Me.ScaleHeight - txtSource.Top
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Screen.MousePointer = vbDefault
   Set mcObjEditor = Nothing
   Set mcFile = Nothing
   Set frmView = Nothing
   
End Sub

Private Function GetColor(ByVal vstrColorCode As String) As String
   
   GetColor = "<FONT COLOR=" & Chr$(34) & "#" & CStr(vstrColorCode) & Chr$(34) & ">"
   'getColor = "<FONT COLOR=" & CStr(vstrColorCode) & ">"
   
End Function

Private Sub InitWords(ByRef rcObjEditor As clsEditor)
   
   On Error GoTo Err_Proc
   
   '// hard code init the basic vb script words -
   '// you can init any words you want with any colors you like
   With rcObjEditor
      .AddEditorWord "On Local Error", vbMagenta
      .AddEditorWord "On Error", vbMagenta
      
      .AddEditorWord "Dim", vbRed
      .AddEditorWord "Redim", vbRed
      .AddEditorWord "Type", vbBlue
      
      .AddEditorWord "Public", vbBlue
      .AddEditorWord "Private", vbBlue
      .AddEditorWord "Function", vbBlue
      .AddEditorWord "Sub", vbBlue
      .AddEditorWord "End", vbBlue
      
      .AddEditorWord "Option", vbBlue
      .AddEditorWord "Explicit", vbBlue
      
      .AddEditorWord "Select ", vbBlue
      .AddEditorWord "Until", vbBlue
      .AddEditorWord "Set", vbBlue
      .AddEditorWord "For", vbBlue
      .AddEditorWord "Next", vbBlue
      .AddEditorWord "Do", vbBlue
      .AddEditorWord "Loop", vbBlue
      .AddEditorWord "If", vbBlue
      .AddEditorWord "Select", vbBlue
      .AddEditorWord "Case", vbBlue
      .AddEditorWord "Then", vbBlue
      .AddEditorWord "Else", vbBlue
      .AddEditorWord "ElseIf", vbBlue
      .AddEditorWord "Open", vbBlue
      .AddEditorWord "Exit", vbBlue
      .AddEditorWord "On", vbBlue
      .AddEditorWord "Resume", vbBlue
      .AddEditorWord "New", vbBlue
      .AddEditorWord "Close", vbBlue
      .AddEditorWord "Print", vbBlue
      .AddEditorWord "Preserve", vbBlue
      .AddEditorWord "Error", vbBlue
      .AddEditorWord "Err.", vbBlue
      .AddEditorWord "False", vbBlue
      .AddEditorWord "True", vbBlue
      
      .AddEditorWord "MsgBox", vbBlue
      .AddEditorWord "Instr", vbBlue
      .AddEditorWord "InstrRev", vbBlue
      .AddEditorWord "GoTo", vbBlue
      .AddEditorWord "GoSub", vbBlue
      .AddEditorWord "Return", vbBlue
      .AddEditorWord "With", vbBlue
      .AddEditorWord "Optional", vbBlue
      .AddEditorWord "String", vbBlue
      .AddEditorWord "Integer", vbBlue
      .AddEditorWord "Boolean", vbBlue
      .AddEditorWord "Long", vbBlue
      .AddEditorWord "Double", vbBlue
      .AddEditorWord "Single", vbBlue
      .AddEditorWord "Byte", vbBlue
      .AddEditorWord "Variant", vbBlue
      .AddEditorWord "String,", vbBlue
      .AddEditorWord "Integer,", vbBlue
      .AddEditorWord "Boolean,", vbBlue
      .AddEditorWord "Long,", vbBlue
      .AddEditorWord "Double,", vbBlue
      .AddEditorWord "Single,", vbBlue
      .AddEditorWord "Byte,", vbBlue
      .AddEditorWord "Variant,", vbBlue
      
      .AddEditorWord "As", vbBlue
      .AddEditorWord "Const", vbBlue
      .AddEditorWord "Static", vbBlue
      .AddEditorWord "Base", vbBlue
      .AddEditorWord "Input", vbBlue
      .AddEditorWord "Output", vbBlue
      .AddEditorWord "Line", vbBlue
   End With
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "InitWords"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub mnuSaveHtml_Click()
   
   Call WriteHTML
   
End Sub

Private Sub mnuSaveRTF_Click()
   
   txtSource.SaveFile mstrFileSaveName & ".rtf"
   
End Sub

Public Sub ShowExport(ByVal strView As String)
   
   On Error GoTo Exit_Proc
   
   mstrFileSaveName = mcFile.RetOnlyPath(gstrProjectName) & "Documentation\"
   Call mcFile.CreateDir(mstrFileSaveName)
   
   mstrFileSaveName = mstrFileSaveName & mcFile.RetOnlyFilename(gstrProjectName) & "_Summary"
   
   txtSource.Text = strView
   
   If Not mblnSaveRTF And Not mblnSaveHTML Then
      Me.Show , frmMain
    Else
      If mblnSaveRTF Then Call mnuSaveRTF_Click
      If mblnSaveHTML Then Call mnuSaveHtml_Click
      Unload Me
   End If
   
Exit_Proc:
   
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "ShowFile"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Public Sub ShowFile(ByVal vstrFilePath As String, _
                    Optional ByVal vblnShowInterface As Boolean = False, _
                    Optional ByVal vstrFileName As String = vbNullString)
   
   On Error GoTo Err_Proc
   
   txtSource.Text = vbNullString
   
   If LenB(vstrFileName) Then
      mstrFileHeader = "'/" & String$(75, "*") & vbCrLf & _
            "'//  File Name: " & vstrFileName & vbCrLf & _
            "'//  File Size: " & Format$(FileLen(vstrFilePath) / 1000, "#,0.0") & " KB" & vbCrLf & _
            "'//  File Date: " & Format$(FileDateTime(vstrFilePath), "m/d/yy  h:mm:ss ampm") & vbCrLf & _
            "'// Printed On: " & Format$(Now, "ddd. mmmm d, yyyy  h:mm:ss ampm") & vbCrLf & _
            "'/" & String$(75, "*") & vbCrLf & vbCrLf
    Else
      mstrFileHeader = "'/" & String$(75, "*") & vbCrLf & _
            "'//  File Name: " & mcFile.RetOnlyFilename(vstrFilePath) & vbCrLf & _
            "'//  File Size: " & Format$(FileLen(vstrFilePath) / 1000, "#,0.0") & " KB" & vbCrLf & _
            "'//  File Date: " & Format$(FileDateTime(vstrFilePath), "m/d/yy  h:mm:ss ampm") & vbCrLf & _
            "'// Printed On: " & Format$(Now, "ddd. mmmm d, yyyy  h:mm:ss ampm") & vbCrLf & _
            "'/" & String$(75, "*") & vbCrLf & vbCrLf
   End If
   
   mstrFileSaveName = mcFile.RetOnlyPath(mcFile.RetOnlyPath(gstrProjectName) & "Documentation\" & vstrFileName)
   Call mcFile.CreateDir(mstrFileSaveName)
   
   mstrFileSaveName = mstrFileSaveName & mcFile.RetOnlyFilename(vstrFilePath)
   
   Call UpdateTextView(vstrFilePath, vblnShowInterface)
   
   mcObjEditor.PaintText True, True
   
   If Not mblnSaveRTF And Not mblnSaveHTML Then
      Me.Show , frmMain
    Else
      If mblnSaveRTF Then Call mnuSaveRTF_Click
      If mblnSaveHTML Then Call mnuSaveHtml_Click
      Unload Me
   End If
   
Exit_Proc:
   
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "ShowFile"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub UpdateTextView(ByVal vstrFilePath As String, _
                           Optional ByVal vblnShowInterface As Boolean = False)
   
  Dim lngFN         As Long
  Dim strLine       As String
  Dim strView       As String
  Dim blnFlag       As Boolean
  Dim blnFlag2      As Boolean
   
   On Error GoTo Err_Proc
   
   '//  Open Source file:
   lngFN = FreeFile
   Open vstrFilePath For Input As #lngFN
   strView = vbNullString
   Do Until EOF(lngFN)
      Line Input #lngFN, strLine
      blnFlag2 = InStr(Trim$(strLine), "Attribute VB_") = 1
      If Not blnFlag Or blnFlag2 Then
         strLine = vbNullString
         blnFlag = blnFlag2
       Else
         strView = strView & strLine & vbNewLine
      End If
   Loop
   
Exit_Proc:
   On Error Resume Next
   Close #lngFN
   txtSource.Text = mstrFileHeader & strView
   
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "UpdateTextView"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub WriteHTML()
   
  Dim lngFN     As Long
  Dim strRTF    As String
  Dim strHTML   As String
  Dim strCT()   As String
  Dim strTemp   As String
  Dim lngI      As Long
  Dim lngN      As Long
  Dim lngX      As Long
   
   On Error GoTo Err_Proc
   
   Screen.MousePointer = vbHourglass
   DoEvents
   
   strRTF = txtSource.TextRTF
   lngI = InStrRev(strRTF, "}")
   If lngI Then strRTF = Mid$(strRTF, 1, lngI - 1)
   
   '// Get color table
   ReDim strCT(0)
   strCT(0) = 0
   lngI = InStr(strRTF, "{\colortbl ;")
   If lngI Then
      lngN = InStr(lngI, strRTF, "}")
      Do
         ReDim Preserve strCT(UBound(strCT) + 1)
         lngI = InStr(lngI, strRTF, "\red") + 4
         strTemp = Mid$(strRTF, lngI, 1)
         lngI = lngI + 1
         If Mid$(strRTF, lngI, 1) <> "\" Then
            strTemp = strTemp & Mid$(strRTF, lngI, 1)
            lngI = lngI + 1
            If Mid$(strRTF, lngI, 1) <> "\" Then strTemp = strTemp & Mid$(strRTF, lngI, 1)
         End If
         strCT(UBound(strCT)) = Right$("00" & Hex(strTemp), 2)
         
         lngI = InStr(lngI, strRTF, "\green") + 6
         strTemp = Mid$(strRTF, lngI, 1)
         lngI = lngI + 1
         If Mid$(strRTF, lngI, 1) <> "\" Then
            strTemp = strTemp & Mid$(strRTF, lngI, 1)
            lngI = lngI + 1
            If Mid$(strRTF, lngI, 1) <> "\" Then strTemp = strTemp & Mid$(strRTF, lngI, 1)
         End If
         strCT(UBound(strCT)) = strCT(UBound(strCT)) & Right$("00" & Hex(strTemp), 2)
         
         lngI = InStr(lngI, strRTF, "\blue") + 5
         strTemp = Mid$(strRTF, lngI, 1)
         lngI = lngI + 1
         If Mid$(strRTF, lngI, 1) <> ";" Then
            strTemp = strTemp & Mid$(strRTF, lngI, 1)
            lngI = lngI + 1
            If Mid$(strRTF, lngI, 1) <> ";" Then strTemp = strTemp & Mid$(strRTF, lngI, 1)
         End If
         strCT(UBound(strCT)) = strCT(UBound(strCT)) & Right$("00" & Hex(strTemp), 2)
         
      Loop Until lngI + 2 >= lngN
   End If
   
   '// get first line with text on it
   lngI = InStr(strRTF, "\viewkind4")
   'If lngI = 0 Then GoTo Exit_Proc
   lngN = InStr(lngI, strRTF, vbNewLine)
   strTemp = Mid$(strRTF, lngI, lngN - lngI)
   lngN = InStr(strTemp, "\cf")
   If lngN Then
      For lngX = 0 To UBound(strCT)
         If InStr(strTemp, "\cf" & CStr(lngX)) < lngN + 5 Then
            strHTML = strHTML & GetColor(strCT(lngX))
            Exit For
         End If
      Next lngX
   End If
   lngN = InStrRev(strTemp, "\fs") + lngI
   
   '// Remove RTF header
   strHTML = strHTML & Mid$(strRTF, lngN + 5)
   
   '// Replace RTF color codes with HTML
   For lngI = 0 To UBound(strCT)
      strHTML = Replace$(strHTML, "\cf" & CStr(lngI), GetColor(strCT(lngI)))
   Next
   
   '// Replace RTF start of line
   strHTML = Replace$(strHTML, "\par ", vbNullString)
   '// Replace RTF Line Feed
   strHTML = Replace$(strHTML, vbCrLf, "<br>" & vbCrLf)
   strHTML = Replace$(strHTML, "\\", "\")
   
   '// Replace tabbed index
   lngI = 1
   lngN = 0
   strTemp = vbNullString
   Do
      lngI = InStr(lngI, strHTML, "<br>" & vbCrLf)
      If lngI Then
         lngI = lngI + 6
         Do
            If Mid$(strHTML, lngI + lngN, 1) = " " Then
               lngN = lngN + 1
               strTemp = strTemp & "&nbsp;"
             Else
               Exit Do
            End If
         Loop
         
         If lngN Then
            strHTML = Mid$(strHTML, 1, lngI - 1) & strTemp & Mid$(strHTML, lngI + lngN)
            lngN = 0
            strTemp = vbNullString
         End If
         
      End If
   Loop Until lngI = 0
   
   strHTML = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCrLf _
         & "<HTML>" & vbCrLf _
         & "<HEAD>" & vbCrLf _
         & "<TITLE>" & mcFile.RetOnlyFilename(mstrFileSaveName) & "</TITLE>" & vbCrLf _
         & "</HEAD>" & vbCrLf _
         & "<BODY>" & vbCrLf _
         & "<span style='font-size:9.0pt;font-family:" & Chr$(34) & "Courier New" & Chr$(34) & "'>" _
         & strHTML _
         & "</BODY>" & vbCrLf _
         & "</HTML>"
   
   lngFN = FreeFile
   Open mstrFileSaveName & ".htm" For Output As #lngFN
   Print #lngFN, strHTML
   
Exit_Proc:
   On Error Resume Next
   Close #lngFN
   Screen.MousePointer = vbDefault
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "WriteHTML"
   Err.Clear
   Resume Exit_Proc
   
End Sub

