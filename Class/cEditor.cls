VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const C_WM_USER                  As Long = &H400
Private Const C_CFM_BACKCOLOR            As Long = &H4000000
Private Const C_EM_GETCHARFORMAT         As Long = (C_WM_USER + 58)
Private Const C_EM_SETCHARFORMAT         As Long = (C_WM_USER + 68)
Private Const C_SCF_SELECTION            As Long = &H1&
Private Const C_LF_FACESIZE              As Integer = 32

Private Type CHARFORMAT2
   cbSize                               As Integer '// 2
   wPad1                                As Integer '// 4
   dwMask                               As Long    '// 8
   dwEffects                            As Long    '// 12
   yHeight                              As Long    '// 16
   yOffset                              As Long    '// 20
   crTextColor                          As Long    '// 24
   bCharSet                             As Byte    '// 25
   bPitchAndFamily                      As Byte    '// 26
   szFaceName(0 To C_LF_FACESIZE - 1)   As Byte    '// 58
   wPad2                                As Integer '// 60
   '// Additional stuff supported by RICHEDIT20
   wWeight                              As Integer '// Font weight (LOGFONT value)
   sSpacing                             As Integer '// Amount to space between letters
   crBackColor                          As Long    '// Background color
   lLCID                                As Long    '// Locale ID
   dwReserved                           As Long    '// Reserved. Must be 0
   sStyle                               As Integer '// Style handle
   wKerning                             As Integer '// Twip size above which to kern Char pair
   bUnderlineType                       As Byte    '// Underline type
   bAnimation                           As Byte    '// Animated text like marching ants
   bRevAuthor                           As Byte    '// Revision author index
   bReserved1                           As Byte
End Type

Private Const C_EM_SCROLL                 As Long = &HB5
Private Const C_EM_GETLINECOUNT           As Long = &HBA
Private Const C_EM_GETFIRSTVISIBLELINE    As Long = &HCE
Private Const C_EM_CHARFROMPOS            As Long = &HD7
Private Const C_EM_GETTEXTRANGE           As Long = 1024 + 75
Private Const C_EM_GETOPTIONS             As Long = 1024 + 78

Private Type InterfaceData
   Class                                 As String
   method                                As String
   IsMethod                              As Boolean
   ToolTip                               As String
End Type
Private mudtInterface()                   As InterfaceData

Private Const C_GREEN                     As Long = 32768

Private Type POINTAPI
   X                                     As Long
   Y                                     As Long
End Type

Private Type EditorWords
   color                                 As Long
   wrd                                   As String
End Type
Private mudtArrWords()                    As EditorWords

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
      (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageP Lib "user32" Alias "SendMessageA" _
      (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As POINTAPI) As Long

Private mblnPaintText                     As Boolean
Private mlngFunctionEndPoint              As Long
Private mblnTextChange                    As Boolean

Private mstrScriptText                    As String
Private mlngPicHeight                     As Long
Private mpicToolTip                       As VB.PictureBox
Private mpicImg(1)                        As stdole.StdPicture
Private WithEvents mrtbScript             As RichTextLib.RichTextBox
Attribute mrtbScript.VB_VarHelpID = -1
Private mrtbDebugTextbox                  As RichTextLib.RichTextBox
Private WithEvents mlvwFlx                As MSComctlLib.ListView
Attribute mlvwFlx.VB_VarHelpID = -1

Private mlngMinPos                        As Long
Private mlngMaxPos                        As Long
Private mstrArrKeywrds()                  As String
Private mlngLastHighLightStart            As Long
Private mlngLastHighlightLen              As Long

Private Sub AddClassKeyword(ByVal vstrCls As String)
   
  Dim lngI      As Long
  Dim blnFound  As Boolean
   
   blnFound = False
   For lngI = 0 To UBound(mstrArrKeywrds)
      blnFound = mstrArrKeywrds(lngI) = vstrCls
      If blnFound Then
         Exit For
      End If
   Next lngI
   
   If Not blnFound Then
      lngI = UBound(mstrArrKeywrds)
      If mstrArrKeywrds(lngI) <> vbNullString Then
         lngI = lngI + 1
      End If
      
      ReDim Preserve mstrArrKeywrds(lngI) As String
      mstrArrKeywrds(lngI) = vstrCls
   End If
   
End Sub

Public Sub AddEditorWord(ByVal vstrWrd As String, ByVal vlngColor As Long)
   
  Dim lngI  As Long
   
   lngI = UBound(mudtArrWords())
   
   If mudtArrWords(0).wrd = vbNullString Then
      lngI = 0
    Else
      lngI = lngI + 1
   End If
   
   ReDim Preserve mudtArrWords(lngI) As EditorWords
   mudtArrWords(lngI).wrd = vstrWrd
   mudtArrWords(lngI).color = vlngColor
   
End Sub

Private Sub AddInterface(ByVal vstrCls As String, ByVal vstrMethod As String, _
                         Optional ByVal vblnIsMethod As Boolean = False, _
                         Optional ByVal vstrToolTip As String = vbNullString)
   
  Dim lngI As Long
   
   lngI = UBound(mudtInterface)
   If mudtInterface(lngI).Class <> vbNullString Then
      lngI = lngI + 1
   End If
   
   ReDim Preserve mudtInterface(lngI) As InterfaceData
   
   With mudtInterface(lngI)
      
      .Class = vstrCls
      .method = vstrMethod
      .IsMethod = vblnIsMethod
      .ToolTip = vstrToolTip
      
      AddClassKeyword LCase$(vstrCls)
      
   End With
   
End Sub

Private Sub APIHighlight(ByVal vlngColor As Long)
   
  Dim udtCF      As CHARFORMAT2
   
   mlngLastHighLightStart = mrtbScript.SelStart
   mlngLastHighlightLen = mrtbScript.SelLength
   '// if you use a Public Variable to store this you have to reset any change or
   '// it will still be set when you make a different call.
   '// EG APIHighLight followed by APIBold
   
   '// both Bolds and Highlights the second instance, If you want to build up a style
   '// with one click
   
   '// this might be useful but mostly not such a good thing.
   On Error GoTo oops
   With udtCF
      .cbSize = LenB(udtCF) '// setup the size of the character format
      .dwMask = C_CFM_BACKCOLOR '// what to test
      If .crBackColor = 0 Then
         .crBackColor = vlngColor
       Else '// NOT .CRBACKCOLOR...
         .crBackColor = mrtbScript.BackColor
      End If
      Call SendMessage(mrtbScript.hWnd, C_EM_SETCHARFORMAT, C_SCF_SELECTION, udtCF)
   End With '// udtCF
   
   Exit Sub
   
oops:
   MsgBox Err.Description, vbCritical
   
End Sub

Private Sub Class_Initialize()
   
   mblnPaintText = False
   mblnTextChange = False
   
   ReDim mudtArrWords(0) As EditorWords
   ReDim mudtInterface(0) As InterfaceData
   ReDim mstrArrKeywrds(0) As String
   
End Sub

Public Sub ClearScript()
   
   mrtbScript.Text = vbNullString
   
End Sub

Private Sub ColorRow(ByVal vObjRTF As RichTextLib.RichTextBox, ByVal vstrWrd As String, _
                     ByVal vlngColor As Long)
   
  Dim lngStart      As Long
  Dim lngEnd        As Long
  Dim lngLineStart  As Long
  Dim strRTF        As String
  Dim strRow        As String
  Dim strFullRow    As String
  Dim strNewRow     As String
  Dim lngColor      As Long
  Dim lngA          As Long
  Dim lngB          As Long
   
   With vObjRTF
      
      strRTF = mstrScriptText
      lngStart = 1
      
      Do
         lngStart = InStr(lngStart, strRTF, vstrWrd, vbTextCompare)
         If lngStart > 0 Then
            '// Check if its hebrew letter
            If Mid$(strRTF, lngStart - 1, 1) <> "\" Then
               '// Check and update the vlngColor table
               If lngColor = 0 Then
                  lngColor = SetColorTable(strRTF, vlngColor)
               End If
               
               lngStart = InStr(lngStart, strRTF, vstrWrd, vbTextCompare)
               lngEnd = InStr(lngStart, strRTF, "\par")
               lngLineStart = InStrRev(strRTF, "\par", lngStart)
               
               If lngEnd > 0 Then
                  strRow = Mid$(strRTF, lngStart, lngEnd - lngStart)
                  strNewRow = "\cf" & lngColor & " " & GetCleanRow(strRow) & "\cf0 "
                  strFullRow = Mid$(strRTF, lngLineStart, lngEnd - lngLineStart)
                  
                  If vstrWrd = "'" Then '// is it lngA comment or is it part of lngA string
                     lngA = InStrRev(strFullRow, Chr$(34))
                     lngB = InStrRev(strFullRow, vstrWrd)
                  End If
                  
                  If lngA < lngB Or lngB = 1 Then
                     strRTF = left(strRTF, lngStart - 1) & _
                           Replace(strRTF, strRow, strNewRow, lngStart, 1)
                  End If
                  
                  lngEnd = InStr(lngEnd - (Len(strRow) - Len(strNewRow)), strRTF, "\par")
               End If
               lngStart = lngEnd
             Else
               lngStart = lngStart + 1
            End If '// Not hebrew letter
         End If '// iStart>0
      Loop Until lngStart = 0
      
      mstrScriptText = strRTF
      
   End With
   
End Sub

Private Sub ColorStrings(ByVal vObjRTF As RichTextLib.RichTextBox, ByVal vstrWrd As String, _
                         ByVal vstrEndWrd As String, ByVal vlngColor As Long)
   
  Dim lngStart        As Long
  Dim lngEnd          As Long
  Dim lngBreak        As Long
  Dim strRTF          As String
  Dim strRow          As String
  Dim strNewRow       As String
  Dim lngColor        As Long
   
   If vstrEndWrd = vbNullString Then
      vstrEndWrd = vstrWrd
   End If
   
   With vObjRTF
      
      strRTF = mstrScriptText
      
      lngStart = 1
      Do
         lngStart = InStr(lngStart, strRTF, vstrWrd, vbTextCompare)
         If lngStart > 0 Then
            
            '// Check && update the vlngColor table
            If lngColor = 0 Then
               lngColor = SetColorTable(strRTF, vlngColor)
            End If
            
            lngStart = InStr(lngStart, strRTF, vstrWrd, vbTextCompare)
            
            If Mid$(strRTF, lngStart - 1, 1) = "\" Then
               lngStart = lngStart + 1
             Else
               lngEnd = GetStringEnd(strRTF, lngStart + 1, vstrEndWrd)
               lngBreak = InStr(lngStart + 1, strRTF, vbNewLine)
               
               If (lngEnd > 0) And ((lngBreak > lngEnd) Or (lngBreak = 0)) Then
                  strRow = Mid$(strRTF, lngStart, lngEnd + Len(vstrEndWrd) - lngStart)
                  strNewRow = GetCleanRow(strRow)
                  strRTF = left(strRTF, lngStart - 1) & "\cf" & lngColor & " " & strNewRow _
                        & "\cf0 " & right(strRTF, Len(strRTF) - lngStart - Len(strRow) + 1)
                  
                  lngEnd = InStr(lngEnd + Len("\cf" & lngColor & " " & strNewRow & "\cf0") _
                        - Len(strNewRow), strRTF, vstrWrd)
               End If
               lngStart = lngEnd
            End If
         End If
      Loop Until lngStart = 0
      
      mstrScriptText = strRTF
      
   End With
   
End Sub

Public Sub ColorWord(ByVal vObjRTF As RichTextLib.RichTextBox, ByVal vlngStartIndx As Long, _
                     ByVal vlngIndxCount As Long, ByVal vlngColor As Long)
   
  Dim lngI         As Long
  Dim lngStart     As Long
  Dim strRTF       As String
  Dim lngColor     As Long
  Dim lngVeryStart As Long
   
   With vObjRTF
      
      strRTF = mstrScriptText
      lngVeryStart = InStr(1, strRTF, "}}")
      
      For lngI = vlngStartIndx To vlngStartIndx + vlngIndxCount
         
         lngStart = InStr(1, strRTF, mudtArrWords(lngI).wrd, vbTextCompare)
         
         If lngStart > 0 And lngStart > lngVeryStart Then
            
            '// Check && update the vlngColor table
            If lngColor = 0 Then
               lngColor = SetColorTable(strRTF, vlngColor)
            End If
            
            If right(mudtArrWords(lngI).wrd, 1) = "." Then
               strRTF = Replace(strRTF, " " & mudtArrWords(lngI).wrd, " \cf" & lngColor & "" _
                     & " " & mudtArrWords(lngI).wrd & "\cf0 ", , , vbTextCompare)
               
             Else
               '// vlngColor the word
               strRTF = Replace(strRTF, " " & mudtArrWords(lngI).wrd & " ", " \cf" & lngColor _
                     & " " & mudtArrWords(lngI).wrd & " \cf0 ", , , vbTextCompare)
               
               '// vlngColor the word + vbnewline
               strRTF = Replace(strRTF, " " & mudtArrWords(lngI).wrd & vbNewLine, " \cf" & _
                     lngColor & " " & mudtArrWords(lngI).wrd & " \cf0", , , vbTextCompare)
               
               strRTF = Replace(strRTF, vbNewLine & mudtArrWords(lngI).wrd & " ", vbNewLine _
                     & " \cf" & lngColor & " " & vbNewLine & mudtArrWords(lngI).wrd & " \cf0", _
                     , , vbTextCompare)
               
            End If
          Else
            If lngStart > 0 Then
               If Mid$(strRTF, lngStart - 1, 1) <> " " Or Mid$(strRTF, lngStart + 1, 1) <> _
                        " " Then
                  
                  '// Check && update the vlngColor table
                  If lngColor = 0 Then
                     lngColor = SetColorTable(strRTF, vlngColor)
                  End If
                  
                  '// vlngColor the word
                  strRTF = Replace(strRTF, " " & mudtArrWords(lngI).wrd & " ", " \cf" & lngColor _
                        & " " & mudtArrWords(lngI).wrd & " \cf0 ", , , vbTextCompare)
                  
                  '// vlngColor the word + vbnewline
                  strRTF = Replace(strRTF, " " & mudtArrWords(lngI).wrd & vbNewLine, " \cf" _
                        & lngColor & " " & mudtArrWords(lngI).wrd & " \cf0", , , vbTextCompare)
                  
                  strRTF = Replace(strRTF, vbNewLine & mudtArrWords(lngI).wrd & " ", vbNewLine _
                        & " \cf" & lngColor & " " & vbNewLine & mudtArrWords(lngI).wrd & _
                        " \cf0", , , vbTextCompare)
                  
               End If
            End If
         End If
      Next lngI
      mstrScriptText = strRTF
   End With
   
End Sub

Private Function CursorOverideFunc() As Boolean
   
   CursorOverideFunc = (mrtbScript.SelStart <= mlngFunctionEndPoint)
   
End Function

Private Function DelColorTable(ByVal vScriptText As String) As String
   
  Dim lngStart As Long
  Dim lngEnd   As Long
   
   '// Do we have color table?
   lngStart = InStr(1, vScriptText, "colortbl", vbTextCompare)
   If lngStart > 0 Then
      lngStart = InStr(1, vScriptText, "{\colortbl", vbTextCompare)
      lngEnd = InStr(lngStart, vScriptText, ";}", vbTextCompare)
      DelColorTable = Replace(vScriptText, Mid$(vScriptText, lngStart, lngEnd - lngStart + 2), vbNullString)
   End If
   
End Function

Private Function GetCleanRow(ByVal vstrRow As String) As String
   
  Dim lngStart As Long
  Dim lngEnd   As Long
  Dim strTemp   As String
   
   GetCleanRow = vstrRow
   lngStart = 1
   Do
      lngStart = InStr(lngStart, vstrRow, "\cf")
      If lngStart > 0 Then
         lngEnd = InStr(lngStart, vstrRow, " ")
         If lngStart > 0 Then
            If lngEnd = 0 Then
               lngEnd = Len(vstrRow) + 1
             Else
               lngEnd = lngEnd + 1
            End If
            strTemp = Mid$(vstrRow, lngStart, lngEnd - lngStart)
            
            '// Ignore special RTF chars
            Select Case Mid$(strTemp, 1, 2)
             Case "\\", "\{", "\}", "\'"
               lngStart = lngStart + 2
             Case Else
               vstrRow = Replace(vstrRow, strTemp, vbNullString)
            End Select
         End If
      End If
      
   Loop Until lngStart = 0
   
   GetCleanRow = vstrRow
   
End Function

Private Function GetColorTable() As String
   
  Dim lngStart As Long
  Dim lngEnd   As Long
   
   '// Do we have color table?
   lngStart = InStr(1, mrtbScript.TextRTF, "colortbl", vbTextCompare)
   If lngStart > 0 Then
      lngStart = InStr(1, mrtbScript.TextRTF, "{\colortbl", vbTextCompare)
      lngEnd = InStr(lngStart, mrtbScript.TextRTF, ";}", vbTextCompare)
      GetColorTable = Mid$(mrtbScript.TextRTF, lngStart, lngEnd - lngStart + 2)
   End If
   
End Function

Public Function GetCurrentPosition(ByRef rRTF As RichTextBox) As Long
   
   GetCurrentPosition = SendMessage(rRTF.hWnd, C_EM_GETTEXTRANGE, ByVal 0&, ByVal 0&)
   
End Function

Private Function GetFixedToolTip(ByVal vstrToolTip As String, ByVal vlngDelta As Long, _
                                 ByVal vlngStartPos As Long) As String
   
  Dim lngI       As Long
  Dim strTemp    As String
  Dim blnPutChar As Boolean
  Dim blnStop    As Boolean
  Dim lngLastPos As Long
  Dim lngLastCut As Long
  Dim strCopy    As String
  Dim blnHardCut As Boolean
   
   blnPutChar = False
   blnStop = False
   blnHardCut = False
   
   strCopy = vstrToolTip
   vstrToolTip = StrReverse(vstrToolTip)
   vlngDelta = Abs(vlngDelta)
   
   Do
      
      lngI = 0
      lngLastPos = 0
      vlngDelta = Abs(vlngDelta)
      
      Do
         lngI = InStr(lngI + 1, vstrToolTip, ",")
         If lngI > 0 Then
            If lngLastCut = lngI Then
               lngI = lngLastPos
               blnHardCut = True
            End If
            
            lngLastPos = lngI
            
            If lngI = 0 Then
               blnStop = True
             Else
               strTemp = left(vstrToolTip, lngI)
               If ((mpicToolTip.TextWidth(strTemp) >= vlngDelta) Or (blnHardCut)) Then
                  blnHardCut = False
                  '// OK this is the correct place to put the vbNewline
                  strCopy = left(strCopy, Len(strCopy) - lngI) & vbNewLine & right(strCopy, lngI)
                  
                  lngLastCut = lngI
                  blnPutChar = True
                  vlngDelta = Screen.Width - (mpicToolTip.TextWidth(strTemp) + vlngStartPos)
                  blnStop = CBool(vlngDelta >= 0)
                  lngI = 0
               End If
            End If
         End If
      Loop Until lngI = 0
      
      '// If we deednt put anything, put it in the last position found
      If (lngI = 0) And (Not blnPutChar) Then
         
         If lngLastPos = 0 Then
            blnStop = True
          Else
            strCopy = left(strCopy, Len(strCopy) - lngLastPos) & vbNewLine & right(strCopy, _
                  lngLastPos)
            
            vlngDelta = Screen.Width - (mpicToolTip.TextWidth(left(strCopy, Len(strCopy) _
                  - lngLastPos)) + vlngStartPos)
            
            blnStop = CBool(vlngDelta >= 0)
            vstrToolTip = right(vstrToolTip, Len(vstrToolTip) - lngLastPos)
         End If
      End If
   Loop Until blnStop
   
   GetFixedToolTip = strCopy
   
End Function

Private Function GetKeyword() As String
   
  Dim lngI      As Long
  Dim lngPos    As Long
  Dim strKeywrd As String
  Dim strTemp   As String
   
   With mrtbScript
      
      lngPos = .SelStart
      
      For lngI = lngPos To 1 Step -1
         strTemp = Mid$(.Text, lngI, 1)
         
         Select Case Asc(strTemp)
          Case 97 To 122
          Case 65 To 90
          Case 48 To 57
          Case Else
            Exit For
         End Select
         
         strKeywrd = strKeywrd & strTemp
      Next lngI
      
   End With
   
   strKeywrd = StrReverse(strKeywrd)
   If Not KeyWordExsist(strKeywrd) Then
      strKeywrd = vbNullString
   End If
   
   '// retval
   GetKeyword = strKeywrd
   
End Function

Public Function GetKeywords() As String()
   
   GetKeywords = mstrArrKeywrds()
   
End Function

Private Sub GetRGB(ByVal vlngColor As Long, _
                   ByRef rlngRed As Long, ByRef rlngGreen As Long, ByRef rlngBlue As Long)
   
   rlngBlue = Int(vlngColor / 65536)
   rlngGreen = Int((vlngColor / 65536 - rlngBlue) * 65536 / 256)
   rlngRed = Int(((vlngColor / 65536 - rlngBlue) * 65536 / 256 - rlngGreen) * 256)
   
End Sub

Private Function GetStringEnd(ByVal vstrRTF As String, ByVal vlngStart As Long, _
                              ByVal vstrEndWrd As String) As Long
   
  Dim blnStop As Boolean
   
   '// return the string end position?
   
   blnStop = False
   Do Until blnStop
      GetStringEnd = InStr(vlngStart, vstrRTF, vstrEndWrd, vbTextCompare)
      '// first, find the string limit character (' or ")
      If GetStringEnd > 0 Then
         '// find one, now check for backslash
         blnStop = CBool(Mid$(vstrRTF, GetStringEnd - 1, 1) <> "\")
         '// is there a backslash
         If Not blnStop Then
            '// yes, is there another one before this backslash?
            '// if yes, then stop the search. if no, continue
            blnStop = CBool(Mid$(vstrRTF, GetStringEnd - 2, 1) = "\")
         End If
       Else
         blnStop = True
      End If
      
      If Not blnStop Then
         vlngStart = GetStringEnd + 1
      End If
   Loop
   
End Function

Public Function GetXMLInterface() As String
   
  Dim strXML As String
  Dim lngI    As Long
   
   strXML = "<ROOT>"
   For lngI = 0 To UBound(mudtInterface)
      With mudtInterface(lngI)
         strXML = strXML & "<NEWITEM>"
         strXML = strXML & "<Class>" & .Class & "</Class>"
         strXML = strXML & "<IsMethod>" & .IsMethod & "</IsMethod>"
         strXML = strXML & "<method>" & .method & "</method>"
         strXML = strXML & "<ToolTip>" & .ToolTip & "</ToolTip>"
         strXML = strXML & "</NEWITEM>"
      End With
   Next lngI
   strXML = strXML & "</ROOT>"
   GetXMLInterface = strXML
   
End Function

Public Sub HighLight(ByVal vlngColor As Long)
   
  Dim lngStart As Long
  Dim lngLen   As Long
   
   If mlngLastHighLightStart > 0 Then
      
      lngStart = mrtbScript.SelStart
      lngLen = mrtbScript.SelLength
      
      mrtbScript.SelStart = mlngLastHighLightStart
      mrtbScript.SelLength = mlngLastHighlightLen
      APIHighlight mrtbScript.BackColor
      
      mrtbScript.SelStart = lngStart
      mrtbScript.SelLength = lngLen
      
   End If
   
   Call APIHighlight(vlngColor)
   mrtbScript.SelLength = 0
   
End Sub

Private Function KeyWordExsist(ByVal vstrKeywrd As String) As Boolean
   
  Dim aKeywrds() As Variant
  Dim lngI            As Long
   
   KeyWordExsist = False
   
   vstrKeywrd = LCase$(vstrKeywrd)
   
   For lngI = 0 To UBound(mstrArrKeywrds)
      KeyWordExsist = CBool(mstrArrKeywrds(lngI) = vstrKeywrd)
      If KeyWordExsist Then
         Exit For
      End If
   Next lngI
   
End Function

Public Function LoadIntelisence(ByVal vstrInterfaceFile As String) As Boolean
   
  Dim lngFN       As Long
  Dim strTemp     As String
  Dim varArr      As Variant
  Dim strToolTip  As String
  Dim blnMethod   As Boolean
   
   lngFN = FreeFile
   Open vstrInterfaceFile For Input As #lngFN
   Do Until EOF(lngFN)
      Line Input #lngFN, strTemp
      varArr = Split(strTemp, ";")
      
      If UBound(varArr) >= 2 Then
         blnMethod = CBool(varArr(2) = "1")
       Else
         blnMethod = False
      End If
      
      If UBound(varArr) >= 3 Then
         strToolTip = varArr(3)
       Else
         strToolTip = vbNullString
      End If
      
      AddInterface varArr(0), varArr(1), blnMethod, strToolTip
      
   Loop
   Close #lngFN
   
End Function

Private Sub mlvwFlx_DblClick()
   
   Call TypeIntelisenceText
   
End Sub

Private Sub mlvwFlx_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   Call ShowToolTip
   
End Sub

Private Sub mlvwFlx_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Select Case KeyCode
    Case vbKeyReturn, vbKeySpace
      TypeIntelisenceText
      SendKeys (" ")
    Case vbKeyEscape
      mlvwFlx.Visible = False
      mpicToolTip.Visible = False
      mrtbScript.SetFocus
    Case vbKeyBack
      mlvwFlx.Visible = False
      mpicToolTip.Visible = False
      mrtbScript.SetFocus
      SendKeys ("{BS}")
    Case 187 '// =
      TypeIntelisenceText
      mrtbScript.SetFocus
      SendKeys ("=")
      
   End Select
   
End Sub

Private Sub mlvwFlx_LostFocus()
   
   mlvwFlx.Visible = False
   With mpicToolTip
      .top = mlvwFlx.top + .TextHeight("A")
   End With
   
End Sub

Private Sub mrtbScript_Change()
   
   If mblnPaintText Then
      mblnPaintText = False
      
      If mblnTextChange Then
         mrtbDebugTextbox.SetFocus
         PaintText
         mrtbScript.SetFocus
         mblnTextChange = False
      End If
   End If
   
End Sub

Public Sub PaintText(Optional vblnMarkStrings As Boolean = False, _
                     Optional vblnMarkREM As Boolean = False)
   
  Dim lngPos       As Long
  Dim lngI         As Long
  Dim lngCnt       As Long
  Dim lngBound     As Long
  Dim lngStart     As Long
  Dim strWorkText  As String
  Dim lngTextStart As Long
   
   lngCnt = 0
   lngStart = 0
   lngPos = mrtbScript.SelStart
   
   mstrScriptText = mrtbScript.TextRTF
   strWorkText = mstrScriptText
   
   lngBound = UBound(mudtArrWords())
   For lngI = 0 To lngBound
      If LenB(mudtArrWords(lngI).wrd) = 0 Then
         Exit For
      End If
      '// Is it the last word to paint ?
      If lngBound = lngI Then
         '// Yes it is
         ColorWord mrtbScript, lngStart, lngCnt, mudtArrWords(lngI).color
       Else
         If mudtArrWords(lngI + 1).color <> mudtArrWords(lngI).color Then
            ColorWord mrtbScript, lngStart, lngCnt, mudtArrWords(lngI).color
            lngStart = lngI + 1
            lngCnt = 0
          Else
            lngCnt = lngCnt + 1
         End If
      End If
   Next lngI
   
   If vblnMarkStrings Then ColorStrings mrtbScript, Chr$(34), Chr$(34), vbMagenta
   If vblnMarkREM Then ColorRow mrtbScript, "'", C_GREEN
   
   If lngTextStart = 0 Then
      mrtbScript.TextRTF = mstrScriptText
    Else
      mrtbScript.TextRTF = Replace(mrtbScript.TextRTF, strWorkText, DelColorTable(mstrScriptText))
   End If
   
   If lngPos > 0 Then
      lngPos = InStr(lngPos, mrtbScript.Text, Chr$(10))
   End If
   
   mrtbScript.SelStart = lngPos
   
   mlngMinPos = mrtbScript.GetLineFromChar(lngPos)
   mlngMaxPos = mlngMinPos
   
End Sub

Private Function SetColorTable(ByRef rstrRTF As String, ByVal vlngColor As Long) As Long
   
  Dim lngRed   As Long
  Dim lngGreen As Long
  Dim lngBlue  As Long
  Dim lngI     As Long
  Dim lngStart As Long
  Dim lngEnd   As Long
  Dim strTable As String
  Dim strTmp   As String
  Dim strColor As String
   
   GetRGB vlngColor, lngRed, lngGreen, lngBlue
   strColor = "\red" & lngRed & "\green" & lngGreen & "\blue" & lngBlue & ";"
   
   '// Do we have vlngColor table?
   lngStart = InStr(1, left$(rstrRTF, 200), "colortbl", vbTextCompare)
   If lngStart = 0 Then
      '// No, Insert vlngColor table
      rstrRTF = Replace(rstrRTF, "}}", "}}{\colortbl ;}", , 1)
      lngStart = InStr(1, rstrRTF, "colortbl", vbTextCompare)
   End If
   
   '// Do we have the current vlngColor?
   If InStr(1, rstrRTF, strColor, vbTextCompare) = 0 Then
      '// No, insert vlngColor
      lngI = InStr(1, rstrRTF, "\colortbl")
      lngI = InStr(lngI, rstrRTF, "}", vbTextCompare)
      strTmp = left(rstrRTF, lngI - 1)
      rstrRTF = Replace(rstrRTF, "}", strColor & "}", lngI, 1, vbTextCompare)
      rstrRTF = strTmp & rstrRTF
   End If
   
   lngEnd = InStr(lngStart, rstrRTF, ";}")
   strTable = Mid$(rstrRTF, lngStart, lngEnd - lngStart + 1)
   
   lngI = 0
   lngStart = 1
   Do
      lngStart = InStr(lngStart, strTable, ";")
      If lngStart <> 0 Then
         lngI = lngI + 1
         lngStart = lngStart + 1
         If InStr(lngStart, strTable, strColor, vbTextCompare) = lngStart Then
            SetColorTable = lngI
            Exit Do
         End If
      End If
   Loop Until lngStart = 0
   
End Function

Public Sub SetEditorObjects(Optional ByRef rObjRTF As RichTextLib.RichTextBox, _
                            Optional ByRef rObjGrid As MSComctlLib.ListView, _
                            Optional ByRef rImg0 As stdole.StdPicture, _
                            Optional ByRef rImg1 As stdole.StdPicture, _
                            Optional ByRef rpicToolTip As VB.PictureBox, _
                            Optional ByRef rDebugTextbox As RichTextLib.RichTextBox)
   
   '// Init editor objects
   
   Set mrtbScript = rObjRTF '// main script text box
   Set mrtbDebugTextbox = rDebugTextbox '// need this for smooth painting
   Set mlvwFlx = rObjGrid '// intelisence grid
   
   If Not (mlvwFlx Is Nothing) Then
      mlvwFlx.ColumnHeaders.Add 1, , , mlvwFlx.Width - 3 * Screen.TwipsPerPixelX
      mlvwFlx.FullRowSelect = True
   End If
   
   Set mpicImg(0) = rImg0 '// intelisence pictures
   Set mpicImg(1) = rImg1
   Set mpicToolTip = rpicToolTip '// tool tip object
   
   If Not (mpicToolTip Is Nothing) Then
      mlngPicHeight = mpicToolTip.Height
   End If
   Call ClearScript
   
End Sub

Public Sub ShowIntelisence()
   
  Dim sKeyword As String
   
   sKeyword = GetKeyword()
   If sKeyword <> vbNullString Then
      ShowInterface sKeyword
   End If
   
End Sub

Private Sub ShowInterface(ByVal vstrKeywrd As String)
   
  Dim objItem As MSComctlLib.ListItem
  Dim lngI       As Long
  Dim udtPoint As POINTAPI
   
   If mudtInterface(0).Class = vbNullString Then
      Exit Sub
   End If
   
   With mlvwFlx
      
      .ListItems.Clear
      
      For lngI = 0 To UBound(mudtInterface)
         If LCase$(mudtInterface(lngI).Class) = LCase$(vstrKeywrd) Then
            Set objItem = .ListItems.Add(, , mudtInterface(lngI).method)
            objItem.Tag = mudtInterface(lngI).ToolTip
            objItem.SmallIcon = IIf(mudtInterface(lngI).IsMethod, "Method", "Property")
         End If
      Next lngI
      
      GetCaretPos udtPoint
      .left = mrtbScript.left + (udtPoint.X + 10) * Screen.TwipsPerPixelX
      If .left + .Width > Screen.Width Then
         .left = Screen.Width - .Width
      End If
      
      .top = mrtbScript.top + (udtPoint.Y + 5) * Screen.TwipsPerPixelY
      
      '// Set grid height
      
      If .ListItems.Count >= 6 Then
         .Height = 6 * .ListItems.Item(1).Height
       Else
         .Height = .ListItems.Count * .ListItems.Item(1).Height + 50
      End If
      
      .ListItems.Item(1).Selected = True
      .Visible = True
      .SetFocus
      
      Call ShowToolTip
      
   End With
   
End Sub

Public Sub ShowToolTip()
   
  Dim strToolTip As String
  Dim lngCount   As Long
  Dim lngPos     As Long
  Dim lngI       As Long
  Dim lngDelta   As Long
   
   With mlvwFlx
      
      '// Show tool tip
      strToolTip = " " & .SelectedItem.Tag
      
      mpicToolTip.left = .left
      mpicToolTip.top = .top + .Height
      mpicToolTip.Height = mlngPicHeight
      mpicToolTip.Width = mpicToolTip.TextWidth(strToolTip)
      
      '// Insert vbNewLine where needed
      lngDelta = Screen.Width - (.left + mpicToolTip.Width)
      If lngDelta < 0 Then
         strToolTip = GetFixedToolTip(strToolTip, lngDelta, .left)
      End If
      
      If Trim$(strToolTip) <> vbNullString Then
         
         lngPos = 0
         lngCount = 0
         Do
            lngPos = InStr(lngPos + 1, strToolTip, vbNewLine)
            If lngPos <> 0 Then
               lngCount = lngCount + 1
            End If
            
         Loop Until lngPos = 0
         
         If lngCount > 0 Then
            For lngI = 1 To lngCount
               mpicToolTip.Height = mpicToolTip.Height + mlngPicHeight
            Next lngI
         End If
         
         mpicToolTip.Width = mpicToolTip.TextWidth(strToolTip)
         
         If mpicToolTip.Width + .left > Screen.Width Then
            mpicToolTip.left = Screen.Width - mpicToolTip.Width
         End If
         
         mpicToolTip.Cls
         mpicToolTip.Print strToolTip
         
         mpicToolTip.Visible = True
       Else
         mpicToolTip.Visible = False
      End If
      
   End With
   
End Sub

Public Sub TypeIntelisenceText()
   
   mrtbScript.SelText = mlvwFlx.SelectedItem.Text '//   mlvwFlx.TextMatrix(mlvwFlx.Row, 1)
   mrtbScript.SetFocus
   mlvwFlx.Visible = False
   '// Copy the function position on the screen
   mlngFunctionEndPoint = mrtbScript.SelStart '// + Len(mlvwFlx.TextMatrix(mlvwFlx.Row, 1))
   
End Sub

