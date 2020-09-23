VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 1998-2003
'//************************************/
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" ( _
   ByVal hdc As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, _
   ByVal ySrc As Long, _
   ByVal nSrcWidth As Long, _
   ByVal nSrcHeight As Long, _
   ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" ( _
   ByVal hdc As Long, _
   ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'// Flag indicating Printing or Previewing
Private PrintFlag As Boolean

'// Object used for Print Preview
Private ObjPrint As Control

'// Storage for the Printer's orignal scale mode
Private pSM As Integer
'// Storage for the Object's orignal scale mode
Private oSM As Integer
'// Default Scale Mode
Private oScaleMode As Integer

'// The actual printable area (something a little less then the paper size)
Private PgWidth As Single
Private PgHeight As Single
Private oOrientation As Integer

'// Remember ColorMode
Private oColorMode As Byte

'// Remember Header Information
Private oTitleMain As String
Private oTitleSub As String
Private oTitleItalic As Boolean

Private PageNumber As Integer
Private TempDir As String

Public Enum PageOrientation
   PagePortrait = vbPRORPortrait
   PageLandscape = vbPRORLandscape
End Enum

Public Enum PrinterColorModeTypes
   cmMonochrome = vbPRCMMonochrome
   cmColor = vbPRCMColor
End Enum

'// -----rotate fonts
Private Const LF_FACESIZE = 32
Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" ( _
   ByVal hdc As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal lpString As String, _
   ByVal nCount As Long) As Long

Public Property Get BackColor() As Long
   
   If PrintFlag Then
      BackColor = GetBkColor(Printer.hdc)
   Else
      BackColor = ObjPrint.BackColor
   End If

End Property

Public Property Let BackColor(ByVal NewColor As Long)
  
  Dim r As Long
   
   If PrintFlag Then
      If NewColor = -1 Then
         Call MakeTrans '// I am not sure why this is necessary
         r = SetBkColor(Printer.hdc, vbWhite)
         Printer.FontTransparent = True
      Else
         Printer.FontTransparent = False
         r = SetBkColor(Printer.hdc, NewColor)
      End If
   Else
      If NewColor = -1 Then
         ObjPrint.FontTransparent = True
         r = SetBkColor(ObjPrint.hdc, vbWhite)
      Else
         ObjPrint.FontTransparent = False
         r = SetBkColor(ObjPrint.hdc, NewColor)
      End If
   End If
   
End Property

Private Sub Class_Initialize()
   
   On Error GoTo Err_Proc
   
   ColorMode = cmColor
   Orientation = PagePortrait
   PrintCopies = 1
   SendToPrinter = True
   oTitleItalic = True
   
   '// Default Scale Mode
   '// vbInches or vbCentimeters
   oScaleMode = vbInches
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "Class_Initialize"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Property Get ColorMode() As PrinterColorModeTypes
   ColorMode = oColorMode
End Property

Public Property Let ColorMode(ByVal vNewValue As PrinterColorModeTypes)
   oColorMode = vNewValue
   If PrintFlag Then Printer.ColorMode = vNewValue
End Property

Public Property Get CurrentX() As Single
   If PrintFlag Then
      CurrentX = Printer.CurrentX
   Else
      CurrentX = ObjPrint.CurrentX
   End If
End Property

Public Property Let CurrentX(ByVal NewXvalue As Single)
   If PrintFlag Then
      Printer.CurrentX = NewXvalue
   Else
      ObjPrint.CurrentX = NewXvalue
   End If
End Property

Public Property Get CurrentY() As Single
   If PrintFlag Then
      CurrentY = Printer.CurrentY
   Else
      CurrentY = ObjPrint.CurrentY
   End If
End Property

Public Property Let CurrentY(ByVal NewYvalue As Single)
   If PrintFlag Then
      Printer.CurrentY = NewYvalue
   Else
      ObjPrint.CurrentY = NewYvalue
   End If
End Property

Public Property Get DrawStyle() As DrawStyleConstants
   
   DrawStyle = Printer.DrawStyle
   
End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)
   
   On Error Resume Next
   Printer.DrawStyle = vNewValue
   ObjPrint.DrawStyle = vNewValue
   
End Property

Public Property Get DrawWidth() As Integer
   
   If PrintFlag Then
      DrawWidth = Printer.DrawWidth
   Else
      DrawWidth = ObjPrint.DrawWidth * 2
   End If

End Property

Public Property Let DrawWidth(ByVal NewWidth As Integer)
   
   If NewWidth < 1 Then NewWidth = 1
   If PrintFlag Then
      Printer.DrawWidth = NewWidth
   Else
      NewWidth = NewWidth / 2
      If NewWidth < 1 Then NewWidth = 1
      ObjPrint.DrawWidth = NewWidth
   End If
   
End Property

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
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
       App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
       Environ("username"), Environ("computername")
   Close #lngFN
   
End Sub

Public Property Get FontBold() As Boolean
   If PrintFlag Then
      FontBold = Printer.FontBold
   Else
      FontBold = ObjPrint.FontBold
   End If
End Property

Public Property Let FontBold(YesNo As Boolean)
   If PrintFlag Then
      Printer.FontBold = YesNo
   Else
      ObjPrint.FontBold = YesNo
   End If
End Property

Public Property Get FontItalic() As Boolean
   If PrintFlag Then
      FontItalic = Printer.FontItalic
   Else
      FontItalic = ObjPrint.FontItalic
   End If
End Property

Public Property Let FontItalic(YesNo As Boolean)
   If PrintFlag Then
      Printer.FontItalic = YesNo
   Else
      ObjPrint.FontItalic = YesNo
   End If
End Property

Public Property Get FontName() As String
   If PrintFlag Then
      FontName = Printer.FontName
   Else
      FontName = ObjPrint.FontName
   End If
End Property

Public Property Let FontName(ByVal newFont As String)
   If PrintFlag Then
      Printer.FontName = newFont
      Printer.Print vbNullString;
   Else
      ObjPrint.FontName = newFont
      ObjPrint.Print vbNullString;
   End If
End Property

Public Property Get FontSize() As Integer
   If PrintFlag Then
      FontSize = Printer.FontSize
   Else
      FontSize = ObjPrint.FontSize
   End If
End Property

Public Property Let FontSize(pSize As Integer)
   If PrintFlag Then
      Printer.FontSize = pSize
   Else
      ObjPrint.FontSize = pSize
   End If
End Property

Public Property Get FontStrikethru() As Boolean
   If PrintFlag Then
      FontStrikethru = Printer.FontStrikethru
   Else
      FontStrikethru = ObjPrint.FontStrikethru
   End If
End Property

Public Property Let FontStrikethru(YesNo As Boolean)
   If PrintFlag Then
      Printer.FontStrikethru = YesNo
   Else
      ObjPrint.FontStrikethru = YesNo
   End If
End Property

Public Property Get FontTransparent() As Boolean
   If PrintFlag Then
      FontTransparent = Printer.FontTransparent
   Else
      FontTransparent = ObjPrint.FontTransparent
   End If
End Property

Public Property Let FontTransparent(ByVal vNewValue As Boolean)
   If PrintFlag Then
      If vNewValue Then
         '// I am not sure why this is necessary but it doesn't work without it
         Call MakeTrans
         Me.BackColor = -1
      Else
         Printer.FontTransparent = False
      End If
   Else
      ObjPrint.FontTransparent = vNewValue
   End If
End Property

Public Property Get FontUnderline() As Boolean
   If PrintFlag Then
      FontUnderline = Printer.FontUnderline
   Else
      FontUnderline = ObjPrint.FontUnderline
   End If
End Property

Public Property Let FontUnderline(YesNo As Boolean)
   If PrintFlag Then
      Printer.FontUnderline = YesNo
   Else
      ObjPrint.FontUnderline = YesNo
   End If
End Property

Public Property Get ForeColor() As Long
   If PrintFlag Then
      ForeColor = Printer.ForeColor
   Else
      ForeColor = ObjPrint.ForeColor
   End If
End Property

Public Property Let ForeColor(NewColor As Long)
   If PrintFlag Then
      Printer.ForeColor = NewColor
   Else
      ObjPrint.ForeColor = NewColor
   End If
End Property

Public Function GetFormalCase(ByVal TextString As String) As String
  Dim x As Integer
   
   '// Cap the first letter if each word
   On Local Error Resume Next
   
   TextString = UCase$(left$(TextString, 1)) & LCase$(Mid$(TextString, 2))
   
   '// Look for space
   x = InStr(TextString, " ")
   If x Then
      Do
         Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
         x = x + 1
         x = InStr(x, TextString, " ")
         If x = 0 Or x + 1 > Len(TextString) Then Exit Do
      Loop
   End If
   
   '// Look for .
   x = InStr(TextString, ".")
   If x Then
      Do
         Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
         x = x + 1
         x = InStr(x, TextString, ".")
         If x = 0 Or x + 1 > Len(TextString) Then Exit Do
      Loop
   End If
   
   '// Look for Mc
   x = InStr(TextString, "Mc")
   If x Then
      Do
         Mid$(TextString, x + 2, 1) = UCase$(Mid$(TextString, x + 2, 1))
         x = x + 2
         x = InStr(x, TextString, "Mc")
         If x = 0 Or x + 2 > Len(TextString) Then Exit Do
      Loop
   End If
   
   '// Look for O'
   x = InStr(TextString, "O'")
   If x Then
      Do
         Mid$(TextString, x + 2, 1) = UCase$(Mid$(TextString, x + 2, 1))
         x = x + 2
         x = InStr(x, TextString, "O'")
         If x = 0 Or x + 2 > Len(TextString) Then Exit Do
      Loop
   End If
   
   '// Look for -
   x = InStr(TextString, "-")
   If x Then
      Do
         Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
         x = x + 1
         x = InStr(x, TextString, "-")
         If x = 0 Or x + 1 > Len(TextString) Then Exit Do
      Loop
   End If
   
   GetFormalCase = LTrim$(TextString)
   
End Function

Public Function GetPage() As Variant
   If PrintFlag Then
      GetPage = Printer.Page
   Else
      GetPage = PageNumber + 1
   End If
End Function

Public Function GetPaperHeight() As Single
   GetPaperHeight = PgHeight
End Function

Public Function GetPaperWidth() As Single
   GetPaperWidth = PgWidth
End Function

Public Function GetRemoveCRLF(ByVal TextString As String) As String
   
  Dim i As Long
  Dim FoundString As Boolean
  Dim FoundFirst As Boolean
   
   On Error GoTo Err_Proc
   
   Do
      FoundString = False
      
      i = InStr(TextString, vbCr)
      If i Then
         Mid$(TextString, i, 1) = " "
         FoundString = True
         FoundFirst = True
      End If
      
      i = InStr(TextString, vbLf)
      If i = 1 Then
         TextString = Mid$(TextString, i + 1)
      ElseIf i > 1 Then
         If FoundFirst Then
            TextString = Mid$(TextString, 1, i - 1) & Mid$(TextString, i + 1)
         Else
            Mid$(TextString, i, 1) = " "
         End If
         FoundString = True
      End If
      FoundFirst = False
      
   Loop Until FoundString = False
   GetRemoveCRLF = TextString
   
Exit_Here:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "GetRemoveCRLF"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Function GetStripQuotes(ByVal TextString As String) As String
   If left$(TextString, 1) = Chr(34) Then TextString = Mid$(TextString, 2)
   If right$(TextString, 1) = Chr(34) Then TextString = left$(TextString, Len(TextString) - 1)
   GetStripQuotes = TextString
End Function

Public Function GetTextHeight(Optional TextString As String = "Sample Text") As Single
   If PrintFlag Then
      GetTextHeight = Printer.TextHeight(TextString)
   Else
      GetTextHeight = ObjPrint.TextHeight(TextString)
   End If
End Function

Public Function GetTextWidth(TextString As Variant) As Single
   If PrintFlag Then
      GetTextWidth = Printer.TextWidth(TextString)
   Else
      GetTextWidth = ObjPrint.TextWidth(TextString)
   End If
End Function

Private Sub MakeTrans()
   
  Dim x As Single
  Dim y As Single
   
   On Error GoTo Err_Proc
   
   x = Me.CurrentX
   y = Me.CurrentY
   '// I am not sure why this is necessary
   '// but it doesn't work without it.
   Sleep 1
   Me.BackColor = vbWhite
   Sleep 1
   pQuarterSpace
   Sleep 1
   Me.CurrentX = x
   Me.CurrentY = y
   Sleep 1
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "MakeTrans"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Property Get Orientation() As PageOrientation
   Orientation = oOrientation
End Property

Public Property Let Orientation(ByVal vNewValue As PageOrientation)
   On Error Resume Next
   Printer.Orientation = vNewValue
   oOrientation = vNewValue
End Property

Public Sub p15Space()
   
  Dim eFont As Integer
  Dim hFont As Integer
   
   On Error GoTo Err_Proc
   pPrint
   eFont = Me.FontSize
   hFont = eFont \ 2
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "p15Space"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pBox(Optional ByVal bLeft As Single = -1, _
                Optional ByVal bTop As Single = -1, _
                Optional ByVal bWidth As Single = -1, _
                Optional ByVal bHeight As Single = -1, _
                Optional ByVal ColorLine As Long = -1, _
                Optional ByVal ColorFill As Long = -1, _
                Optional FilledBox As FillStyleConstants = vbFSTransparent)
   
  Dim x As Single
  Dim y As Single
   
   On Error GoTo Err_Proc
   
   y = Me.CurrentY
   x = Me.CurrentX
   
   If ColorLine = -1 Then ColorLine = Me.ForeColor
   If ColorFill = -1 Then ColorFill = ColorLine
   If bLeft = -1 Then bLeft = Me.CurrentX
   If bTop = -1 Then bTop = Me.CurrentY
   If bWidth = -1 Then bWidth = PgWidth
   If bHeight = -1 Then bHeight = GetTextHeight
   
   If FilledBox <> vbFSTransparent Then
      If PrintFlag Then
         Printer.FillColor = ColorFill
         Printer.FillStyle = FilledBox
         Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
         Printer.FillStyle = vbFSTransparent
      Else
         ObjPrint.FillColor = ColorFill
         ObjPrint.FillStyle = FilledBox
         ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
         ObjPrint.FillStyle = vbFSTransparent
      End If
   Else
      If PrintFlag Then
         Printer.FillStyle = vbFSTransparent
         Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
      Else
         ObjPrint.FillStyle = vbFSTransparent
         ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
      End If
   End If
   
   Me.CurrentX = x
   Me.CurrentY = y
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pBox"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pBullet(Optional ByVal LeftMargin As Single = -1)
   
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean
   
   On Error GoTo Err_Proc
   
   If LeftMargin = -1 Then LeftMargin = Me.CurrentX
   
   '// Save current setting
   eFontN = Me.FontName
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru
   
   '// Change settings
   '// FontName = "Wingdings"
   Me.FontName = "Symbol"
   Me.FontBold = False
   Me.FontItalic = False
   Me.FontUnderline = False
   Me.FontStrikethru = False
   
   'pPrint Chr(164) & "  ", LeftMargin, True
   pPrint Chr(183) & "  ", LeftMargin, True
   
   '// Restore setting
   Me.FontName = eFontN
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pBullet"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pCancled()
   Me.FontSize = 12
   Me.FontBold = True
   Me.ForeColor = vbRed
   pPrint
   pPrint "**** PRINTING CANCLED ****", 0.5
End Sub

Public Sub pCenter(ByVal PrintVar As String, _
                   Optional ByVal SameLine As Boolean = False, _
                   Optional COLWIDTH As Single = -1, _
                   Optional LeftMargin As Single = 0)
   
   On Error GoTo Err_Proc
   
   If COLWIDTH = -1 Then COLWIDTH = PgWidth - LeftMargin
   
   If GetTextWidth(PrintVar) > PgWidth Then
      pCenterMultiline PrintVar, LeftMargin, LeftMargin + COLWIDTH, , SameLine
   Else
      Me.CurrentX = LeftMargin + ((COLWIDTH - GetTextWidth(PrintVar)) / 2)
      pPrint PrintVar, , SameLine
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCenter"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pCenterMultiline(ByVal PrintVar As Variant, _
                            Optional ByVal LeftMargin As Single = 0, _
                            Optional ByVal RightMargin As Single = -1, _
                            Optional ByVal RemoveCrLf As Boolean = True, _
                            Optional ByVal SameLine As Boolean = False)
   
  Dim StartChar As Integer
  Dim CharLength As Single
  Dim CurrentPos As Single
  Dim TxtLen As Single
  Dim TxtWidth As Single
  Dim tString As String
  Dim NeedsStrip As Boolean
  Dim COLWIDTH As Single
  Dim x As Long
  Dim y As Long
   
   On Error GoTo Err_Proc
   
   If RightMargin = -1 Then RightMargin = PgWidth
   COLWIDTH = RightMargin - LeftMargin
   
   TxtLen = Len(PrintVar)
   StartChar = 1
   CurrentPos = 0
   CharLength = TxtLen
   
   If PrintVar = vbNullString Then
      pPrint
      Exit Sub
   End If
   
   If InStr(PrintVar, vbCr) Or InStr(PrintVar, vbLf) Then NeedsStrip = True
   
   For x = 1 To TxtLen
      y = x - CurrentPos
      
      '// Mark space between words
      If Mid$(PrintVar, x, 1) < Chr(33) Then CharLength = y
      
      If (GetTextWidth(Mid$(PrintVar, StartChar, y)) >= COLWIDTH) Or (Not RemoveCrLf And Mid$(PrintVar, x, 1) = vbCr) Then
      
         '// If there are no spaces then break line here */
         If CharLength > y Then CharLength = y - 1
         
         tString = Trim$(Mid$(PrintVar, StartChar, CharLength))
         If NeedsStrip Then tString = GetRemoveCRLF(tString)
         Me.CurrentX = LeftMargin + ((COLWIDTH - GetTextWidth(tString)) / 2)
         If PrintFlag Then
            Printer.Print tString
         Else
            ObjPrint.Print tString
         End If
         
         CurrentPos = CharLength + CurrentPos
         StartChar = CurrentPos + 1
         CharLength = TxtLen
         
      End If
   Next x
   
   tString = Trim$(Mid$(PrintVar, StartChar))
   If NeedsStrip Then tString = GetRemoveCRLF(tString)
   Me.CurrentX = LeftMargin + ((COLWIDTH - GetTextWidth(tString)) / 2)
   If PrintFlag Then
      If SameLine Then
         Printer.Print tString;
      Else
         Printer.Print tString
      End If
   Else
      If SameLine Then
         ObjPrint.Print tString;
      Else
         ObjPrint.Print tString
      End If
   End If

Exit_Here:
Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCenterMultiline"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pCircle(ByVal bLeft As Single, _
                   ByVal bTop As Single, _
                   ByVal bRadius As Single, _
                   Optional ByVal ColorLine As Long = -1, _
                   Optional ByVal ColorFill As Long = -1, _
                   Optional FilledCircle As FillStyleConstants = vbFSTransparent, _
                   Optional AspectRatio As Single = 1)
   
   On Error GoTo Err_Proc
   
   If ColorLine = -1 Then ColorLine = Me.ForeColor
   If ColorFill = -1 Then ColorFill = ColorLine
   
   If PrintFlag Then
      If FilledCircle <> vbFSTransparent Then
         Printer.FillStyle = FilledCircle
         Printer.FillColor = ColorFill
      End If
      
      Printer.Circle (bLeft, bTop), bRadius, ColorLine, , , AspectRatio
      
      Printer.FillStyle = vbFSTransparent
   Else
      If FilledCircle <> vbFSTransparent Then
         ObjPrint.FillStyle = FilledCircle
         ObjPrint.FillColor = ColorFill
      End If
      
      ObjPrint.Circle (bLeft, bTop), bRadius, ColorLine, , , AspectRatio
      
      ObjPrint.FillStyle = vbFSTransparent
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pCircle"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pDots(ByVal RightMargin As Single, _
                 Optional ByVal LeftMargin As Single = 0)
                 
   If LeftMargin > 0 Then Me.CurrentX = LeftMargin
   If Me.CurrentX >= RightMargin Then Exit Sub
   Do
      pPrint ".", , True
   Loop Until Me.CurrentX >= RightMargin
   
End Sub

Public Sub pDoubleLine(Optional ByVal LeftPos As Single = 0, _
                       Optional ByVal RightPos As Single = -1, _
                       Optional ByVal LineWidth As Integer = 1, _
                       Optional IncludeSpace As Boolean = True)
   
  Dim eFont As Integer
  Dim eDrawWidth As Integer
  Dim cY As Single
  Dim cX As Single
   
   On Error GoTo Err_Proc
   
   cX = Me.CurrentX
   Select Case oScaleMode
   Case vbCentimeters
      cY = 0.07
   Case Else '// vbinches
      cY = 0.03
   End Select
   
   eDrawWidth = Me.DrawWidth
   If LineWidth > 0 Then Me.DrawWidth = LineWidth
   'If IncludeSpace Then me.CurrentY = me.CurrentY + cY
   
   If PrintFlag Then
      cX = Printer.CurrentX
      If RightPos <= LeftPos Then RightPos = PgWidth
      Printer.CurrentY = Printer.CurrentY + cY
      Printer.Line (LeftPos, Printer.CurrentY)-(RightPos, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + cY
      Printer.Line (LeftPos, Printer.CurrentY)-(RightPos, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + cY
      Printer.CurrentX = cX
   Else
      cX = ObjPrint.CurrentX
      If RightPos <= LeftPos Then RightPos = PgWidth
      ObjPrint.CurrentY = ObjPrint.CurrentY + cY
      ObjPrint.Line (LeftPos, ObjPrint.CurrentY)-(RightPos, ObjPrint.CurrentY)
      ObjPrint.CurrentY = ObjPrint.CurrentY + cY
      ObjPrint.Line (LeftPos, ObjPrint.CurrentY)-(RightPos, ObjPrint.CurrentY)
      ObjPrint.CurrentY = ObjPrint.CurrentY + cY
      ObjPrint.CurrentX = cX
   End If
   
   If LineWidth > 0 Then Me.DrawWidth = eDrawWidth
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + cY
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pDoubleLine"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pEndDoc(Optional ByVal oModal As Byte = 1, Optional OwnerForm As Form)
  
  Dim i As Integer
   
   If PrintFlag Then
      Printer.EndDoc
      Printer.ScaleMode = pSM
      Me.SendToPrinter = False
   Else
      On Local Error Resume Next
      ObjPrint.ScaleMode = oSM
      
      SavePicture ObjPrint.Image, TempDir & "PPview" & CStr(PageNumber) & ".bmp"
      
      frmMultiPgPreview.PageNumber = PageNumber
      frmMultiPgPreview.Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(0) & ".bmp")
      frmMultiPgPreview.Show oModal, OwnerForm
   End If
End Sub

Public Function pEndOfPage(Optional ByVal Less As Single = 0, _
                           Optional ByVal SaveRoomForFooter As Boolean = True) As Boolean
   
  Dim n As Single
  Dim fTextHeight As Single
  Dim eFontSize As Integer
   
   On Error GoTo Err_Proc
   
   If PrintFlag Then
      '// Make sure there is room for the footer
      fTextHeight = Printer.TextHeight("TextString")
      If SaveRoomForFooter Then
         eFontSize = Printer.FontSize
         Printer.FontSize = 10
         fTextHeight = Printer.TextHeight("TextString") * 2
         Printer.FontSize = eFontSize
      End If
      n = Printer.CurrentY + Printer.TextHeight("TextString") + fTextHeight + Less
   Else
      '// Make sure there is room for the footer
      fTextHeight = ObjPrint.TextHeight("TextString")
      If SaveRoomForFooter Then
         eFontSize = ObjPrint.FontSize
         ObjPrint.FontSize = 10
         fTextHeight = ObjPrint.TextHeight("TextString") * 2
         ObjPrint.FontSize = eFontSize
      End If
      n = ObjPrint.CurrentY + ObjPrint.TextHeight("TextString") + fTextHeight + Less
   End If
   
   If n >= PgHeight Then
      pEndOfPage = True
   Else
      pEndOfPage = False
   End If
   
Exit_Here:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pEndOfPage"
   Err.Clear
   Resume Exit_Here
   
End Function

Public Sub pFontName(Optional ByVal pFontName As String = "Times New Roman")
   
   On Error GoTo Err_Proc
   
   If PrintFlag Then
      Printer.FontName = pFontName
      Printer.Print vbNullString;
   Else
      ObjPrint.FontName = pFontName
      ObjPrint.Print vbNullString;
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pFontName"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pFooter(Optional FooterText As String = vbNullString)
   
  Dim eFontS As Integer
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean
  Dim tMargin As Single
   
   On Error GoTo Err_Proc
   
   Select Case oScaleMode
   Case vbCentimeters
      tMargin = 1.25
   Case Else '// vbinches
      tMargin = 0.5
   End Select
   
   '// Save current setting
   eFontN = Me.FontName
   eFontS = Me.FontSize
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru
   
   '// Change settings
   pFontName
   Me.FontSize = 10
   Me.FontBold = False
   Me.FontItalic = False
   Me.FontUnderline = False
   Me.FontStrikethru = False
   Me.CurrentY = PgHeight - (GetTextHeight * 2)
   pLine 0, , 6
   
   If FooterText = vbNullString Then FooterText = "Printed: " & Format$(Now, "ddd. mmmm d, yyyy \@ h:mm ampm")
   
   pPrint FooterText, tMargin, True
   pRightJust "Pg. " & GetPage, PgWidth - tMargin
   
   '// Restore setting
   Me.FontName = eFontN
   Me.FontSize = eFontS
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pFooter"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pHalfSpace()
   
  Dim eFont As Integer
  Dim hFont As Integer
   
   On Error GoTo Err_Proc
   
   eFont = Me.FontSize
   hFont = eFont \ 2
   If hFont < 1 Then hFont = 1
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pHalfSpace"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pHeader(Optional ByVal MainTitle As String = vbNullString, _
                   Optional SubTitle As Variant, _
                   Optional ItalicMain As Variant)
   
  Dim eFontS As Integer
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean
   
   On Error GoTo Err_Proc
   
   If MainTitle = vbNullString And oTitleMain = vbNullString Then Exit Sub
   
   '// Save current setting
   eFontN = Me.FontName
   eFontS = Me.FontSize
   eFontB = Me.FontBold
   eFontI = Me.FontItalic
   eFontU = Me.FontUnderline
   eFontK = Me.FontStrikethru
   Me.CurrentY = 0
   Me.CurrentX = 0
   
   If MainTitle > vbNullString Then oTitleMain = MainTitle
   If Not IsMissing(SubTitle) Then oTitleSub = SubTitle
   If Not IsMissing(ItalicMain) Then oTitleItalic = CBool(ItalicMain)
   
   '// Change settings
   pFontName
   Me.FontSize = 16
   Me.FontBold = True
   Me.FontItalic = oTitleItalic
   pCenter oTitleMain
   
   Me.FontSize = 11
   Me.FontBold = False
   Me.FontItalic = False
   If oTitleSub > vbNullString Then pCenter oTitleSub
   pDoubleLine
   pHalfSpace
   
   '// Restore setting
   Me.FontName = eFontN
   Me.FontSize = eFontS
   Me.FontBold = eFontB
   Me.FontItalic = eFontI
   Me.FontUnderline = eFontU
   Me.FontStrikethru = eFontK
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pHeader"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pLine(Optional ByVal LeftMargin As Single = -1, _
                 Optional ByVal RightMargin As Single = -1, _
                 Optional ByVal LineWidth As Integer = 1, _
                 Optional ByVal IncludeSpace As Boolean = True)
   
  Dim eDrawWidth As Integer
  Dim cY As Single
  Dim cX As Single
   
   On Error GoTo Err_Proc
   
   Select Case oScaleMode
   Case vbCentimeters
      cY = 0.07
   Case Else '// vbinches
      cY = 0.03
   End Select
   
   If LeftMargin = -1 Then LeftMargin = Me.CurrentX
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + cY
   
   If LineWidth > 0 Then
      eDrawWidth = Me.DrawWidth
      Me.DrawWidth = LineWidth
   End If
   
   If PrintFlag Then
      cX = Printer.CurrentX
      If RightMargin <= LeftMargin Then RightMargin = PgWidth
      Printer.Line (LeftMargin, Printer.CurrentY)-(RightMargin, Printer.CurrentY)
      Printer.CurrentX = cX
   Else
      cX = ObjPrint.CurrentX
      If RightMargin <= LeftMargin Then RightMargin = PgWidth
      ObjPrint.Line (LeftMargin, ObjPrint.CurrentY)-(RightMargin, ObjPrint.CurrentY)
      ObjPrint.CurrentX = cX
   End If
   
   If LineWidth > 0 Then Me.DrawWidth = eDrawWidth
   If IncludeSpace Then Me.CurrentY = Me.CurrentY + cY
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pLine"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pMultiline(ByVal PrintVar As Variant, _
                      Optional ByVal LeftMargin As Single = -1, _
                      Optional ByVal RightMargin As Single = -1, _
                      Optional ByVal IndentChar As String = vbNullString, _
                      Optional ByVal SameLine As Boolean = False, _
                      Optional ByVal UsePageBreaks As Boolean = True)
   
  Dim StartChar As Integer
  Dim SecondLine As Boolean
  Dim CharLength As Single
  Dim CurrentPos As Single
  Dim TxtLen As Single
  Dim TxtWidth As Single
  Dim IndentText As String
  Dim tString As String
  Dim NeedsStrip As Boolean
  Dim x As Long
  Dim y As Long
   
   On Error GoTo Err_Proc
   
   If LeftMargin = -1 Then LeftMargin = Me.CurrentX
   If LeftMargin > PgWidth - 0.1 Then LeftMargin = PgWidth - 0.5
   If RightMargin < LeftMargin Then RightMargin = PgWidth - 0.1
   RightMargin = RightMargin - LeftMargin
   
   If PrintVar = vbNullString Then
      If SameLine Then
         pPrint vbNullString, , True
      Else
         pPrint
      End If
      Exit Sub
   End If
   
   TxtLen = Len(PrintVar)
   StartChar = 1
   CurrentPos = 0
   CharLength = TxtLen
   IndentText = vbNullString
   
   If InStr(PrintVar, vbCr) Or InStr(PrintVar, vbLf) Then NeedsStrip = True
   
   For x = 1 To TxtLen
      y = x - CurrentPos
      If Mid$(PrintVar, x, 1) < Chr(33) Then CharLength = y
      If (GetTextWidth(IndentText) + GetTextWidth(Mid$(PrintVar, StartChar, y)) >= RightMargin) Or (Mid$(PrintVar, x, 1) = vbCr) Then
      
         '// If there are no spaces then break line here */
         If CharLength > y Then CharLength = y - Len(IndentText) - 1
         If NeedsStrip Then
            tString = IndentText & Trim$(GetRemoveCRLF(Mid$(PrintVar, StartChar, CharLength)))
         Else
            tString = IndentText & Mid$(PrintVar, StartChar, CharLength)
         End If
         
         Me.CurrentX = LeftMargin
         If PrintFlag Then
            Printer.Print tString
         Else
            ObjPrint.Print tString
         End If
         
         CurrentPos = CharLength + CurrentPos
         StartChar = CurrentPos + 1
         CharLength = TxtLen
         If Not SecondLine Then
            SecondLine = True
            IndentText = IndentChar
         End If
         
         If UsePageBreaks Then
            If pEndOfPage Then
               pFooter
               pNewPage
               pHeader
            End If
         End If
      End If
   Next x

   If NeedsStrip Then
      tString = IndentText & Trim$(GetRemoveCRLF(Mid$(PrintVar, StartChar)))
   Else
      tString = IndentText & Mid$(PrintVar, StartChar)
   End If
   
   Me.CurrentX = LeftMargin
   If SameLine Then
      If PrintFlag Then
         Printer.Print tString;
      Else
         ObjPrint.Print tString;
      End If
   Else
      If PrintFlag Then
         Printer.Print tString
      Else
         ObjPrint.Print tString
      End If
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pMultiline"
   Err.Clear
   Resume Exit_Here

End Sub

Public Sub pNewPage()
   
   On Local Error Resume Next
   If PrintFlag Then
      Printer.NewPage
   Else
      SavePicture ObjPrint.Image, TempDir & "PPview" & CStr(PageNumber) & ".bmp"
      ObjPrint.Cls
      PageNumber = PageNumber + 1
   End If
End Sub

Public Sub pPrint(Optional ByVal PrintVar As String = vbNullString, _
                  Optional ByVal LeftMargin As Single = -1, _
                  Optional ByVal SameLine As Boolean = False)
   
   On Error GoTo Err_Proc
   
   If PrintVar = vbNullString Then
      '// Empty String */
   Else
      If LeftMargin = -1 Then LeftMargin = Me.CurrentX
      If (GetTextWidth(PrintVar) + LeftMargin > PgWidth) Or InStr(PrintVar, vbCr) Or InStr(PrintVar, vbLf) Then
         
         pMultiline PrintVar, LeftMargin, PgWidth - 0.1, , SameLine
         Exit Sub
      End If
   End If
   
   If LeftMargin >= 0 Then Me.CurrentX = LeftMargin
   
   If SameLine Then
      If PrintFlag Then
         Printer.Print PrintVar;
      Else
         ObjPrint.Print PrintVar;
      End If
   Else
      If PrintFlag Then
         Printer.Print PrintVar
      Else
         ObjPrint.Print PrintVar
      End If
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrint"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pPrintedDate(Optional PrintCentered As Boolean = False, _
                        Optional ByVal LeftMargin As Single = -1, _
                        Optional ByVal SameLine As Boolean = False)
   
  Dim PrintVar As String
  Dim FSize As Integer
  Dim FBold As Boolean
  Dim FItalic As Boolean
  Dim FUnderline As Boolean
  Dim FStrikethru As Boolean
   
   On Error GoTo Err_Proc
   
   FSize = Me.FontSize
   FBold = Me.FontBold
   FItalic = Me.FontItalic
   FUnderline = Me.FontUnderline
   FStrikethru = Me.FontStrikethru
   
   Me.FontSize = 9
   Me.FontBold = False
   Me.FontUnderline = False
   Me.FontItalic = False
   Me.FontStrikethru = False
   PrintVar = "Printed: " & Format$(Now, "ddd. mmm. d, yyyy \@ h:mm ampm")
   
   If PrintCentered Then
      pCenter PrintVar, SameLine
   Else
      pPrint PrintVar, LeftMargin, SameLine
   End If
   
   Me.FontSize = FSize
   Me.FontBold = FBold
   Me.FontItalic = FItalic
   Me.FontUnderline = FUnderline
   Me.FontStrikethru = FStrikethru
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrintedDate"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pPrintPicture(ByRef newPic As StdPicture, _
                         Optional LeftMargin As Single = -1, _
                         Optional TopMargin As Single = -1, _
                         Optional pWidth As Single = 0, _
                         Optional pHeight As Single = 0, _
                         Optional ByVal ScaleToFit As Boolean = False, _
                         Optional ByVal MaintainRatio As Boolean = True)
   
  Dim xmin As Single
  Dim ymin As Single
  Dim wid As Single
  Dim hgt As Single
  Dim aspect As Single
  Dim picBox As PictureBox
   
   On Error GoTo Err_Proc
   
   If pWidth = 0 Then pWidth = pHeight
   If pHeight = 0 Then pHeight = pWidth
   If pWidth = 0 And pHeight = 0 Then ScaleToFit = True
   
   If PrintFlag Then
      Load frmMultiPgPreview
      Set picBox = frmMultiPgPreview.picPrintPic
      picBox.Picture = newPic
      
      aspect = picBox.ScaleHeight / picBox.ScaleWidth
      If ScaleToFit Then
         wid = Printer.ScaleWidth
         hgt = Printer.ScaleHeight
      Else
         wid = pWidth
         hgt = pHeight
      End If
      
      If MaintainRatio Then
         If hgt / wid > aspect Then
            hgt = aspect * wid
            If LeftMargin = -1 Then
               xmin = Printer.ScaleLeft
            Else
               xmin = LeftMargin
            End If
            If TopMargin = -1 Then
               ymin = (Printer.ScaleHeight - hgt) / 2
            Else
               ymin = TopMargin
            End If
         Else
            wid = hgt / aspect
            If LeftMargin = -1 Then
               xmin = (Printer.ScaleWidth - wid) / 2
            Else
               xmin = LeftMargin
            End If
            If TopMargin = -1 Then
               ymin = Printer.ScaleTop
            Else
               ymin = TopMargin
            End If
         End If
      Else
         If LeftMargin = -1 Then
            xmin = Printer.ScaleLeft
         Else
            xmin = LeftMargin
         End If
         If TopMargin = -1 Then
            ymin = (Printer.ScaleHeight - hgt) / 2
         Else
            ymin = TopMargin
         End If
      End If
      
      Printer.PaintPicture picBox.Picture, xmin, ymin, wid, hgt, , , , , vbSrcCopy
      Unload frmMultiPgPreview
      
   Else
      
      Set picBox = frmMultiPgPreview.picPrintPic
      picBox.Picture = newPic
      
      aspect = picBox.ScaleHeight / picBox.ScaleWidth
      If ScaleToFit Then
         wid = ObjPrint.ScaleWidth
         hgt = ObjPrint.ScaleHeight
      Else
         wid = pWidth
         hgt = pHeight
      End If
      
      If MaintainRatio Then
         If hgt / wid > aspect Then
            hgt = aspect * wid
            If LeftMargin = -1 Then
               xmin = ObjPrint.ScaleLeft
            Else
               xmin = LeftMargin
            End If
            If TopMargin = -1 Then
               ymin = (ObjPrint.ScaleHeight - hgt) / 2
            Else
               ymin = TopMargin
            End If
         Else
            wid = hgt / aspect
            If LeftMargin = -1 Then
               xmin = (ObjPrint.ScaleWidth - wid) / 2
            Else
               xmin = LeftMargin
            End If
            If TopMargin = -1 Then
               ymin = ObjPrint.ScaleTop
            Else
               ymin = TopMargin
            End If
         End If
      Else
         If LeftMargin = -1 Then
            xmin = ObjPrint.ScaleLeft
         Else
            xmin = LeftMargin
         End If
         If TopMargin = -1 Then
            ymin = (ObjPrint.ScaleHeight - hgt) / 2
         Else
            ymin = TopMargin
         End If
      End If
      
      ObjPrint.PaintPicture picBox.Picture, xmin, ymin, wid, hgt, , , , , vbSrcCopy
      
      picBox.Picture = Nothing
      
   End If
   
   Set picBox = Nothing
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pPrintPicture"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pPrintRotate(Optional ByVal PrintVar As String = vbNullString, _
                        Optional ByVal Degree As Integer = 0, _
                        Optional ByVal LeftMargin As Single = -1)
   
   '// By Diomidis Kiriakopoulos modified by Me
   
   If PrintVar = vbNullString Then Exit Sub
   If Degree > 359 Then
      Degree = 359
   ElseIf Degree < 0 Then
      Degree = 0
   End If
   If LeftMargin = -1 Then
      LeftMargin = Me.CurrentX
   Else
      Me.CurrentX = LeftMargin
   End If
   
   On Local Error GoTo GetOut
   
  Dim F As LOGFONT
  Dim hPrevFont As Long
  Dim hFont As Long
  Dim ObjhDC As Long
  Dim x As Long
  Dim y As Long
   
   '// Save the hDC.
   If PrintFlag Then
      Printer.Print vbNullString;
      ObjhDC = Printer.hdc
      F.lfHeight = (FontSize * -20) / Printer.TwipsPerPixelY
      x = Printer.CurrentX * 600 '// Printer dots/inch
      y = Printer.CurrentY * 600
   Else
      ObjhDC = ObjPrint.hdc
      F.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
   End If
   
   F.lfEscapement = 10 * Degree         '// rotation angle, in tenths
   F.lfFaceName = Me.FontName & vbNullChar '// null terminated
   hFont = CreateFontIndirect(F)
   
   hPrevFont = SelectObject(ObjhDC, hFont)
   
   If PrintFlag Then
      '// Draw the text.
      TextOut ObjhDC, x, y, PrintVar, Len(PrintVar)
   Else
      ObjPrint.Print PrintVar;
   End If
   
   '//  Clean up, restore original font
   hFont = SelectObject(ObjhDC, hPrevFont)
   DeleteObject hFont
   
GetOut:
   On Local Error GoTo 0
   
End Sub

Public Sub pQuarterSpace()
   
  Dim eFont As Integer
  Dim hFont As Integer
   
   On Error GoTo Err_Proc
   
   eFont = Me.FontSize
   hFont = eFont \ 4
   If hFont < 1 Then hFont = 1
   Me.FontSize = hFont
   pPrint
   Me.FontSize = eFont
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pQuarterSpace"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pRightJust(ByVal PrintVar As Variant, _
                      Optional ByVal RightMargin As Single = -1, _
                      Optional ByVal SameLine As Boolean = False)
   
  Dim TxtWidth As Single
   
   On Error GoTo Err_Proc
   
   TxtWidth = GetTextWidth(PrintVar)
   If RightMargin = -1 Then RightMargin = Me.CurrentX + TxtWidth
   Me.CurrentX = RightMargin - TxtWidth
   
   If SameLine Then
      If PrintFlag Then
         Printer.Print PrintVar;
      Else
         ObjPrint.Print PrintVar;
      End If
   Else
      If PrintFlag Then
         Printer.Print PrintVar
      Else
         ObjPrint.Print PrintVar
      End If
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pRightJust"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pRightTab(ByVal PrintVar As Variant, _
                     Optional ByVal xFromRight As Single = 0.1, _
                     Optional ByVal SameLine As Boolean = False)
   
   On Error GoTo Err_Proc
   
   Me.CurrentX = PgWidth - (GetTextWidth(PrintVar) + xFromRight)
   
   If SameLine Then
      If PrintFlag Then
         Printer.Print PrintVar;
      Else
         ObjPrint.Print PrintVar;
      End If
   Else
      If PrintFlag Then
         Printer.Print PrintVar
      Else
         ObjPrint.Print PrintVar
      End If
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pRightTab"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Property Get PrintCopies() As Integer
   On Local Error Resume Next
   PrintCopies = Printer.Copies
   On Local Error GoTo 0
End Property

Public Property Let PrintCopies(pNumber As Integer)
   On Local Error Resume Next
   Printer.Copies = pNumber
   On Local Error GoTo 0
End Property

Public Sub pSpaces(Optional ByVal RightMargin As Single = -1, _
                   Optional ByVal LeftMargin As Single = -1, _
                   Optional ByVal UseSymbol As Boolean = False)
   
  Dim xFontname As String
  Dim xForeColor As Long
  Dim tString As String
   
   On Error GoTo Err_Proc
   
   xFontname = Me.FontName
   xForeColor = Me.ForeColor
   
   If LeftMargin <> -1 Then Me.CurrentX = LeftMargin
   If RightMargin = -1 Then RightMargin = PgWidth
   
   If UseSymbol Then
      Me.FontName = "Symbol"
      tString = "\"
      Me.ForeColor = vbGreen
   Else
      tString = " "
   End If
   
   If Me.CurrentX >= RightMargin Then GoTo ExitSpaceSub
   Do
      pPrint tString, , True
   Loop Until Me.CurrentX >= RightMargin
   
ExitSpaceSub:
   Me.FontName = xFontname
   Me.ForeColor = xForeColor
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pSpaces"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Sub pStartDoc()
   
   PageNumber = 0
   TempDir = Environ("TEMP") & "\"
   
   On Local Error Resume Next
   
   '// Set the Printer's scale mode
   pSM = Printer.ScaleMode
   Printer.ScaleMode = oScaleMode
   
   '// Get the physical printable area
   PgWidth = Printer.ScaleWidth
   PgHeight = Printer.ScaleHeight
   
   If PrintFlag Then
      '// Initialize printer
      Printer.Print vbNullString;
   Else
      '// Initialize the preview object
      Load frmMultiPgPreview
      Set ObjPrint = frmMultiPgPreview.Picture1
      
      '// Scale Object to Printer's printable area
      oSM = ObjPrint.ScaleMode
      ObjPrint.ScaleMode = oScaleMode
      
      '// Full Page size (1440 twips = 1 inch or 567 twips = 1 centimeter)
      Select Case oScaleMode
      Case vbCentimeters
         ObjPrint.Width = (PgWidth + 0.6) * 567
         ObjPrint.Height = (PgHeight + 0.6) * 567
      Case Else '// vbinches
         ObjPrint.Width = (PgWidth + 0.25) * 1440
         ObjPrint.Height = (PgHeight + 0.25) * 1440
      End Select
      
      '// Set default properties of the scroll bars
      frmMultiPgPreview.VScroll1.Max = Val(ObjPrint.Height * 0.5)
      frmMultiPgPreview.VScroll1.Min = -500
      frmMultiPgPreview.VScroll1.SmallChange = Val(frmMultiPgPreview.VScroll1.Max * 0.1)
      frmMultiPgPreview.VScroll1.LargeChange = Val(frmMultiPgPreview.VScroll1.Max * 0.5)
      
      frmMultiPgPreview.HScroll1.Max = Val(ObjPrint.Width * 0.25)
      frmMultiPgPreview.HScroll1.Min = -500
      frmMultiPgPreview.HScroll1.SmallChange = Val(frmMultiPgPreview.HScroll1.Max * 0.1)
      frmMultiPgPreview.HScroll1.LargeChange = Val(frmMultiPgPreview.HScroll1.Max * 0.5)
      
      '// Set default properties of the object to match printer
      ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
      ObjPrint.FontName = Printer.FontName
      ObjPrint.FontSize = Printer.FontSize
      ObjPrint.ForeColor = Printer.ForeColor
      ObjPrint.Picture = Nothing
      
   End If
   
End Sub

Public Sub pVerticalLine(Optional ByVal LeftPos As Single = -1, _
                         Optional ByVal TopPos As Single = -1, _
                         Optional ByVal BottomPos As Single = -1, _
                         Optional ByVal LineWidth As Integer = 0)
   
  Dim eDrawWidth As Integer
  Dim cY As Single
  Dim cX As Single
  Dim tH As Single
  Dim eFontSize As Integer
   
   On Error GoTo Err_Proc
   
   eDrawWidth = Me.DrawWidth
   cX = Me.CurrentX
   cY = Me.CurrentY
   If LineWidth > 0 Then Me.DrawWidth = LineWidth
   
   If BottomPos = -1 Then
      eFontSize = Me.FontSize
      FontSize = 10
      BottomPos = PgHeight - (GetTextHeight * 2)
      Me.FontSize = eFontSize
   End If
   
   If LeftPos = -1 Then LeftPos = Me.CurrentX
   If LeftPos > PgWidth Then LeftPos = PgWidth - 0.01
   If TopPos = -1 Then TopPos = Me.CurrentY
   
   If PrintFlag Then
      Printer.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
   Else
      ObjPrint.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
   End If
   
   Me.CurrentX = cX
   Me.CurrentY = cY
   
   If LineWidth > 0 Then Me.DrawWidth = eDrawWidth
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsMultiPgPreview", "pVerticalLine"
   Err.Clear
   Resume Exit_Here
   
End Sub

Public Property Get ReportSubTitle() As String
   ReportSubTitle = oTitleSub
End Property

Public Property Let ReportSubTitle(ByVal vNewValue As String)
   oTitleSub = vNewValue
End Property

Public Property Get ReportTitle() As String
   ReportTitle = oTitleMain
End Property

Public Property Let ReportTitle(ByVal vNewValue As String)
   oTitleMain = vNewValue
End Property

Public Property Get SendToPrinter() As Boolean
   SendToPrinter = PrintFlag
End Property

Public Property Let SendToPrinter(ByVal vNewValue As Boolean)
   PrintFlag = vNewValue
End Property
