VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'// Copyright (c) 1997-2004
'//************************************/
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

'// Windows API/Global Declarations for : Determine the Available Screen Real Estate
Private Type Rect
   left      As Long
   top       As Long
   right     As Long
   bottom    As Long
End Type

Private Const SPI_GETWORKAREA As Long = 48&
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
      (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
      ByVal X As Long, ByVal Y As Long, ByVal cX As Long, _
      ByVal cY As Long, ByVal wFlags As Long) As Long

'// Windows API/Global Declarations for IsScreenFontSmall
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" _
      (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
      (ByVal hWnd As Long, ByVal hdc As Long) As Long

'// Windows API/Global Declarations for Resizing
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Private Const C_lngSSTAB_DIALOG_OFFSET            As Long = 75000
Private Const C_strSSTAB_TYPE_NAME                As String = "SSTAB"
Private Const C_strResizeRepositionCommand        As String = "@"
Private Const C_strResizeRepositionLeft           As String = "lngLeft"
Private Const C_strResizeRepositionTop            As String = "lngTop"
Private Const C_strResizeRepositionWidth          As String = "W"
Private Const C_strResizeRepositionHeight         As String = "H"

Private Type ControlPosition
   ControlInstance          As Control      'Reference to the control instance.
   OriginalLeft             As Long         'Original Left position of the Control.
   OriginalTop              As Long         'Original Top position of the Control.
   OriginalWidth            As Long         'Original Width of the Control.
   OriginalHeight           As Long         'Original Height of the Control.
End Type

Private mfrm_SourceForm      As Form       'The form to be resized.
Private mlng_FormWidth       As Long       'Original form width.
Private mlng_FormHeight      As Long       'Original form height.
Private mudt_Controls()      As ControlPosition 'Array for storing control information.
Private mbln_IsFirstResize   As Boolean    'Flag indicating first resize.

Public Enum enuOnTopPar
   ON_TOP_YES = False
   ON_TOP_NO = True
End Enum

Public Sub CenterForm(ByRef rfrmForm As Form, _
                      Optional ByVal vsngMoveUpDn As Single = 1)
   
  Dim Rc        As Rect
  Dim lngTop    As Long
  Dim lngBottom As Long
  Dim lngLeft   As Long
  Dim lngRight  As Long
  Dim lngTopT   As Long
  Dim lngLeftT  As Long
   
   On Error Resume Next
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   
   lngTop = Rc.top * Screen.TwipsPerPixelY
   lngBottom = Rc.bottom * Screen.TwipsPerPixelY
   lngLeft = Rc.left * Screen.TwipsPerPixelX
   lngRight = Rc.right * Screen.TwipsPerPixelX
   
   lngTopT = Abs((lngBottom / 2 * vsngMoveUpDn) - (rfrmForm.Height / 2))
   lngLeftT = Abs((lngRight / 2) - (rfrmForm.Width / 2))
   
   If lngTopT < lngTop Then lngTopT = lngTop
   If lngTopT > lngBottom - rfrmForm.Height Then lngTopT = lngBottom - rfrmForm.Height
   If lngLeftT < lngLeft Then lngLeftT = lngLeft
   
   rfrmForm.Move lngLeftT, lngTopT
   
End Sub

Public Sub CenterObject(ByRef ParentObject As Object, _
                        ByRef ChildObject As Object)
   
   ChildObject.top = (ParentObject.Height - ChildObject.Height) / 2
   ChildObject.left = (ParentObject.Width - ChildObject.Width) / 2
   
End Sub

Private Sub Class_Initialize()
   
   Set mfrm_SourceForm = Nothing
   mbln_IsFirstResize = True
   
End Sub

Public Sub FitScreen(ByRef rfrmForm As Form, _
                     Optional ByVal vsngPercent As Single = 100)
   
  Dim Rc As Rect
   
   On Error Resume Next
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   'left - " & rc.Left
   'top - " & rc.Top
   'rght - " & rc.Right
   'bttm - " & rc.Bottom
   
   If vsngPercent > 100 Then vsngPercent = 100
   vsngPercent = vsngPercent / 100
   
   rfrmForm.Move Rc.left * Screen.TwipsPerPixelX * vsngPercent, _
         Rc.top * Screen.TwipsPerPixelY * vsngPercent, _
         Rc.right * Screen.TwipsPerPixelX * vsngPercent, _
         Rc.bottom * Screen.TwipsPerPixelY * vsngPercent
   
   If vsngPercent < 1 Then Call CenterForm(rfrmForm)
   
   On Local Error GoTo 0
   
End Sub

Public Sub HCenterObject(ByRef rObj As Object)
   
  Dim Rc       As Rect
  Dim lngLeft  As Long
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   
   lngLeft = Rc.right * Screen.TwipsPerPixelX
   lngLeft = lngLeft / 2 - rObj.Width / 2
   rObj.left = lngLeft
   
End Sub

Public Property Get Height() As Variant
   
   Height = GetSystemMetrics(1&)
   
End Property

Public Sub hFitScreen(ByRef rfrmForm As Form, _
                      Optional vblnCenter As Boolean = True)
   
  Dim Rc As Rect
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   rfrmForm.left = Rc.left * Screen.TwipsPerPixelX
   rfrmForm.Width = (Rc.right - Rc.left) * Screen.TwipsPerPixelX
   If vblnCenter Then CenterForm rfrmForm
   
End Sub

Public Sub OnTop(ByRef rfrmForm As Form, _
                 Optional vblnTurnOff As enuOnTopPar = ON_TOP_YES)
   
  Const HWND_TOPMOST    As Long = -1
  Const HWND_NOTTOPMOST As Long = -2
   
   If vblnTurnOff Then
      Call SetWindowPos(rfrmForm.hWnd, HWND_NOTTOPMOST, 0, 0, 0, 0, 3)
    Else
      Call SetWindowPos(rfrmForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
   End If
   
End Sub

Public Sub ResizeControls(ByRef rfrmForm As Form)
   
  Dim dblHeightChange          As Double
  Dim dblWidthChange           As Double
  Dim ctlControlInstance       As Control
  Dim strResizeRepositionFlags As String
  Dim lngI                     As Long
   
   Set mfrm_SourceForm = rfrmForm
   
   '// Check if the form has any controls on it.
   If mfrm_SourceForm.Controls.Count = 0 Then Exit Sub
   
   '// Check if the form is being minimized.
   If mfrm_SourceForm.WindowState = vbMinimized Then Exit Sub
   
   '// Check if it's the first time the form is to be resized.
   If mbln_IsFirstResize Then
      Call ResizeSaveInitialStates
      Exit Sub
   End If
   
   '// Prevent the Window from redrawing until the user releases the mouse.
   Call LockWindowUpdate(mfrm_SourceForm.hWnd)
   
   '// Calculate the change in the form's size.
   If (TypeOf mfrm_SourceForm Is MDIForm) Then
      dblHeightChange = mfrm_SourceForm.Height / mlng_FormHeight
      dblWidthChange = mfrm_SourceForm.Width / mlng_FormWidth
    Else
      dblHeightChange = mfrm_SourceForm.ScaleHeight / mlng_FormHeight
      dblWidthChange = mfrm_SourceForm.ScaleWidth / mlng_FormWidth
   End If
   
   '// Iterate through the array containing the form's controls.
   For lngI = 0 To UBound(mudt_Controls)
      Set ctlControlInstance = mudt_Controls(lngI).ControlInstance
      
      '// Use the Control that is currently being referenced.
      With ctlControlInstance
         On Local Error Resume Next
         '// Check if the Tag Property contains valid resize or reposition instructions.
         If (left(UCase$(.Tag), 1) = C_strResizeRepositionCommand) Then
            'Retrieve any specific resize and reposition instructions.
            strResizeRepositionFlags = UCase$(.Tag)
          Else
            '// Default all resize and resposition instructions.
            strResizeRepositionFlags = C_strResizeRepositionCommand & C_strResizeRepositionLeft & C_strResizeRepositionTop & C_strResizeRepositionWidth & C_strResizeRepositionHeight
         End If
         
         '// Resize and Reposition the control.
         If (TypeOf ctlControlInstance Is Line) Then
            '// Line control.
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionLeft) <> 0) Then
               .X1 = (mudt_Controls(lngI).OriginalLeft * dblWidthChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionTop) <> 0) Then
               .Y1 = (mudt_Controls(lngI).OriginalTop * dblHeightChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionWidth) <> 0) Then
               .X2 = (mudt_Controls(lngI).OriginalWidth * dblWidthChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionHeight) <> 0) Then
               .Y2 = (mudt_Controls(lngI).OriginalHeight * dblHeightChange)
            End If
          ElseIf (UCase$(TypeName(ctlControlInstance.Container)) = C_strSSTAB_TYPE_NAME) Then
            '// Control is placed on a SSTab Dialog Control.
            If (.left < 0) Then
               If (mudt_Controls(lngI).OriginalLeft > 0) Then
                  mudt_Controls(lngI).OriginalLeft = mudt_Controls(lngI).OriginalLeft - C_lngSSTAB_DIALOG_OFFSET
               End If
               
               If (InStr(strResizeRepositionFlags, C_strResizeRepositionLeft) <> 0) Then .left = ((mudt_Controls(lngI).OriginalLeft + C_lngSSTAB_DIALOG_OFFSET) * dblWidthChange) - C_lngSSTAB_DIALOG_OFFSET
             Else
               If (mudt_Controls(lngI).OriginalLeft < 0) Then
                  mudt_Controls(lngI).OriginalLeft = mudt_Controls(lngI).OriginalLeft + C_lngSSTAB_DIALOG_OFFSET
               End If
               
               If (InStr(strResizeRepositionFlags, C_strResizeRepositionLeft) <> 0) Then .left = mudt_Controls(lngI).OriginalLeft * dblWidthChange
            End If
            
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionTop) <> 0) Then
               .top = (mudt_Controls(lngI).OriginalTop * dblHeightChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionWidth) <> 0) Then
               .Width = (mudt_Controls(lngI).OriginalWidth * dblWidthChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionHeight) <> 0) Then
               .Height = (mudt_Controls(lngI).OriginalHeight * dblHeightChange)
            End If
          Else
            '// All other controls.
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionLeft) <> 0) Then
               .left = (mudt_Controls(lngI).OriginalLeft * dblWidthChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionTop) <> 0) Then
               .top = (mudt_Controls(lngI).OriginalTop * dblHeightChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionWidth) <> 0) Then
               .Width = (mudt_Controls(lngI).OriginalWidth * dblWidthChange)
            End If
            If (InStr(strResizeRepositionFlags, C_strResizeRepositionHeight) <> 0) Then
               .Height = (mudt_Controls(lngI).OriginalHeight * dblHeightChange)
            End If
         End If
         On Local Error GoTo 0
      End With
   Next lngI
   
   '// Release the memory being used by the ctlControlInstance variable.
   Set ctlControlInstance = Nothing
   
   '// Allow Windows to refresh the form's display.
   Call LockWindowUpdate(0)
   
End Sub

Public Sub ResizeFonts(ByRef rfrmForm As Form)
   
  Dim lngI                  As Long
  Dim sngFontRatio          As Single
  Const C_intDESIGNFONTSIZE As Integer = 96 '// 96 = Small & 120 = Large
   
   '// check to see what size of fonts are being used
   sngFontRatio = C_intDESIGNFONTSIZE / ScreenFontSize()
   If sngFontRatio = 1 Then Exit Sub
   
   On Local Error Resume Next
   
   With rfrmForm
      For lngI = 0 To .Controls.Count - 1 '// Number of controls on the form
         
         If .Controls(lngI).Font.Size <= 8 And sngFontRatio < 1 Then
            .Controls(lngI).Font.Name = "Small Fonts"
            .Controls(lngI).Font.Size = .Controls(lngI).Font.Size - 0.5
          Else
            .Controls(lngI).Font.Size = .Controls(lngI).Font.Size * sngFontRatio
         End If
         
         If TypeOf .Controls(lngI) Is Line Then
            .Controls(lngI).X1 = .Controls(lngI).X1 * sngFontRatio
            .Controls(lngI).Y1 = .Controls(lngI).Y1 * sngFontRatio
            .Controls(lngI).X2 = .Controls(lngI).X2 * sngFontRatio
            .Controls(lngI).Y2 = .Controls(lngI).Y2 * sngFontRatio
          ElseIf TypeOf .Controls(lngI) Is PictureBox Then
            .Controls(lngI).Width = .Controls(lngI).Width * sngFontRatio
            .Controls(lngI).Height = .Controls(lngI).Height * sngFontRatio
            .Controls(lngI).top = .Controls(lngI).top * sngFontRatio
            .Controls(lngI).left = .Controls(lngI).left * sngFontRatio
            .Controls(lngI).ScaleHeight = .Controls(lngI).ScaleHeight * sngFontRatio
            .Controls(lngI).ScaleWidth = .Controls(lngI).ScaleWidth * sngFontRatio
          Else
            .Controls(lngI).Width = .Controls(lngI).Width * sngFontRatio
            .Controls(lngI).Height = .Controls(lngI).Height * sngFontRatio
            .Controls(lngI).top = .Controls(lngI).top * sngFontRatio
            .Controls(lngI).left = .Controls(lngI).left * sngFontRatio
            '// If control is a grid
            .Controls(lngI).COLWIDTH = .Controls(lngI).COLWIDTH * sngFontRatio
            .Controls(lngI).RowHeight = .Controls(lngI).RowHeight * sngFontRatio
            '// If Control is a Toolbar
            .Controls(lngI).ButtonHeight = .Controls(lngI).ButtonHeight * sngFontRatio
            .Controls(lngI).ButtonWidth = .Controls(lngI).ButtonWidth * sngFontRatio
         End If
      Next lngI
      
   End With
   
   On Local Error GoTo 0
   
End Sub

Public Sub ResizeSaveInitialSizes(ByRef rfrmForm As Form)
   
   Set mfrm_SourceForm = rfrmForm
   mbln_IsFirstResize = True
   Call ResizeSaveInitialStates
   
End Sub

Private Sub ResizeSaveInitialStates()
   
   '// Description : Store the dimensions and positions of all controls on the Form.            '
   
  Dim lngI As Long
   
   '// Check if the form has any controls on it.
   If mfrm_SourceForm.Controls.Count = 0 Then Exit Sub
   
   With mfrm_SourceForm
      '// Check if the form being resized is an MDI form.
      If (TypeOf mfrm_SourceForm Is MDIForm) Then
         '// Set the FormWidth and FormHeight variables to the Form's Width and Height.
         mlng_FormWidth = .Width
         mlng_FormHeight = .Height
       Else
         '// Set the FormWidth and FormHeight variables to the Form's Scale Width and Height.
         mlng_FormWidth = .ScaleWidth
         mlng_FormHeight = .ScaleHeight
      End If
   End With
   
   '// Frees the memory used by the dynamic array storing all control information.
   Erase mudt_Controls
   
   '// Redimension the array storing all control information to the number of controls on the
   '// form that is being resized.
   ReDim mudt_Controls(0 To (mfrm_SourceForm.Controls.Count - 1))
   
   '// Iterate through each of the Controls on the form to be resized.
   For lngI = 0 To (mfrm_SourceForm.Controls.Count - 1)
      '// Use the control that is currently being referenced.
      With mfrm_SourceForm.Controls(lngI)
         '// Add the controls reference to the array containing all control information.
         Set mudt_Controls(lngI).ControlInstance = mfrm_SourceForm.Controls(lngI)
         '// Check if the control currently being referenced is a line control.
         If (TypeOf mfrm_SourceForm.Controls(lngI) Is Line) Then
            '// Add a reference to the control and it's position into the Controls array.
            mudt_Controls(lngI).OriginalLeft = .X1
            mudt_Controls(lngI).OriginalTop = .Y1
            mudt_Controls(lngI).OriginalWidth = .X2
            mudt_Controls(lngI).OriginalHeight = .Y2
          Else
            On Local Error Resume Next
            '// Add a reference to the control and it's position into the Controls array.
            mudt_Controls(lngI).OriginalLeft = .left
            mudt_Controls(lngI).OriginalTop = .top
            mudt_Controls(lngI).OriginalWidth = .Width
            mudt_Controls(lngI).OriginalHeight = .Height
            On Local Error GoTo 0
         End If
      End With
   Next lngI
   
   '// Notify the class that the form has been resized.
   mbln_IsFirstResize = False
   
End Sub

Private Function ScreenFontSize() As Integer
   
  Dim lng_hWndDesk      As Long
  Dim lng_hDCDesk       As Long
  Dim lng_logPix        As Long
  Const C_lngLOGPIXELSX As Long = 88&
   'Const LOGPIXELSY As Long = 90
   
   lng_hWndDesk = GetDesktopWindow()
   lng_hDCDesk = GetWindowDC(lng_hWndDesk)
   lng_logPix = GetDeviceCaps(lng_hDCDesk, C_lngLOGPIXELSX)
   Call ReleaseDC(lng_hWndDesk, lng_hDCDesk)
   '// 96 = Small & 120 = Large
   ScreenFontSize = lng_logPix
   
End Function

Public Function scrnBottom() As Long
   
  Dim Rc As Rect
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   scrnBottom = Rc.bottom * Screen.TwipsPerPixelY
   
End Function

Public Function scrnLeft() As Long
   
  Dim Rc As Rect
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   scrnLeft = Rc.left * Screen.TwipsPerPixelX
   
End Function

Public Function scrnRight() As Long
   
  Dim Rc As Rect
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   scrnRight = Rc.right * Screen.TwipsPerPixelX
   
End Function

Public Function scrnTop() As Long
   
  Dim Rc As Rect
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   scrnTop = Rc.top * Screen.TwipsPerPixelY
   
End Function

Public Sub vFitScreen(ByRef rfrmForm As Form, _
                      Optional vblnCenter As Boolean = True, _
                      Optional vblnUseFullScreen As Boolean = False)
   
  Dim Rc          As Rect
  Dim lngScTop    As Long
  Dim lngScBottom As Long
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
   
   lngScTop = Rc.top * Screen.TwipsPerPixelY
   lngScBottom = Rc.bottom * Screen.TwipsPerPixelY
   
   If lngScTop = 0 And Not vblnUseFullScreen Then lngScTop = 250
   rfrmForm.Height = lngScBottom - lngScTop
   If vblnCenter Then CenterForm rfrmForm
   rfrmForm.top = lngScTop
   
End Sub

Public Property Get Width() As Variant
   
   Width = GetSystemMetrics(0&)
   
End Property

