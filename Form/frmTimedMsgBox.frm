VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E1FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   2730
   ClientTop       =   2580
   ClientWidth     =   5835
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTimedMsgBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   690
      TabIndex        =   6
      Top             =   930
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Help"
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
      Index           =   3
      Left            =   3405
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
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
      Index           =   2
      Left            =   2355
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
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
      Index           =   1
      Left            =   1305
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
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
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer tmrCountDown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   5430
      ToolTipText     =   "Close"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   2
      Left            =   5295
      Picture         =   "frmTimedMsgBox.frx":000C
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   1
      Left            =   4935
      Picture         =   "frmTimedMsgBox.frx":0410
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   0
      Left            =   4575
      Picture         =   "frmTimedMsgBox.frx":0837
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label txtMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Top             =   495
      UseMnemonic     =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009EF5F3&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   375
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 2003-2004
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

'// Used to keep form always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
      ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'// Used to get screen size
Private Type Rect
   left       As Long
   top        As Long
   right      As Long
   bottom     As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
      (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA As Long = 48&

'// Used to get positions of cursor
Private Type POINTAPI
   X  As Long
   Y  As Long
End Type
Private CursorXY                           As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'// Button and Icon types
Public Enum enuShowIconTypes
   None_i = 0
   vbCritical = 16         '// Display Critical Message icon.
   vbQuestion = 32         '// Display Warning Query icon.
   vbExclamation = 48      '// Display Warning Message icon.
   vbInformation = 64      '// Display Information Message icon.
   WinLogo_I = 128         '// Display WinLogo icon.
   Folder_I = 144          '// Display Folder icon.
   Printer_I = 160         '// Display Printer icon.
   Find_I = 176            '// Display Find icon.
   Save_I = 240            '// Display Save icon.
   Hourglass_I = 80        '// Display Hourglass icon.
   
   vbDefaultButton1 = 0    '// First button is default.
   vbDefaultButton2 = 256  '// Second button is default.
   vbDefaultButton3 = 512  '// Third button is default.
   vbDefaultButton4 = 768  '// Fourth button is default.
   
   vbOKCancel = 1          '// Display OK and Cancel buttons.
   vbAbortRetryIgnore = 2  '// Display Abort, Retry, and Ignore buttons.
   vbYesNoCancel = 3       '// Display Yes, No, and Cancel buttons.
   vbYesNo = 4             '// Display Yes and No buttons.
   vbRetryCancel = 5       '// Display Retry and Cancel buttons.
   vbOkButton = 6          '// Display OK button only.
   vbMsgBoxHelpButton = 16384 '// Display the Help button
   
   vbHelp = 8              '// Help button pressed
End Enum

'// Used for moving the form around by draging the caption bar
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
      (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'// Used to draw the form's border
Private Declare Function RoundRect Lib "gdi32" _
      (ByVal hdc As Long, ByVal left As Long, ByVal top As Long, ByVal right As Long, _
      ByVal bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'// Used to round the corners of the form and make trasnparent
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
      ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" _
      (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
      ByVal RectY2 As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'// Used to play system sounds
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Const MB_IconAsterisk    As Long = &H10&
Private Const MB_IconQuestion    As Long = &H20&
Private Const MB_IconExclamation As Long = &H30&
Private Const MB_IconInformation As Long = &H40&

'// Used to draw system icons
Private Enum enuSystemIconConstants
   IDI_Application = 32512
   IDI_Error = 32513       'vbCritical (Critical)
   IDI_Question = 32514    'vbQuestion
   IDI_Warning = 32515     'vbExlamation (Exclamation)
   IDI_Information = 32516 'vbInformation (Asterisk)
   IDI_WinLogo = 32517
End Enum
Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" _
      (ByVal hInstance As Long, ByVal lpIconNum As enuSystemIconConstants) As Long
Private Declare Function DrawIcon Lib "user32" _
      (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

'// Used to draw system icons from Shell32.dll
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" _
      (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, _
      ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
      (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'// GradientFill API - Requires Windows 2000 or later; Requires Windows 98 or later
Private Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type
Private Type TRIVERTEX
   X       As Long
   Y       As Long
   Red     As Integer    '// Ushort value (-256 to 0)
   Green   As Integer    '// Ushort value (-256 to 0)
   Blue    As Integer    '// Ushort value (-256 to 0)
   Alpha   As Integer    '// Ushort value (-256 to 0)
End Type
Private Const GRADIENT_FILL_RECT_H         As Long = &H0&
Private Const GRADIENT_FILL_RECT_V         As Long = &H1&
Private Const GRADIENT_FILL_TRIANGLE       As Long = &H2&
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" _
      (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
      ByRef pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'// Operating system version information
Private Type OSVersionInfo
   OSVSize       As Long
   dwVerMajor    As Long
   dwVerMinor    As Long
   dwBuildNumber As Long
   PlatformID    As Long
   szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
      (ByRef lpVersionInformation As OSVersionInfo) As Long

'// Form Variables
Private mlngStandardIcon     As Long
Private mstrCaption          As String
Private mlngAutoCloseSeconds As Long
Private mintButtonResponse   As Integer
Private mbytButtonFocus      As Byte
Private mblnNonModal         As Boolean
Private mblnInputBox         As Boolean
Private mlngCountDown        As Long

Private Sub CheckIfLoaded()
   
  Dim Frm As Form
   
   On Local Error Resume Next
   For Each Frm In Forms
      If LCase$(Frm.Name) = "frmmsgbox" Then
         Unload Frm
         Exit For
      End If
   Next Frm
   
End Sub

Private Sub cmdButton_Click(Index As Integer)
   
   mintButtonResponse = cmdButton(Index).Tag
   Me.Hide
   If mblnNonModal Then Unload Me
   
End Sub

Private Sub DisplayInputBox(ByVal vstrPrompt As String, _
                            ByVal vstrTitle As String, _
                            Optional ByVal vstrDefault As String = vbNullString, _
                            Optional ByVal vblnShowClose As Boolean = True, _
                            Optional ByVal vblnCenter As Boolean = False, _
                            Optional ByVal vstrFont As String = "Tahoma")
   
  Dim bytI     As Byte
  Dim lngIcon  As Long
  Dim lngPosX  As Long
  Dim lngPosY  As Long
  Dim lngWidth As Long
   
   '// Set defaults
   On Error Resume Next
   Me.ScaleMode = vbPixels
   Me.DrawWidth = 1
   Me.FillStyle = 1
   Me.Font = vstrFont
   txtMessage.Font = vstrFont
   txtMessage.FontSize = 10
   lblCaption.Font = vstrFont
   imgClose.Picture = imgX(0).Picture
   On Error GoTo 0
   
   '// Get display position from mouse position
   Call GetCursorPos(CursorXY)
   lngPosX = CursorXY.X * Screen.TwipsPerPixelX
   lngPosY = CursorXY.Y * Screen.TwipsPerPixelY
   
   mstrCaption = vstrTitle
   txtUserText.Text = vstrDefault
   
   '// Resize the Form's width to fit the title bar/messagebox width
   Me.FontSize = 10
   lngWidth = 5000
   Me.FontSize = 8
   If lngWidth < (Me.TextWidth(vstrPrompt) + 90) * Screen.TwipsPerPixelX Then
      lngWidth = (Me.TextWidth(vstrPrompt) + 90) * Screen.TwipsPerPixelX
   End If
   lblCaption.Caption = vstrTitle
   
   Me.Width = lngWidth
   Me.Height = 800
   
   '// Resize the Form's height based on the amount of text to display
   txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
   txtMessage.Caption = vstrPrompt
   If txtMessage.top + txtMessage.Height >= Me.ScaleHeight - 10 Then
      Me.Height = (txtMessage.top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
   End If
   
   txtUserText.Move 25, txtMessage.top + txtMessage.Height + 10, txtMessage.Width - 25
   
   '// Locate Buttons and resize Form if required
   If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
      '// How many buttons are visible?
      If Val(cmdButton(1).Tag) > 0 Then bytI = 1
      If Val(cmdButton(2).Tag) > 0 Then bytI = 2
      If Val(cmdButton(3).Tag) > 0 Then bytI = 3
      
      cmdButton(0).top = txtUserText.top + txtUserText.Height + 10
      cmdButton(1).top = txtUserText.top + txtUserText.Height + 10
      cmdButton(2).top = txtUserText.top + txtUserText.Height + 10
      cmdButton(3).top = txtUserText.top + txtUserText.Height + 10
      If Me.Width < (cmdButton(bytI).left + cmdButton(bytI).Width + 15) * Screen.TwipsPerPixelX Then
         Me.Width = (cmdButton(bytI).left + cmdButton(bytI).Width + 15) * Screen.TwipsPerPixelX
      End If
      
      Me.Height = (cmdButton(0).top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY
   End If
   
   '// Show or don't show the close button
   If vblnShowClose Then
      imgClose.Visible = True
    Else
      imgClose.Visible = False
   End If
   
   '// Locate title bar and close button
   imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
   lblCaption.Move 2, 5, Me.ScaleWidth, 25
   
   Call GradientFill
   
   '// Draw box around Title Bar
   Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
   
   '// Draw border around the Form
   Me.ForeColor = &H80000015
   RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, _
         (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
   Me.ForeColor = &H8000000F
   RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, _
         (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
   
   '// Make corners transparent
   SetWindowRgn Me.hWnd, _
         CreateRoundRectRgn(0, 0, _
         (Me.Width / Screen.TwipsPerPixelX), _
         (Me.Height / Screen.TwipsPerPixelY), _
         25, 25), True
   
   '// Position form on screen
   If Not vblnCenter Then
      Me.Move lngPosX, lngPosY
   End If
   Call PositionForm(vblnCenter)
   
   lngIcon = LoadStandardIcon(0&, IDI_Question)
   Call DrawIcon(Me.hdc, 4&, 4&, lngIcon)
   DestroyIcon lngIcon
   
   txtUserText.SelStart = 0
   txtUserText.SelLength = Len(txtUserText.Text)
   
End Sub

Private Sub DisplayMessage(ByVal vstrText As String, _
                           Optional ByVal venuIcon As enuShowIconTypes = None_i, _
                           Optional ByVal vstrTitle As String = vbNullString, _
                           Optional ByVal vlngAutoCloseSeconds As Long = 0, _
                           Optional ByVal vblnShowClose As Boolean = True, _
                           Optional ByVal vblnCenter As Boolean = False, _
                           Optional ByVal vlngWidth As Long = -1, _
                           Optional ByVal vstrFont As String = "Tahoma")
   
  Dim bytI           As Byte
  Dim lngIcon        As Long
  Dim lngPosX        As Long
  Dim lngPosY        As Long
  Dim blnShell32Icon As Boolean
   
   '// Set defaults
   On Error Resume Next
   Me.ScaleMode = vbPixels
   Me.DrawWidth = 1
   Me.FillStyle = 1
   Me.Font = vstrFont
   txtMessage.Font = vstrFont
   txtMessage.FontSize = 10
   lblCaption.Font = vstrFont
   imgClose.Picture = imgX(0).Picture
   On Error GoTo 0
   
   '// Get display position from mouse position
   Call GetCursorPos(CursorXY)
   lngPosX = CursorXY.X * Screen.TwipsPerPixelX
   lngPosY = CursorXY.Y * Screen.TwipsPerPixelY
   
   '// Set Title bar
   Select Case venuIcon
    Case vbInformation '// The "bytI" icon - Information
      If vstrTitle = vbNullString Then vstrTitle = "Information"
      MessageBeep MB_IconInformation
      mlngStandardIcon = IDI_Information
      
    Case vbCritical '// The "x" icon - Critical
      If vstrTitle = vbNullString Then vstrTitle = "ERROR!"
      MessageBeep MB_IconAsterisk
      mlngStandardIcon = IDI_Error
      
    Case vbExclamation '// The "!" icon - Exclamation
      If vstrTitle = vbNullString Then vstrTitle = "Warning!"
      MessageBeep MB_IconExclamation
      mlngStandardIcon = IDI_Warning
      
    Case vbQuestion '// The "?" icon - Question
      If vstrTitle = vbNullString Then vstrTitle = "Question?"
      MessageBeep MB_IconQuestion
      mlngStandardIcon = IDI_Question
      
    Case WinLogo_I '// Winlogo icon
      mlngStandardIcon = IDI_WinLogo
      
    Case Printer_I '// Printer icon
      If vstrTitle = vbNullString Then vstrTitle = "Printing.. Please Wait"
      MessageBeep MB_IconInformation
      mlngStandardIcon = 16
      blnShell32Icon = True
      
    Case Folder_I '// Open folder icon
      MessageBeep MB_IconInformation
      mlngStandardIcon = 4
      blnShell32Icon = True
      
    Case Find_I '// Find icon
      MessageBeep MB_IconInformation
      mlngStandardIcon = 22
      blnShell32Icon = True
      
    Case Save_I '// Save icon
      MessageBeep MB_IconInformation
      mlngStandardIcon = 6
      blnShell32Icon = True
      
    Case Hourglass_I '// Hourglass icon
      If vstrTitle = vbNullString Then vstrTitle = "Working.. Please Wait"
      MessageBeep MB_IconInformation
      mlngStandardIcon = 76
      blnShell32Icon = True
      
    Case Else 'Use no icon
      
   End Select
   mstrCaption = vstrTitle
   
   '// Resize the Form's width to fit the title bar/messagebox width
   Me.FontSize = 10
   If vlngWidth = -1 Then
      vlngWidth = (Me.TextWidth(vstrText) + 20) * Screen.TwipsPerPixelX
      If vlngWidth > 5000 Then vlngWidth = 5000
   End If
   If vlngWidth < 1500 Then vlngWidth = 1500
   If vlngAutoCloseSeconds > 0 Then
      If vstrTitle > vbNullString Then
         vstrTitle = vstrTitle & " -" & CStr(vlngAutoCloseSeconds)
       Else
         vstrTitle = CStr(vlngAutoCloseSeconds)
      End If
   End If
   Me.FontSize = 8
   If vlngWidth < (Me.TextWidth(vstrTitle) + 90) * Screen.TwipsPerPixelX Then
      vlngWidth = (Me.TextWidth(vstrTitle) + 90) * Screen.TwipsPerPixelX
   End If
   lblCaption.Caption = vstrTitle
   
   Me.Width = vlngWidth
   Me.Height = 800
   
   '// Resize the Form's height based on the amount of text to display
   txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
   txtMessage.Caption = vstrText
   If txtMessage.top + txtMessage.Height >= Me.ScaleHeight - 10 Then
      Me.Height = (txtMessage.top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
   End If
   
   '// Locate Buttons and resize Form if required
   If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
      '// How many buttons are visible?
      If Val(cmdButton(1).Tag) > 0 Then bytI = 1
      If Val(cmdButton(2).Tag) > 0 Then bytI = 2
      If Val(cmdButton(3).Tag) > 0 Then bytI = 3
      
      'Me.Height = Me.Height + 500
      cmdButton(0).top = txtMessage.top + txtMessage.Height + 10
      cmdButton(1).top = txtMessage.top + txtMessage.Height + 10
      cmdButton(2).top = txtMessage.top + txtMessage.Height + 10
      cmdButton(3).top = txtMessage.top + txtMessage.Height + 10
      If Me.Width < (cmdButton(bytI).left + cmdButton(bytI).Width + 15) * Screen.TwipsPerPixelX Then
         Me.Width = (cmdButton(bytI).left + cmdButton(bytI).Width + 15) * Screen.TwipsPerPixelX
      End If
      Me.Height = (cmdButton(0).top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY 'Me.Height + 500
   End If
   
   '// Show or don't show the close button
   If vblnShowClose Then
      imgClose.Visible = True
    Else
      imgClose.Visible = False
   End If
   
   '// Enable or disable auto close timer
   If vlngAutoCloseSeconds = 0 Then
      tmrCountDown.Enabled = False
    Else
      If mstrCaption > vbNullString Then mstrCaption = mstrCaption & " -"
      mlngAutoCloseSeconds = vlngAutoCloseSeconds
      tmrCountDown.Enabled = True
   End If
   
   Call GradientFill
   
   '// Locate title bar and close button
   imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
   lblCaption.Move 2, 5, Me.ScaleWidth, 25
   
   '// Draw box around Title Bar
   Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
   
   '// Draw border around the Form
   Me.ForeColor = &H80000015
   RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
   Me.ForeColor = &H8000000F
   RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
   
   '// Draw Icon
   If blnShell32Icon Then
      Call LoadShell32Icon(mlngStandardIcon)
    Else
      lngIcon = LoadStandardIcon(0&, mlngStandardIcon)
      Call DrawIcon(Me.hdc, 4&, 4&, lngIcon)
      DestroyIcon lngIcon
   End If
   
   '// Make corners transparent
   SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), 25, 25), True
   
   '// Position form on screen
   If Not vblnCenter Then
      Me.Move lngPosX, lngPosY
   End If
   Call PositionForm(vblnCenter)
   
End Sub

Private Sub Form_Activate()
   
   If cmdButton(mbytButtonFocus).Visible Then
      cmdButton(mbytButtonFocus).SetFocus
   End If
   If mblnInputBox Then txtUserText.SetFocus
   
   Me.ZOrder
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   imgClose.Picture = imgX(0).Picture
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmMsgBox = Nothing
   
End Sub

Private Sub GradientFill()
   
  Dim udtTriVert(4) As TRIVERTEX
  Dim udtTRi(1)     As GRADIENT_TRIANGLE
  Dim udtOSV        As OSVersionInfo
  Dim lngOSver      As Long
   
   '// Get OS compatability flag
   udtOSV.OSVSize = Len(udtOSV)
   If GetVersionEx(udtOSV) = 1 Then
      If udtOSV.PlatformID = 1 And udtOSV.dwVerMinor >= 10 Then lngOSver = 1 '// Win 98/ME
      If udtOSV.PlatformID = 2 And udtOSV.dwVerMajor >= 5 Then lngOSver = 2  '// Win 2000/XP
   End If
   '// Requires Windows 2000 or later; Requires Windows 98/ME
   If lngOSver = 0 Then Exit Sub
   
   Me.AutoRedraw = True
   
   '// Top Left Trangle
   udtTriVert(0).X = 0
   udtTriVert(0).Y = 0
   Call GradientFillColor(udtTriVert(0), vbWhite)
   
   '// Top Right Trangle
   udtTriVert(1).X = Me.ScaleWidth
   udtTriVert(1).Y = 0
   Call GradientFillColor(udtTriVert(1), vbWhite)
   
   '// Bottom Right Trangle
   udtTriVert(2).X = Me.ScaleWidth
   udtTriVert(2).Y = Me.ScaleHeight
   Call GradientFillColor(udtTriVert(2), &HC0FFFF)
   
   '// Bottom Left Trangle
   udtTriVert(3).X = 0
   udtTriVert(3).Y = Me.ScaleHeight
   Call GradientFillColor(udtTriVert(3), vbWhite)
   
   udtTRi(0).Vertex1 = 0
   udtTRi(0).Vertex2 = 1
   udtTRi(0).Vertex3 = 2
   
   udtTRi(1).Vertex1 = 0
   udtTRi(1).Vertex2 = 2
   udtTRi(1).Vertex3 = 3
   
   Call GradientFillTriangle(Me.hdc, udtTriVert(0), 4, udtTRi(0), 2, GRADIENT_FILL_TRIANGLE)
   
End Sub

Private Sub GradientFillColor(ByRef rudtTV As TRIVERTEX, _
                              ByVal vlngColor As Long)
   
  Dim lngRed   As Long
  Dim lngGreen As Long
  Dim lngBlue  As Long
   
   '// Separate color into RGB
   lngRed = (vlngColor And &HFF&) * &H100&
   lngGreen = (vlngColor And &HFF00&)
   lngBlue = (vlngColor And &HFF0000) \ &H100&
   
   '// Make Red color a UShort
   If (lngRed And &H8000&) = &H8000& Then
      rudtTV.Red = (lngRed And &H7F00&)
      rudtTV.Red = rudtTV.Red Or &H8000
    Else
      rudtTV.Red = lngRed
   End If
   '// Make Green color a UShort
   If (lngGreen And &H8000&) = &H8000& Then
      rudtTV.Green = (lngGreen And &H7F00&)
      rudtTV.Green = rudtTV.Green Or &H8000
    Else
      rudtTV.Green = lngGreen
   End If
   '// Make Blue color a UShort
   If (lngBlue And &H8000&) = &H8000& Then
      rudtTV.Blue = (lngBlue And &H7F00&)
      rudtTV.Blue = rudtTV.Blue Or &H8000
    Else
      rudtTV.Blue = lngBlue
   End If
   
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If Button = vbLeftButton Then imgClose.Picture = imgX(2).Picture
   
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If imgClose.Picture <> imgX(2).Picture Then
      imgClose.Picture = imgX(1).Picture
   End If
   
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Unload Me
   
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If Me.WindowState <> vbMaximized Then
      ReleaseCapture
      SendMessage Me.hWnd, &HA1, 2, 0&
   End If
   
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   imgClose.Picture = imgX(0).Picture
   
End Sub

Private Sub LoadShell32Icon(ByVal vlngIndex As Long)
   
  Dim strSysDir    As String
  Dim strCurFile   As String
  Dim lngIcon      As Long
  Dim lngIconCount As Long
  Dim lngRv        As Long
   
   strSysDir = Space$(260)
   lngRv = GetSystemDirectory(strSysDir, 260)
   strSysDir = left(strSysDir, lngRv) & "\"
   
   strCurFile = strSysDir & "Shell32.dll"
   lngIconCount = ExtractIconEx(strCurFile, -1, 0, 0, 0)
   
   If lngIconCount >= vlngIndex Then
      Call ExtractIconEx(strCurFile, vlngIndex, lngIcon, 0&, 1&)
      Call DrawIcon(Me.hdc, 4&, 4&, lngIcon)
      DestroyIcon lngIcon
   End If
   
End Sub

Private Sub PositionForm(ByVal vblnCenter As Boolean)
   
  Dim Rc            As Rect
  Dim lngTop        As Long
  Dim lngBottom     As Long
  Dim lngLeft       As Long
  Dim lngRight      As Long
  Dim lngTopT       As Long
  Dim lngLeftT      As Long
  Const C_lngOffset As Long = 150&
   
   '// Get screen size
   SystemParametersInfo SPI_GETWORKAREA, 0&, Rc, 0&
   lngTop = Rc.top * Screen.TwipsPerPixelY
   lngBottom = Rc.bottom * Screen.TwipsPerPixelY
   lngLeft = Rc.left * Screen.TwipsPerPixelX
   lngRight = Rc.right * Screen.TwipsPerPixelX
   
   If vblnCenter Then
      '// vblnCenter Form on screen
      lngTopT = Abs((lngBottom / 2) - (Me.Height / 2))
      lngLeftT = Abs((lngRight / 2) - (Me.Width / 2))
      
      If lngTopT < lngTop Then lngTopT = lngTop
      If lngTopT > lngBottom - Me.Height Then lngTopT = lngBottom - Me.Height
      If lngLeftT < lngLeft Then lngLeftT = lngLeft
    Else
      '// Make sure all the Form is on the screen
      lngTopT = Me.top
      lngLeftT = Me.left
      
      If Me.top - C_lngOffset < lngTop Then lngTopT = lngTop + C_lngOffset
      If Me.left - C_lngOffset < lngLeft Then lngLeftT = lngLeft + C_lngOffset
      If Me.top + Me.Height + C_lngOffset > lngBottom Then lngTopT = lngBottom - Me.Height - C_lngOffset
      If Me.left + Me.Width + C_lngOffset > lngRight Then lngLeftT = lngRight - Me.Width - C_lngOffset
   End If
   
   Me.Move lngLeftT, lngTopT
   
End Sub

Public Function SInputBox(ByVal vstrPrompt As String, _
                          Optional ByVal vstrTitle As String = vbNullString, _
                          Optional ByVal vstrDefault As String = vbNullString, _
                          Optional ByVal vblnShowClose As Boolean = False, _
                          Optional ByVal vblnCenter As Boolean = True, _
                          Optional ByVal vstrFont As String = "Tahoma") As String
   
   Call CheckIfLoaded
   
   With cmdButton(0)
      .Visible = True
      .Caption = "Ok"
      .Tag = vbOK
      .Default = True
   End With
   With cmdButton(1)
      .Visible = True
      .Caption = "Cancel"
      .Tag = vbCancel
      .Cancel = True
   End With
   
   txtUserText.Visible = True
   mblnInputBox = True
   
   If vstrTitle = vbNullString Then vstrTitle = App.Title
   
   DisplayInputBox vstrPrompt, vstrTitle, vstrDefault, vblnShowClose, vblnCenter, vstrFont
   
   DoEvents
   Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
   Me.Show vbModal
   
   If mintButtonResponse = vbCancel Then
      SInputBox = vbNullString
    Else
      SInputBox = txtUserText.Text
   End If
   
   DoEvents
   Unload Me
   
End Function

Public Function sMessage(ByVal vstrText As String, _
                         Optional ByVal venuIcon As enuShowIconTypes = None_i, _
                         Optional ByVal vstrTitle As String = vbNullString, _
                         Optional ByVal vlngAutoCloseSeconds As Long = 0, _
                         Optional ByVal vblnShowClose As Boolean = True, _
                         Optional ByVal vblnCenter As Boolean = True, _
                         Optional ByVal lngWidth As Long = -1, _
                         Optional ByVal vstrFont As String = "Tahoma", _
                         Optional ByRef rOwnerForm As Form) As Integer
   
  Dim udtMsgType As enuShowIconTypes
  Dim blnTesthDC As Boolean
   
   Call CheckIfLoaded
   
   '// Separate Message Icon from input
   udtMsgType = venuIcon And 240
   
   '// Only the OK button allowed for a non-modal message box
   If (venuIcon And 15) = vbOkButton Then
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Ok"
      cmdButton(0).Tag = vbOK
      cmdButton(0).Cancel = True
      mblnNonModal = True
   End If
   
   Call DisplayMessage(vstrText, udtMsgType, vstrTitle, vlngAutoCloseSeconds, _
         vblnShowClose, vblnCenter, lngWidth, vstrFont)
   
   DoEvents
   On Local Error Resume Next
   blnTesthDC = rOwnerForm.HasDC
   If Not blnTesthDC Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
   
   Show , rOwnerForm
   DoEvents
   Me.ZOrder
   
End Function

Public Function SMessageModal(ByVal vstrText As String, _
                              Optional ByVal venuIcon As enuShowIconTypes = None_i, _
                              Optional ByVal vstrTitle As String = vbNullString, _
                              Optional ByVal vlngAutoCloseSeconds As Long = 0, _
                              Optional ByVal vblnShowClose As Boolean = True, _
                              Optional ByVal vblnCenter As Boolean = True, _
                              Optional ByVal lngWidth As Long = -1, _
                              Optional ByVal vstrFont As String = "Tahoma", _
                              Optional ByRef rOwnerForm As Form) As Integer
   
  Dim udtMsgType As enuShowIconTypes
  Dim blnTesthDC As Boolean
   
   Call CheckIfLoaded
   
   '// Separate Message Icon from input
   udtMsgType = venuIcon And 240
   
   '// Separate button default from input
   Select Case venuIcon And 1792
    Case 256
      mbytButtonFocus = 1 '// Second button is default.
    Case 512
      mbytButtonFocus = 2 '// Third button is default.
    Case 768
      mbytButtonFocus = 3 '// Fourth button is default.
    Case Else
      mbytButtonFocus = 0 '// First button is default.
   End Select
   
   '// Separate Button type from input
   If vlngAutoCloseSeconds = 0 Then vblnShowClose = True
   Select Case venuIcon And 15
    Case vbRetryCancel
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Retry"
      cmdButton(0).Tag = vbRetry
      cmdButton(1).Visible = True
      cmdButton(1).Caption = "Cancel"
      cmdButton(1).Tag = vbCancel
      cmdButton(1).Cancel = True
      vblnShowClose = False
      vlngAutoCloseSeconds = 0
    Case vbYesNo
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Yes"
      cmdButton(0).Tag = vbYes
      cmdButton(1).Visible = True
      cmdButton(1).Caption = "No"
      cmdButton(1).Tag = vbNo
      vblnShowClose = False
      vlngAutoCloseSeconds = 0
    Case vbYesNoCancel
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Yes"
      cmdButton(0).Tag = vbYes
      cmdButton(1).Visible = True
      cmdButton(1).Caption = "No"
      cmdButton(1).Tag = vbNo
      cmdButton(2).Visible = True
      cmdButton(2).Caption = "Cancel"
      cmdButton(2).Tag = vbCancel
      cmdButton(2).Cancel = True
      vblnShowClose = False
      vlngAutoCloseSeconds = 0
    Case vbAbortRetryIgnore
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Abort"
      cmdButton(0).Tag = vbAbort
      cmdButton(1).Visible = True
      cmdButton(1).Caption = "Retry"
      cmdButton(1).Tag = vbRetry
      cmdButton(2).Visible = True
      cmdButton(2).Caption = "Ignore"
      cmdButton(2).Tag = vbIgnore
      vblnShowClose = False
    Case vbOKCancel
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Ok"
      cmdButton(0).Tag = vbOK
      cmdButton(1).Visible = True
      cmdButton(1).Caption = "Cancel"
      cmdButton(1).Tag = vbCancel
      cmdButton(1).Cancel = True
      vblnShowClose = False
      vlngAutoCloseSeconds = 0
    Case Else
      cmdButton(0).Visible = True
      cmdButton(0).Caption = "Ok"
      cmdButton(0).Tag = vbOK
      cmdButton(0).Cancel = True
      
   End Select
   '// Show Help button?
   If venuIcon And 16384 Then
      cmdButton(3).Visible = True
      cmdButton(3).Tag = vbHelp
   End If
   
   Call DisplayMessage(vstrText, udtMsgType, vstrTitle, vlngAutoCloseSeconds, _
         vblnShowClose, vblnCenter, lngWidth, vstrFont)
   
   DoEvents
   
   On Local Error Resume Next
   blnTesthDC = rOwnerForm.HasDC
   If Not blnTesthDC Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
   
   Me.Show vbModal
   SMessageModal = mintButtonResponse
   DoEvents
   
   Unload Me
   
End Function

Private Sub tmrCountDown_Timer()
   
   mlngCountDown = mlngCountDown + 1
   If mlngCountDown >= mlngAutoCloseSeconds Then
      Unload Me
    Else
      lblCaption.Caption = mstrCaption & CStr(mlngAutoCloseSeconds - mlngCountDown)
   End If
   
End Sub

