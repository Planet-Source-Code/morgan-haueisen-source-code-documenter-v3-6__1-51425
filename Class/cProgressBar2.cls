VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net       */
'// Copyright (c) 1997-2004
'//************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, vlngWhole or in vsngPart, as source code or in binary form, alone or
'        as vsngPart of a larger distribution or product, is forbidden for any commercial or for-profit
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
'        Redistributions in binary form, as vsngPart of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

Private mstrTextBefore      As String
Private mstrTextAfter       As String

Private mblnShowStatus      As Boolean
Private mblnShowCounts      As Boolean

Private mblnHorizontal      As Boolean
Private mblnBackwards       As Boolean
Private mlngBackColor       As Long
Private mlngForeColor       As Long
Private mlngForeColor2      As Long
Private mlngTextColor       As Long
Private msngMax             As Single
Private msngValue           As Single
Private mlngStyle           As Long
Private mpicPicBox          As PictureBox

Public Enum enuBarType
   pbSolidBar = 0
   pbSolid2Color = 1
   pbStepped = 2
   pbStepped2Color = 4
   pbValueDependant = 3
End Enum

Public Property Get BackColor() As Long
   
   BackColor = mlngBackColor
   
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
   
   mlngBackColor = vNewValue
   
End Property

Public Property Get Backwards() As Boolean
   
   Backwards = mblnBackwards
   
End Property

Public Property Let Backwards(ByVal New_Backwards As Boolean)
   
   mblnBackwards = New_Backwards
   
End Property

Private Sub Class_Initialize()
   
   mlngBackColor = vbWhite 'vbWindowBackground
   mlngForeColor = RGB(100, 100, 255)
   mlngForeColor2 = RGB(100, 255, 100)
   mlngTextColor = vbBlack
   
   mstrTextBefore = vbNullString
   mstrTextAfter = vbNullString
   
   mblnHorizontal = True
   mblnShowStatus = True
   mblnShowCounts = True
   
End Sub

Private Sub DrawBoarder(ByRef vPicStatus As PictureBox)
   
   '// Draw 3D boarder */
   
   With vPicStatus
      .DrawMode = vbCopyPen '// Copy Pen
      vPicStatus.Line (0, 0)-(.ScaleWidth, .ScaleHeight), vb3DDKShadow, B
      vPicStatus.Line (1, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), vb3DHighlight
      vPicStatus.Line (.ScaleWidth - 1, 1)-(.ScaleWidth - 1, .ScaleHeight), vb3DHighlight
      .Refresh
   End With
   
End Sub

Public Property Get ForeColor() As Long
   
   ForeColor = mlngForeColor
   
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
   
   mlngForeColor = vNewValue
   
End Property

Public Property Get ForeColor2() As Long
   
   ForeColor2 = mlngForeColor2
   
End Property

Public Property Let ForeColor2(ByVal vNewValue As Long)
   
   mlngForeColor2 = vNewValue
   
End Property

Private Sub GetRGB(ByVal vlngColor As Long, _
                   ByRef rlngRed As Long, _
                   ByRef rlngGreen As Long, _
                   ByRef rlngBlue As Long)
   
  Dim srtHV As String
   
   If vlngColor < 0 Then
      rlngRed = 105
      rlngGreen = 105
      rlngBlue = 255
    Else
      srtHV = Hex$(vlngColor)
      srtHV = right("000000" & srtHV, 6)
      rlngRed = CLng("&H" & right(srtHV, 2))
      rlngGreen = CLng("&H" & Mid$(srtHV, 3, 2))
      rlngBlue = CLng("&H" & left(srtHV, 2))
   End If
   
End Sub

Public Property Get Horizontal() As Boolean
   
   Horizontal = mblnHorizontal
   
End Property

Public Property Let Horizontal(ByVal New_Horizontal As Boolean)
   
   mblnHorizontal = New_Horizontal
   
End Property

Public Property Get Max() As Single
   
   Max = msngMax
   
End Property

Public Property Let Max(ByVal vNewValue As Single)
   
   msngMax = vNewValue
   
End Property

Public Property Get picBox() As PictureBox
   
End Property

Public Property Let picBox(ByRef vNewValue As PictureBox)
   
   Set mpicPicBox = vNewValue
   
End Property

Public Sub Progress2ColorBar(ByRef vPicStatus As PictureBox, _
                             Optional ByVal vsngPart As Single = 0, _
                             Optional ByVal vlngWhole As Single = 0)
   
  Dim sngPercent As Single
  Dim lngI       As Long
  Dim lngX       As Long
  Dim lngY       As Long
  Dim lngWidth   As Long
  Dim lngHeight  As Long
  Dim strCaption As String
  Dim lngRed1    As Long
  Dim lngGrn1    As Long
  Dim lngBlu1    As Long
  Dim lngRed2    As Long
  Dim lngGrn2    As Long
  Dim lngBlu2    As Long
  Dim sngRed     As Single
  Dim sngGrn     As Single
  Dim sngBlu     As Single
  Dim lngMaxH    As Long
   
   On Error Resume Next
   
   If vsngPart = 0 Then vsngPart = msngValue
   If vlngWhole = 0 Then vlngWhole = msngMax
   
   Call SetDefaults(vPicStatus)
   
   If vlngWhole > 0 Then
      '// calculate the percentage from current value and total value
      sngPercent = vsngPart / vlngWhole
    Else
      sngPercent = 0
   End If
   
   If mblnShowStatus Or mblnShowCounts Or LenB(mstrTextBefore) Or LenB(mstrTextAfter) Then
      If mblnShowCounts Then
         strCaption = CStr(vsngPart) & " of " & CStr(vlngWhole)
      End If
      If mblnShowStatus Then
         If LenB(strCaption) Then
            strCaption = strCaption & " (" & Format$(sngPercent, "0%") & ")"
          Else
            strCaption = Format$(sngPercent, "0%")
         End If
      End If
      strCaption = IIf(mstrTextBefore > vbNullString, mstrTextBefore & " ", vbNullString) & strCaption
      strCaption = strCaption & IIf(mstrTextAfter > vbNullString, " " & mstrTextAfter, vbNullString)
      With vPicStatus
         '// Back to the center print position and print the text
         lngWidth = .TextWidth(strCaption)
         lngHeight = .TextHeight(strCaption)
         If mblnHorizontal Then
            If .ScaleHeight < lngHeight Then .Height = lngHeight * 15
          Else
            If .ScaleWidth < lngWidth Then .Width = lngWidth * 15
         End If
         
         '// Now set lngX and lngY to the starting location for frmPrinting the percentage
         lngX = (.ScaleWidth / 2) - (lngWidth / 2)
         lngY = (.ScaleHeight / 2) - (lngHeight / 2)
         
         .CurrentX = lngX
         .CurrentY = lngY
         .ForeColor = mlngTextColor
         vPicStatus.Print strCaption
         
      End With
   End If
   
   Call GetRGB(mlngForeColor, lngRed1, lngGrn1, lngBlu1)
   Call GetRGB(mlngForeColor2, lngRed2, lngGrn2, lngBlu2)
   
   '// Use the "Not XOR" pen so that we change the color of the text wherever we touch it
   vPicStatus.DrawMode = vbMergePenNot 'vbNotXorPen
   
   If sngPercent > 0 Then
      If mblnHorizontal Then
         
         lngMaxH = CInt(vPicStatus.ScaleWidth * vsngPart / vlngWhole)
         sngRed = (lngRed2 - lngRed1) / lngMaxH
         sngGrn = (lngGrn2 - lngGrn1) / lngMaxH
         sngBlu = (lngBlu2 - lngBlu1) / lngMaxH
         
         If mblnBackwards Then
            For lngI = 0 To CInt(vPicStatus.ScaleWidth * vsngPart / vlngWhole)
               lngX = vPicStatus.ScaleWidth - lngI
               vPicStatus.Line (lngX, 0)-(lngX, vPicStatus.ScaleHeight), RGB(lngRed1 + sngRed * lngI, lngGrn1 + sngGrn * lngI, lngBlu1 + sngBlu * lngI)
            Next lngI
          Else
            For lngI = 0 To CInt(vPicStatus.ScaleWidth * vsngPart / vlngWhole)
               vPicStatus.Line (lngI, 0)-(lngI, vPicStatus.ScaleHeight), RGB(lngRed1 + sngRed * lngI, lngGrn1 + sngGrn * lngI, lngBlu1 + sngBlu * lngI)
            Next lngI
         End If
         
       Else '// Vertical
         lngMaxH = CInt(vPicStatus.ScaleHeight * vsngPart / vlngWhole)
         sngRed = (lngRed2 - lngRed1) / lngMaxH
         sngGrn = (lngGrn2 - lngGrn1) / lngMaxH
         sngBlu = (lngBlu2 - lngBlu1) / lngMaxH
         
         If mblnBackwards Then
            For lngI = 0 To CInt(vPicStatus.ScaleHeight * vsngPart / vlngWhole)
               vPicStatus.Line (0, lngI)-(vPicStatus.ScaleWidth, lngI), RGB(lngRed1 + sngRed * lngI, lngGrn1 + sngGrn * lngI, lngBlu1 + sngBlu * lngI)
            Next lngI
          Else
            For lngI = 0 To CInt(vPicStatus.ScaleHeight * vsngPart / vlngWhole)
               lngY = vPicStatus.ScaleHeight - lngI
               vPicStatus.Line (0, lngY)-(vPicStatus.ScaleWidth, lngY), RGB(lngRed1 + sngRed * lngI, lngGrn1 + sngGrn * lngI, lngBlu1 + sngBlu * lngI)
            Next lngI
         End If
      End If
      
   End If
   
   '// Draw 3D boarder */
   Call DrawBoarder(vPicStatus)
   
   On Error GoTo 0
   
End Sub

Public Sub ProgressBar(ByRef vPicStatus As PictureBox, _
                       Optional ByVal vsngPart As Single = 0, _
                       Optional ByVal vlngWhole As Single = 0, _
                       Optional ByVal vblnUseOneColor As Boolean = True)
   
  Dim sngPercent As Single
  Dim lngX       As Long
  Dim lngY       As Long
  Dim lngWidth   As Long
  Dim lngHeight  As Long
  Dim strCaption As String
  Dim lngRed1    As Long
  Dim lngGrn1    As Long
  Dim lngBlu1    As Long
  Dim lngRed2    As Long
  Dim lngGrn2    As Long
  Dim lngBlu2    As Long
  Dim sngRed     As Single
  Dim sngGrn     As Single
  Dim sngBlu     As Single
   
   On Error Resume Next
   
   Call SetDefaults(vPicStatus)
   
   If vsngPart = 0 Then vsngPart = msngValue
   If vlngWhole = 0 Then vlngWhole = msngMax
   
   If vlngWhole > 0 Then
      '// calculate the percentage from current value and total value
      sngPercent = vsngPart / vlngWhole
    Else
      sngPercent = 0
   End If
   
   If mblnShowStatus Or mblnShowCounts Or LenB(mstrTextBefore) Or LenB(mstrTextAfter) Then
      If mblnShowCounts Then
         strCaption = CStr(vsngPart) & " of " & CStr(vlngWhole)
      End If
      If mblnShowStatus Then
         If LenB(strCaption) Then
            strCaption = strCaption & " (" & Format$(sngPercent, "0%") & ")"
          Else
            strCaption = Format$(sngPercent, "0%")
         End If
      End If
      strCaption = IIf(mstrTextBefore > vbNullString, mstrTextBefore & " ", vbNullString) & strCaption
      strCaption = strCaption & IIf(mstrTextAfter > vbNullString, " " & mstrTextAfter, vbNullString)
      
      With vPicStatus
         lngWidth = .TextWidth(strCaption)
         lngHeight = .TextHeight(strCaption)
         If mblnHorizontal Then
            If .ScaleHeight < lngHeight Then .Height = lngHeight * 15
          Else
            If .ScaleWidth < lngWidth Then .Width = lngWidth * 15
         End If
         '// Now set lngX and lngY to the starting location for frmPrinting the percentage
         lngX = (.ScaleWidth / 2) - (lngWidth / 2)
         lngY = (.ScaleHeight / 2) - (lngHeight / 2)
         '// Need to draw a filled box with the picStatus background color to wipe out previous
         '// percentage display (if any)
         vPicStatus.Line (lngX, lngY)-(lngWidth, lngHeight), vPicStatus.BackColor
         '// Back to the center pri position and pri the text
         .CurrentX = lngX
         .CurrentY = lngY
         .ForeColor = mlngTextColor
         vPicStatus.Print strCaption
      End With
   End If
   
   '// Now fill in the box with the ribbon color to the desired percentage
   '// Use the "Not XOR" pen so that we change the color of the text wherever we touch it
   vPicStatus.DrawMode = vbMergePenNot 'vbNotXorPen
   
   Call GetRGB(mlngForeColor, lngRed1, lngGrn1, lngBlu1)
   If Not vblnUseOneColor Then
      Call GetRGB(mlngForeColor2, lngRed2, lngGrn2, lngBlu2)
    Else
      lngRed2 = lngRed1
      lngGrn2 = lngGrn1
      lngBlu2 = lngBlu1
   End If
   
   sngRed = (lngRed2 - lngRed1)
   sngGrn = (lngGrn2 - lngGrn1)
   sngBlu = (lngBlu2 - lngBlu1)
   
   On Error Resume Next
   vPicStatus.ForeColor = RGB(CInt(lngRed1 + sngRed * sngPercent), CInt(lngGrn1 + sngGrn * sngPercent), CInt(lngBlu1 + sngBlu * sngPercent))
   
   If mblnHorizontal Then
      If sngPercent > 0 Then
         If mblnBackwards Then
            vPicStatus.Line (vPicStatus.ScaleWidth, 0)-(vPicStatus.ScaleWidth - (vPicStatus.ScaleWidth * sngPercent), vPicStatus.ScaleHeight), vPicStatus.ForeColor, BF
          Else
            vPicStatus.Line (0, 0)-(vPicStatus.ScaleWidth * sngPercent, vPicStatus.ScaleHeight), vPicStatus.ForeColor, BF
         End If
      End If
      
    Else '// Vertical
      If sngPercent > 0 Then
         If mblnBackwards Then
            vPicStatus.Line (0, 0)-(vPicStatus.ScaleWidth, vPicStatus.ScaleHeight * sngPercent), vPicStatus.ForeColor, BF
          Else
            vPicStatus.Line (0, vPicStatus.ScaleHeight)-(vPicStatus.ScaleWidth, vPicStatus.ScaleHeight - (vPicStatus.ScaleHeight * sngPercent)), vPicStatus.ForeColor, BF
         End If
      End If
   End If
   
   '// Draw 3D boarder */
   Call DrawBoarder(vPicStatus)
   
   On Error GoTo 0
   
End Sub

Public Sub ProgressSteppedBar(ByRef vPicStatus As PictureBox, _
                              Optional ByVal vsngPart As Single = 0, _
                              Optional ByVal vlngWhole As Single = 0, _
                              Optional ByVal vblnUseOneColor As Boolean = False)
   
  Dim sngPercent As Single
  Dim lngI       As Long
  Dim lngX       As Long
  Dim lngY       As Long
  Dim lngWidth   As Long
  Dim lngHeight  As Long
  Dim strCaption As String
  Dim lngRed1    As Long
  Dim lngGrn1    As Long
  Dim lngBlu1    As Long
  Dim lngRed2    As Long
  Dim lngGrn2    As Long
  Dim lngBlu2    As Long
  Dim sngRed     As Single
  Dim sngGrn     As Single
  Dim sngBlu     As Single
  Dim lngMaxH    As Long
  Dim lngCounter As Long
   
   On Error Resume Next
   
   If vlngWhole = 0 Then vlngWhole = msngMax
   If vsngPart = 0 Then vsngPart = msngValue
   
   Call SetDefaults(vPicStatus)
   
   If vlngWhole > 0 Then
      '// calculate the percentage from current value and total value
      sngPercent = vsngPart / vlngWhole
    Else
      sngPercent = 0
   End If
   
   Call GetRGB(mlngForeColor, lngRed1, lngGrn1, lngBlu1)
   If Not vblnUseOneColor Then
      Call GetRGB(mlngForeColor2, lngRed2, lngGrn2, lngBlu2)
    Else
      lngRed2 = lngRed1
      lngGrn2 = lngGrn1
      lngBlu2 = lngBlu1
   End If
   lngCounter = 10
   
   If sngPercent > 0 Then
      
      If mblnHorizontal Then
         lngMaxH = CInt(vPicStatus.ScaleWidth * vsngPart / vlngWhole) / 3
         sngRed = (lngRed2 - lngRed1) / lngMaxH
         sngGrn = (lngGrn2 - lngGrn1) / lngMaxH
         sngBlu = (lngBlu2 - lngBlu1) / lngMaxH
         
         For lngI = 0 To (vsngPart / vlngWhole) * (vPicStatus.ScaleWidth) Step 3
            lngX = lngI
            If mblnBackwards Then lngX = vPicStatus.ScaleWidth - lngI
            lngCounter = lngCounter + 3
            If lngCounter >= 8 Then
               vPicStatus.Line (lngX, 0)-(lngX + 7, vPicStatus.ScaleHeight), RGB(lngRed1 + sngRed * lngI / 3, lngGrn1 + sngGrn * lngI / 3, lngBlu1 + sngBlu * lngI / 3), BF
               
               vPicStatus.Line (lngX, 0)-(lngX + 7, vPicStatus.ScaleHeight / 3), RGB((lngRed1 + sngRed * lngI / 3) + 50, (lngGrn1 + sngGrn * lngI / 3) + 50, (lngBlu1 + sngBlu * lngI / 3) + 50), BF
               lngCounter = 0
            End If
         Next lngI
         
       Else '// Vertical
         lngMaxH = CInt(vPicStatus.ScaleHeight * vsngPart / vlngWhole) / 3
         sngRed = (lngRed2 - lngRed1) / lngMaxH
         sngGrn = (lngGrn2 - lngGrn1) / lngMaxH
         sngBlu = (lngBlu2 - lngBlu1) / lngMaxH
         
         For lngI = 2 To (vsngPart / vlngWhole) * (vPicStatus.ScaleHeight) Step 3
            lngY = vPicStatus.ScaleHeight - 7 - lngI
            If mblnBackwards Then lngY = lngI
            lngCounter = lngCounter + 3
            If lngCounter >= 8 Then
               vPicStatus.Line (2, lngY)-(vPicStatus.ScaleWidth - 2, lngY + 7), RGB(lngRed1 + sngRed * lngI / 3, lngGrn1 + sngGrn * lngI / 3, lngBlu1 + sngBlu * lngI / 3), BF
               
               vPicStatus.Line (3, lngY)-(vPicStatus.ScaleWidth / 3, lngY + 7), RGB((lngRed1 + sngRed * lngI / 3) + 50, (lngGrn1 + sngGrn * lngI / 3) + 50, (lngBlu1 + sngBlu * lngI / 3) + 50), BF
               lngCounter = 0
            End If
         Next lngI
      End If
   End If
   
   If mblnShowStatus Or mblnShowCounts Or LenB(mstrTextBefore) Or LenB(mstrTextAfter) Then
      If mblnShowCounts Then
         strCaption = CStr(vsngPart) & " of " & CStr(vlngWhole)
      End If
      If mblnShowStatus Then
         If LenB(strCaption) Then
            strCaption = strCaption & " (" & Format$(sngPercent, "0%") & ")"
          Else
            strCaption = Format$(sngPercent, "0%")
         End If
      End If
      strCaption = IIf(mstrTextBefore > vbNullString, mstrTextBefore & " ", vbNullString) & strCaption
      strCaption = strCaption & IIf(mstrTextAfter > vbNullString, " " & mstrTextAfter, vbNullString)
      
      With vPicStatus
         '// Back to the center print position and print the text
         lngWidth = .TextWidth(strCaption)
         lngHeight = .TextHeight(strCaption)
         If mblnHorizontal Then
            If .ScaleHeight < lngHeight Then .Height = lngHeight * 15
          Else
            If .ScaleWidth < lngWidth Then .Width = lngWidth * 15
         End If
         '// Now set lngX and lngY to the starting location for frmPrinting the percentage
         lngX = (.ScaleWidth / 2) - (lngWidth / 2)
         lngY = (.ScaleHeight / 2) - (lngHeight / 2)
         
         .ForeColor = mlngTextColor
         .CurrentX = lngX
         .CurrentY = lngY
         vPicStatus.Print strCaption
      End With
      
   End If
   
   '// Draw 3D boarder */
   Call DrawBoarder(vPicStatus)
   
   On Error GoTo 0
   
End Sub

Private Sub SetDefaults(ByRef vPicStatus As PictureBox)
   
   With vPicStatus
      .AutoRedraw = True
      .ScaleMode = vbPixels
      .DrawMode = vbCopyPen
      .BorderStyle = vbBSNone
      .BackColor = mlngBackColor
      .ForeColor = mlngForeColor
      .FontBold = False
      .Visible = True
      If .Height < 100 Then .Height = 100
   End With
   
End Sub

Public Property Get ShowCounts() As Boolean
   
   ShowCounts = mblnShowCounts
   
End Property

Public Property Let ShowCounts(ByVal vNewValue As Boolean)
   
   mblnShowCounts = vNewValue
   
End Property

Public Property Get ShowStatus() As Boolean
   
   ShowStatus = mblnShowStatus
   
End Property

Public Property Let ShowStatus(ByVal vNewValue As Boolean)
   
   mblnShowStatus = vNewValue
   
End Property

Public Property Get Style() As enuBarType
   
   Style = mlngStyle
   
End Property

Public Property Let Style(ByVal vNewValue As enuBarType)
   
   mlngStyle = vNewValue
   
End Property

Public Property Get TextAfter() As String
   
   TextAfter = mstrTextAfter
   
End Property

Public Property Let TextAfter(ByVal vNewValue As String)
   
   mstrTextAfter = vNewValue
   
End Property

Public Property Get TextBefore() As String
   
   TextBefore = mstrTextBefore
   
End Property

Public Property Let TextBefore(ByVal vNewValue As String)
   
   mstrTextBefore = vNewValue
   
End Property

Public Property Get TextColor() As Long
   
   TextColor = mlngTextColor
   
End Property

Public Property Let TextColor(ByVal vNewValue As Long)
   
   mlngTextColor = vNewValue
   
End Property

Public Property Get Value() As Single
   
   Value = msngValue
   
End Property

Public Property Let Value(ByVal vNewValue As Single)
   
   msngValue = vNewValue
   Select Case mlngStyle
    Case 0
      ProgressBar mpicPicBox, , , True
    Case 1
      Progress2ColorBar mpicPicBox
    Case 3
      ProgressBar mpicPicBox, , , False
    Case 4
      ProgressSteppedBar mpicPicBox, , , False
    Case Else
      ProgressSteppedBar mpicPicBox, , , True
   End Select
   
End Property

