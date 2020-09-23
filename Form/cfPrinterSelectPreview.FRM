VERSION 5.00
Begin VB.Form frmPrinterSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Setup"
   ClientHeight    =   4710
   ClientLeft      =   2970
   ClientTop       =   2220
   ClientWidth     =   6090
   Icon            =   "cfPrinterSelectPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintProperties 
      Caption         =   "Properties..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4845
      TabIndex        =   30
      Top             =   555
      Width           =   1140
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4800
      TabIndex        =   14
      Top             =   2955
      Width           =   1140
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4800
      TabIndex        =   1
      Top             =   3930
      Width           =   1140
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   3375
      TabIndex        =   11
      Top             =   2715
      Width           =   1335
      Begin VB.VScrollBar VScroll 
         Height          =   390
         Left            =   840
         Max             =   9
         Min             =   1
         TabIndex        =   13
         Top             =   375
         Value           =   1
         Width           =   240
      End
      Begin VB.TextBox txtCopies 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Text            =   "1"
         Top             =   435
         Width           =   585
      End
      Begin VB.Image imgCopies 
         Height          =   510
         Left            =   210
         Picture         =   "cfPrinterSelectPreview.frx":030A
         Top             =   960
         Width           =   900
      End
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      TabIndex        =   7
      Top             =   150
      Width           =   4785
   End
   Begin VB.TextBox txtDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   900
      Width           =   4860
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1170
      Width           =   4860
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4800
      TabIndex        =   0
      Top             =   3442
      Width           =   1140
   End
   Begin VB.Frame fraQuality 
      Caption         =   "Quality"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   225
      TabIndex        =   4
      Top             =   2715
      Width           =   3045
      Begin VB.PictureBox picfraQuality 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   30
         ScaleHeight     =   1185
         ScaleWidth      =   2910
         TabIndex        =   19
         Top             =   195
         Width           =   2910
         Begin VB.OptionButton optQuality 
            Caption         =   "Best"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   22
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton optQuality 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1065
            TabIndex        =   21
            Top             =   840
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optQuality 
            Caption         =   "Draft"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2160
            TabIndex        =   20
            Top             =   855
            Width           =   765
         End
         Begin VB.Image imgPrinterQuality 
            Height          =   480
            Index           =   2
            Left            =   225
            Picture         =   "cfPrinterSelectPreview.frx":08E6
            Top             =   210
            Width           =   390
         End
         Begin VB.Image imgPrinterQuality 
            Height          =   480
            Index           =   1
            Left            =   1290
            Picture         =   "cfPrinterSelectPreview.frx":0DD8
            Top             =   210
            Width           =   390
         End
         Begin VB.Image imgPrinterQuality 
            Height          =   480
            Index           =   0
            Left            =   2325
            Picture         =   "cfPrinterSelectPreview.frx":13CF
            Top             =   210
            Width           =   390
         End
      End
   End
   Begin VB.Frame fraDuplex 
      Caption         =   "Duplix"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   210
      TabIndex        =   3
      Top             =   1515
      Width           =   3045
      Begin VB.PictureBox picfraDuplex 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   45
         ScaleHeight     =   945
         ScaleWidth      =   2955
         TabIndex        =   23
         Top             =   195
         Width           =   2955
         Begin VB.OptionButton optDuplex 
            Caption         =   "Double Sided Book"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1050
            TabIndex        =   26
            Top             =   540
            Width           =   2100
         End
         Begin VB.OptionButton optDuplex 
            Caption         =   "Double Sided Tablet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1050
            TabIndex        =   25
            Top             =   270
            Width           =   2100
         End
         Begin VB.OptionButton optDuplex 
            Caption         =   "Single Sided"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1050
            TabIndex        =   24
            Top             =   0
            Value           =   -1  'True
            Width           =   2100
         End
         Begin VB.Image imgPrinterDuplex 
            Height          =   300
            Left            =   120
            Top             =   30
            Width           =   405
         End
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      TabIndex        =   2
      Top             =   1515
      Width           =   2580
      Begin VB.PictureBox picfraOrientation 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   150
         ScaleHeight     =   645
         ScaleWidth      =   2280
         TabIndex        =   27
         Top             =   285
         Width           =   2280
         Begin VB.OptionButton optOrien 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   930
            TabIndex        =   29
            Top             =   330
            Width           =   1590
         End
         Begin VB.OptionButton optOrien 
            Caption         =   "Portrait"
            Height          =   255
            Index           =   0
            Left            =   930
            TabIndex        =   28
            Top             =   0
            Value           =   -1  'True
            Width           =   1590
         End
         Begin VB.Image imgPrinterOrien 
            Height          =   465
            Left            =   105
            Top             =   60
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraColor 
      Height          =   510
      Left            =   225
      TabIndex        =   15
      Top             =   4110
      Width           =   3045
      Begin VB.PictureBox picFrameColor 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   240
         ScaleHeight     =   300
         ScaleWidth      =   2655
         TabIndex        =   16
         Top             =   180
         Width           =   2655
         Begin VB.OptionButton optColor 
            Caption         =   "Grayscale"
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
            Index           =   1
            Left            =   945
            TabIndex        =   18
            Top             =   15
            Width           =   1200
         End
         Begin VB.OptionButton optColor 
            Caption         =   "Color"
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
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Value           =   -1  'True
            Width           =   915
         End
      End
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   5
      Left            =   4410
      Picture         =   "cfPrinterSelectPreview.frx":19DD
      Top             =   7725
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   4
      Left            =   4425
      Picture         =   "cfPrinterSelectPreview.frx":1EB6
      Top             =   7110
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   3
      Left            =   4410
      Picture         =   "cfPrinterSelectPreview.frx":248C
      Top             =   6495
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   2
      Left            =   3570
      Picture         =   "cfPrinterSelectPreview.frx":2A4F
      Top             =   7725
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   1
      Left            =   3555
      Picture         =   "cfPrinterSelectPreview.frx":2F41
      Top             =   7110
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgQuality 
      Height          =   480
      Index           =   0
      Left            =   3540
      Picture         =   "cfPrinterSelectPreview.frx":3538
      Top             =   6525
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgDuplex 
      Height          =   345
      Index           =   5
      Left            =   1980
      Picture         =   "cfPrinterSelectPreview.frx":3B46
      Top             =   8145
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgDuplex 
      Height          =   660
      Index           =   4
      Left            =   2025
      Picture         =   "cfPrinterSelectPreview.frx":427B
      Top             =   7170
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgDuplex 
      Height          =   345
      Index           =   3
      Left            =   1995
      Picture         =   "cfPrinterSelectPreview.frx":4996
      Top             =   6570
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgDuplex 
      Height          =   465
      Index           =   2
      Left            =   540
      Picture         =   "cfPrinterSelectPreview.frx":4F1B
      Top             =   8115
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDuplex 
      Height          =   900
      Index           =   1
      Left            =   585
      Picture         =   "cfPrinterSelectPreview.frx":5677
      Top             =   7140
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgDuplex 
      Height          =   465
      Index           =   0
      Left            =   555
      Picture         =   "cfPrinterSelectPreview.frx":5DD8
      Top             =   6555
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " Printer: "
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
      Index           =   0
      Left            =   525
      TabIndex        =   10
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   9
      Top             =   900
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   8
      Top             =   1170
      Width           =   855
   End
End
Attribute VB_Name = "frmPrinterSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 1999-2003
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
Private Const C_MaxCopies          As Long = 999

Private mstrPrinterName            As String
Private mblnPrinterSetupFormLoaded As Boolean

Private Sub cboPrinter_Click()
   
  Dim xPrinter As Printer
   
   On Local Error Resume Next
   
   For Each xPrinter In Printers
      If xPrinter.DeviceName = cboPrinter.Text Then
         
         Set Printer = xPrinter
         
         txtDriver = Printer.DriverName
         mstrPrinterName = cboPrinter.Text
         txtPort = Printer.Port
         
         Printer.Orientation = gcPrint.Orientation
         optDuplex(Printer.Duplex - 1).Value = True
         optOrien(Printer.Orientation - 1).Value = True
         
         If Printer.Orientation = vbPRORPortrait Then
            optOrien(1) = False
            optOrien(0) = True
          Else
            optOrien(0) = True
            optOrien(1) = False
         End If
         
         If Printer.ColorMode = vbPRCMMonochrome Then
            optColor(1).Value = True
          Else
            optColor(0).Value = True
         End If
         
         Exit For
      End If
   Next
   
End Sub

Private Sub cmdPreview_Click()
   
   gcPrint.SendToPrinter = False
   gcPrint.Orientation = Printer.Orientation
   Call PrintPreview
   
End Sub

Private Sub cmdPrintProperties_Click()
   
  Dim cPrintDialog As clsPrintDialog
   
   Set cPrintDialog = New clsPrintDialog
   cPrintDialog.AllowOrientation = fraOrientation.Visible
   cPrintDialog.AllowPrintToFile = False
   cPrintDialog.Owner_hwnd = Me.hWnd
   cPrintDialog.ShowPrinter DLG_PRINTSETUP
   Set cPrintDialog = Nothing
   
   optOrien(Printer.Orientation - 1).Value = True
   If Printer.ColorMode = vbPRCMMonochrome Then
      optColor(1).Value = True
    Else
      optColor(0).Value = True
   End If
   
End Sub

Private Sub cmdPrint_Click()
   
   gcPrint.SendToPrinter = True
   gcPrint.Orientation = Printer.Orientation
   Call PrintPreview
   
End Sub

Private Sub cmdQuit_Click()
   
   gblnQuitCommand = True
   Me.Hide
   
End Sub

Private Sub Form_Activate()
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Initialize()
   
   Me.ZOrder
   gblnQuitCommand = True
   imgPrinterOrien.Picture = imgDuplex(0).Picture
   imgPrinterDuplex.Picture = imgDuplex(0).Picture
   
End Sub

Private Sub Form_Load()
   
  Dim xPrinter As Printer
  Dim lngIndex    As Long
   
   Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
   
   On Local Error Resume Next
   
   VScroll.Max = C_MaxCopies
   VScroll.Min = 1
   
   mstrPrinterName = GetSetting(App.Title, "Options", "Printer", "None")
   txtCopies = GetSetting(App.Title, "Options", "Copies", "1")
   
   lngIndex = -1
   For Each xPrinter In Printers
      cboPrinter.AddItem xPrinter.DeviceName
      If xPrinter.DeviceName = mstrPrinterName Then lngIndex = cboPrinter.NewIndex
      If xPrinter.DeviceName = Printer.DeviceName And lngIndex = -1 Then lngIndex = cboPrinter.NewIndex
   Next
   If lngIndex >= 0 Then cboPrinter.ListIndex = lngIndex
   
   Printer.Orientation = gcPrint.Orientation
   optOrien(Printer.Orientation - 1).Value = True
   
   mblnPrinterSetupFormLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmPrinterSetUp = Nothing
   
End Sub

Private Sub optColor_Click(Index As Integer)
   
   If Index > 0 Then Index = 3
   imgPrinterQuality(0).Picture = imgQuality(Index + 0).Picture
   imgPrinterQuality(1).Picture = imgQuality(Index + 1).Picture
   imgPrinterQuality(2).Picture = imgQuality(Index + 2).Picture
   
End Sub

Private Sub optDuplex_Click(Index As Integer)
   
   If Not mblnPrinterSetupFormLoaded Then Exit Sub
   If optOrien(1).Value Then Index = Index + 3
   imgPrinterDuplex.Picture = imgDuplex(Index).Picture
   
End Sub

Private Sub optOrien_Click(Index As Integer)
   
   On Local Error Resume Next
   
   Printer.Orientation = Index + 1
   If Err.Number Then
      optOrien(0).Value = True
      Index = 0
   End If
   
   If Index > 0 Then Index = 3
   imgPrinterOrien.Picture = imgDuplex(Index).Picture
   
   If optDuplex(0).Value Then
      imgPrinterDuplex.Picture = imgDuplex(Index).Picture
    ElseIf optDuplex(1).Value Then
      imgPrinterDuplex.Picture = imgDuplex(Index + 1).Picture
    ElseIf optDuplex(2).Value Then
      imgPrinterDuplex.Picture = imgDuplex(Index + 2).Picture
   End If
   
End Sub

Private Sub optQuality_Click(Index As Integer)
   
   On Local Error Resume Next
   Select Case Index
    Case 0
      Printer.PrintQuality = vbPRPQDraft
    Case 1
      Printer.PrintQuality = vbPRPQMedium
    Case Else
      Printer.PrintQuality = vbPRPQHigh
   End Select
   
End Sub

Private Sub PrintPreview()
   
  Dim lngI As Long
   
   On Local Error Resume Next
   For lngI = 0 To 2
      If optDuplex(lngI).Value Then
         Select Case lngI
          Case 1 '// Double Sided Tablet
            If Printer.Orientation = vbPRORPortrait Then
               Printer.Duplex = vbPRDPVertical
             Else
               Printer.Duplex = vbPRDPHorizontal
            End If
          Case 2 '// Double Sided Book
            If Printer.Orientation = vbPRORPortrait Then
               Printer.Duplex = vbPRDPHorizontal
             Else
               Printer.Duplex = vbPRDPVertical
            End If
          Case Else '// Single Sided
            Printer.Duplex = vbPRDPSimplex
         End Select
      End If
   Next lngI
   
   If optColor(1).Value Then
      Printer.ColorMode = vbPRCMMonochrome
      gcPrint.ColorMode = cmMonochrome
    Else
      Printer.ColorMode = vbPRCMColor
      gcPrint.ColorMode = cmColor
   End If
   
   Printer.Copies = Val(txtCopies)
   
   SaveSetting App.Title, "Options", "Printer", mstrPrinterName
   SaveSetting App.Title, "Options", "Copies", txtCopies
   gblnQuitCommand = False
   Unload Me
   
End Sub

Private Sub txtCopies_Change()
   
   On Local Error Resume Next
   
   If Val(txtCopies) > C_MaxCopies Then
      txtCopies = CStr(C_MaxCopies)
    ElseIf Val(txtCopies) < 1 Then
      txtCopies = "1"
   End If
   VScroll.Value = Val(txtCopies)
   
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
   
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
      KeyAscii = False
   End If
   If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
   
End Sub

Private Sub VScroll_Change()
   
   txtCopies = Abs(VScroll.Value)
   
End Sub

