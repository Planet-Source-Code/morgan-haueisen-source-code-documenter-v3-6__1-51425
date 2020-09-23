VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Project Documenter"
   ClientHeight    =   5160
   ClientLeft      =   3450
   ClientTop       =   2475
   ClientWidth     =   6720
   ForeColor       =   &H00000000&
   Icon            =   "fmrMainSingle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5160
   ScaleWidth      =   6720
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   1635
      ScaleHeight     =   3090
      ScaleWidth      =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   3390
      Begin VBProjDoc.chameleonButton cmdOptClose 
         Height          =   285
         Left            =   2460
         TabIndex        =   26
         Top             =   2760
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":030A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkExport 
         BackColor       =   &H80000014&
         Caption         =   "Export Project as HTML"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   2445
         Width           =   3090
      End
      Begin VB.CheckBox chkExport 
         BackColor       =   &H80000014&
         Caption         =   "Export Project as RTF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   23
         Top             =   2145
         Width           =   3090
      End
      Begin VB.CheckBox chkSummaryOnly 
         BackColor       =   &H80000014&
         Caption         =   "Print Summary Report ONLY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1740
         Width           =   3090
      End
      Begin VB.CheckBox chkPageBreakFiles 
         BackColor       =   &H80000014&
         Caption         =   "Page Break Between Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   135
         Value           =   1  'Checked
         Width           =   3090
      End
      Begin VB.CheckBox chkIncludeDeclarations 
         BackColor       =   &H80000014&
         Caption         =   "Include Declarations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   990
         Value           =   1  'Checked
         Width           =   3090
      End
      Begin VB.CheckBox chkIncludeCover 
         BackColor       =   &H80000014&
         Caption         =   "Include Cover Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   705
         Value           =   1  'Checked
         Width           =   3090
      End
      Begin VB.CheckBox chkIncludeSummary 
         BackColor       =   &H80000014&
         Caption         =   "Include Summary Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1410
         Value           =   1  'Checked
         Width           =   3090
      End
      Begin VB.CheckBox chkPageBreaks 
         BackColor       =   &H80000014&
         Caption         =   "Page Break Between Procedures"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   3090
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   4020
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   4020
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin MSComctlLib.TreeView trvPrintList 
      Height          =   2850
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   5027
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6660
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6720
      Begin VBProjDoc.chameleonButton cmd_add 
         Height          =   585
         Left            =   45
         TabIndex        =   15
         Top             =   15
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Open"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":0326
         PICN            =   "fmrMainSingle.frx":0342
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmd_clear 
         Height          =   585
         Left            =   645
         TabIndex        =   16
         Top             =   15
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":0808
         PICN            =   "fmrMainSingle.frx":0824
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmdOptions 
         Height          =   585
         Left            =   1605
         TabIndex        =   17
         Top             =   15
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Options"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":0CEA
         PICN            =   "fmrMainSingle.frx":0D06
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmd_begin 
         Height          =   585
         Left            =   2715
         TabIndex        =   18
         Top             =   15
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Print Project"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":11CC
         PICN            =   "fmrMainSingle.frx":11E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmdAbout 
         Height          =   585
         Left            =   5430
         TabIndex        =   19
         Top             =   15
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":16AE
         PICN            =   "fmrMainSingle.frx":16CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmd_quit 
         Height          =   585
         Left            =   6030
         TabIndex        =   20
         Top             =   15
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":1A1C
         PICN            =   "fmrMainSingle.frx":1A38
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmd_abort 
         Cancel          =   -1  'True
         Height          =   585
         Left            =   2715
         TabIndex        =   21
         Top             =   15
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "Abort Printing"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":1EFE
         PICN            =   "fmrMainSingle.frx":1F1A
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjDoc.chameleonButton cmdView 
         Height          =   585
         Left            =   4680
         TabIndex        =   25
         Top             =   15
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "View"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   255
         MPTR            =   1
         MICON           =   "fmrMainSingle.frx":23E0
         PICN            =   "fmrMainSingle.frx":23FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6720
      TabIndex        =   1
      Top             =   4635
      Width           =   6720
      Begin VB.PictureBox picProg1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   205
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   5745
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   5805
      End
      Begin VB.PictureBox picProg2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   5745
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   5805
      End
      Begin VB.Label lblPageNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   525
         Left            =   5865
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lbl_printing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   0
         TabIndex        =   2
         Top             =   100
         Width           =   5325
      End
   End
   Begin MSComctlLib.ImageList imglst1 
      Left            =   4845
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":29CA
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":2E1C
            Key             =   "Classes"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":316E
            Key             =   "Forms"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":34C0
            Key             =   "Modules"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":3812
            Key             =   "User controls"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":3B64
            Key             =   "User documents"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":3C76
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":3FC8
            Key             =   "ROOT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmrMainSingle.frx":431A
            Key             =   "SUB"
         EndProperty
      EndProperty
   End
   Begin VBProjDoc.chameleonButton cmdSelectAll 
      Height          =   345
      Left            =   135
      TabIndex        =   22
      Top             =   1065
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Collapse Tree"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "fmrMainSingle.frx":466C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgBegin 
      Height          =   240
      Index           =   1
      Left            =   7080
      Picture         =   "fmrMainSingle.frx":4688
      Top             =   3465
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBegin 
      Height          =   285
      Index           =   0
      Left            =   7095
      Picture         =   "fmrMainSingle.frx":49FD
      Top             =   3135
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblProject 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   795
      Width           =   6510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'Public VBInstance As VBIDE.VBE
'Public Connect As Connect

Private mcFile  As clsFileUtilities
Private mcProg1 As clsProgressBar
Private mcProg2 As clsProgressBar

Private mlngModuleCount  As Long
Private mlngFileNo       As Long
Private mlngFN           As Long
Private mblnSumOnly      As Boolean
Private mblnExport       As Boolean

Private mastrIndexFile() As String
Private mastrIndexSub()  As String

Private mastrReference()     As String
Private mastrReferencePath() As String
Private mastrObjects()       As String
Private mastrObjectsPath()   As String
Private mastrDLLs()          As String
Private mastrDLLsPath()      As String


Private Function CheckForStartEnd(ByVal vstrLine As String, _
                                  ByVal vstrParentName As String, _
                                  ByRef vstrProcName As String, _
                                  ByRef vblnStartSub As Boolean, _
                                  ByRef vblnEndSub As Boolean, _
                                  Optional ByVal vblnAlwaysAdd As Boolean = False) As Long
   
   On Error GoTo Err_Proc
   
   vstrLine = Trim$(vstrLine)
   
   If Not vblnStartSub Then '// Start of SUB?
      If left$(vstrLine, 12) = "Private Sub " Or _
               left$(vstrLine, 17) = "Private Function " Or _
               left$(vstrLine, 11) = "Public Sub " Or _
               left$(vstrLine, 16) = "Public Function" & " " Or _
               left$(vstrLine, 14) = "Public Static " Or _
               left$(vstrLine, 15) = "Private Static" & " " Or _
               left$(vstrLine, 11) = "Friend Sub " Or _
               left$(vstrLine, 11) = "Static Sub " Or _
               left$(vstrLine, 4) = "Sub " Or _
               left$(vstrLine, 9) = "Function " Or _
               left$(vstrLine, 21) Like "Private Property [LGS]et " Or _
               left$(vstrLine, 20) Like "Public Property" & " [LGS]et " Or _
               left$(vstrLine, 20) Like "Static Property [LGS]et " Or _
               left$(vstrLine, 20) Like "Friend Property [LGS]et " Or _
               left$(vstrLine, 13) Like "Property [LGS]et" & " " Then
         
         vstrProcName = GetProcName(vstrLine, InStr(vstrLine, "("))
         If vblnAlwaysAdd Then
            vblnStartSub = True
          Else
            vblnStartSub = modTreeView.IsNodeChecked(vstrProcName, vstrParentName)
         End If
      End If
      
    Else
      '// End of SUB?
      If left$(vstrLine, 7) = "End Sub" Or left$(vstrLine, 12) = "End Function" Or left$(vstrLine, _
               12) = "End Property" Then
         
         vblnEndSub = True
      End If
   End If
   
Exit_Proc:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "CheckForStartEnd"
   Err.Clear
   Resume Exit_Proc
   
End Function

Private Function CheckPageLen(ByVal pIndex As Long) As Boolean
   
   If mblnSumOnly Then Exit Function
   
   If gcPrint.pEndOfPage Then
      gcPrint.pNewPage
      Call PrintHeader(pIndex, True)
      CheckPageLen = True
   End If
   
End Function

Private Sub chkExport_Click(Index As Integer)
   
   If chkExport(0).Value Or chkExport(1).Value Then
      chkIncludeDeclarations.Enabled = False
      chkIncludeCover.Enabled = False
      chkPageBreaks.Enabled = False
      chkPageBreakFiles.Enabled = False
      cmd_begin.Caption = "Export Project"
      cmd_abort.Caption = "Abort Export"
      Set cmd_begin.PictureNormal = imgBegin(1).Picture
    Else
      chkIncludeDeclarations.Enabled = True
      chkIncludeCover.Enabled = True
      chkPageBreaks.Enabled = True
      chkPageBreakFiles.Enabled = True
      cmd_begin.Caption = "Print Project"
      cmd_abort.Caption = "Abort Printing"
      Set cmd_begin.PictureNormal = imgBegin(0).Picture
   End If
   
End Sub

Private Sub chkSummaryOnly_Click()
   
   If chkExport(0).Value = vbUnchecked And chkExport(1).Value = vbUnchecked Then
      chkPageBreakFiles.Enabled = Not CBool(chkSummaryOnly.Value)
      chkPageBreaks.Enabled = Not CBool(chkSummaryOnly.Value)
      chkIncludeDeclarations.Enabled = Not CBool(chkSummaryOnly.Value)
   End If
   chkIncludeSummary.Enabled = Not CBool(chkSummaryOnly.Value)
   
End Sub

Private Sub cmdAbout_Click()
   
   frmAbout.Show , Me
   
End Sub

Private Sub cmdOptClose_Click()
   
   picOptions.Visible = False
   
End Sub

Private Sub cmdOptions_Click()
   
   picOptions.Visible = True
   picOptions.ZOrder
   
End Sub

Private Sub cmdSelectAll_Click()
   
   modTreeView.CollapseAllNodes (2)
   
End Sub

Private Sub cmdView_Click()
   
  Dim strFilename As String
   
   If LenB(gstrProjectName) = 0 Then Exit Sub
   
   On Error GoTo Err_Proc
   
   With trvPrintList
      '// view source code
      If Not .SelectedItem Is Nothing And LenB(.SelectedItem.Tag) Then
         strFilename = mcFile.RetOnlyFilename(.SelectedItem.Tag)
         If LenB(Dir$(.SelectedItem.Tag)) > 0 Then
            frmStatus.Show , Me
            frmStatus.lblInfo.Caption = .SelectedItem.Tag
            frmStatus.lblInfo.Refresh
            DoEvents
            frmView.ShowFile .SelectedItem.Tag, , .SelectedItem.Text
         End If
      End If
   End With
   
Exit_Proc:
   Unload frmStatus
   
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdView_Click"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub cmd_abort_Click()
   
   gblnQuitCommand = True
   
End Sub

Private Sub cmd_add_Click()
   
   Call OpenProjectFile(True)
   
End Sub

Private Sub cmd_begin_Click()
   
  Dim lngI As Long
   
   If LenB(gstrProjectName) = 0 Then Exit Sub
   
   picOptions.Visible = False
   
   gblnQuitCommand = True
   For lngI = 6 To trvPrintList.Nodes.Count
      If trvPrintList.Nodes(lngI).Checked Then
         gblnQuitCommand = False
         Exit For
      End If
   Next lngI
   If gblnQuitCommand Then Exit Sub
   
   mblnExport = (chkExport(0).Value Or chkExport(1).Value)
   
   '// Show Print Options
   If Not mblnExport Then
      frmPrinterSetUp.Show vbModal
      If gblnQuitCommand Then
         gblnQuitCommand = False
         Exit Sub
      End If
   End If
   
   cmd_add.Visible = False
   cmd_clear.Visible = False
   cmdOptions.Visible = False
   cmd_begin.Visible = False
   cmdView.Visible = False
   cmd_quit.Visible = False
   cmdAbout.Visible = False
   trvPrintList.Enabled = False
   cmd_abort.Visible = True
   lbl_printing.Caption = vbNullString
   
   glngLinesOfCode = 0
   glngLinesOfCom = 0
   glngLinesOfBlank = 0
   mblnSumOnly = CBool(chkSummaryOnly.Value)
   lblPageNo.Visible = Not mblnSumOnly
   
   If mblnExport Then
      lblPageNo.Visible = False
      mblnSumOnly = True
      If chkSummaryOnly.Value Then
         Call MainPrint
       Else
         Call ExportRTF
         If chkIncludeSummary.Value Then Call MainPrint
      End If
      frmMsgBox.sMessage "All code was saved in " & mcFile.RetOnlyPath(gstrProjectName) _
            & "Documentation\", vbInformation + vbOkButton, , 10, , , , , Me
    Else
      Call MainPrint
   End If
   
   cmd_add.Visible = True
   cmd_clear.Visible = True
   cmdOptions.Visible = True
   cmd_abort.Visible = False
   cmd_begin.Visible = True
   cmdView.Visible = True
   cmd_quit.Visible = True
   cmdAbout.Visible = True
   trvPrintList.Enabled = True
   picProg1.Visible = False
   picProg2.Visible = False
   lblPageNo.Caption = vbNullString
   lblPageNo.Visible = False
   
End Sub

Private Sub cmd_clear_Click()
   
   lbl_printing.Caption = vbNullString
   lblProject.Caption = vbNullString
   picProg1.Visible = False
   picProg2.Visible = False
   picOptions.Visible = False
   cmd_add.Enabled = True
   gstrProjectName = vbNullString
   cmd_begin.Enabled = False
   cmdView.Enabled = False
   
   Erase gastrFileNameList
   Erase gastrFilePathList
   Erase galngFileSizeList
   Erase gastrProjectGroups
   Erase mastrObjects
   Erase mastrObjectsPath
   Erase mastrReference
   Erase mastrReferencePath
   Erase mastrDLLs
   Erase mastrDLLsPath
   Erase gaudtProjectInfo
   
   modTreeView.SetDefaultTree
   
End Sub

Private Sub cmd_quit_Click()
   
   Unload Me
   
End Sub

Private Sub ExportRTF()
   
  Dim lngI As Long
   
   On Error GoTo Err_Proc
   
   With frmStatus
      .lblTitle.Caption = "Exporting.."
      .Show , Me
      .mcProg.Max = mlngModuleCount
   End With
   DoEvents
   
   With trvPrintList
      For lngI = 1 To .Nodes.Count
         If .Nodes(lngI).Checked And LenB(.Nodes(lngI).Tag) Then
            If LenB(Dir$(.Nodes(lngI).Tag)) Then
               
               frmStatus.lblInfo.Caption = .Nodes(lngI).Tag
               frmStatus.lblInfo.Refresh
               frmStatus.mcProg.Value = lngI
               
               If chkExport(0).Value Then
                  frmView.mblnSaveRTF = True
               End If
               If chkExport(1).Value Then
                  frmView.mblnSaveHTML = True
               End If
               frmView.ShowFile .Nodes(lngI).Tag, , .Nodes(lngI).Text
               
            End If
         End If
      Next lngI
   End With
   
Exit_Proc:
   Unload frmStatus
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "ExportRTF"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub ExportSummary(ByVal vnlgIndexCount As Long, ByVal vstrIndexTitle As String)
   
  Dim strLine  As String
  Dim lngI     As Long
  Dim lngN     As Long
   
   If gaudtProjectInfo(0).VersionFileDescription > vbNullString Then
      strLine = strLine & "File Description:  " & gaudtProjectInfo(0).VersionFileDescription & vbCrLf
   End If
   
   If gbln_IsVBG Then
      strLine = strLine & "Group Project Name:  " & mcFile.RetOnlyFilename(gstrProjectName) & vbCrLf
      strLine = strLine & "   Startup Project: " & gaudtProjectInfo(0).PName & vbCrLf
    Else
      strLine = strLine & "         Project: " & gaudtProjectInfo(0).PName & vbCrLf
   End If
   strLine = strLine & " Revision Number:  " & gaudtProjectInfo(0).MajorVer & "." & gaudtProjectInfo(0).MinorVer _
         & "." & gaudtProjectInfo(0).RevisionVer & vbCrLf
   strLine = strLine & "   Revision Date:  " & gstrRevisionDate$ & vbCrLf
   strLine = strLine & "       File Name:  " & gaudtProjectInfo(0).ExeName32 & vbCrLf
   strLine = strLine & "    Company Name:  " & gaudtProjectInfo(0).VersionCompanyName & vbCrLf
   strLine = strLine & "Location: " & gstrProjectName & vbCrLf
   strLine = strLine & String$(79, "-") & vbCrLf
   
   For lngN = 1 To vnlgIndexCount
      If mcFile.RetOnlyFilename(vstrIndexTitle) <> mastrIndexFile(lngN) Then
         strLine = strLine & vbCrLf
         strLine = strLine & mastrIndexFile(lngN) & vbCrLf
         vstrIndexTitle = mastrIndexFile(lngN)
      End If
      strLine = strLine & String$(8, Chr$(32)) & mastrIndexSub(lngN) & vbCrLf
      
      With frmStatus
         .lblInfo.Caption = mastrIndexFile(lngN)
         .lblInfo.Refresh
         .mcProg.Value = lngN
      End With
      
   Next lngN
   
   strLine = strLine & vbCrLf & vbCrLf
   strLine = strLine & String$(79, "-") & vbCrLf
   
   lngI = UBound(mastrReference)
   If lngI Then
      strLine = strLine & "References:" & vbCrLf
      For lngN = 1 To lngI
         strLine = strLine & String$(4, Chr$(32)) & mastrReference(lngN) & vbCrLf
         strLine = strLine & String$(6, Chr$(32)) & "Path: " & mastrReferencePath(lngN) & vbCrLf
      Next lngN
      strLine = strLine & vbCrLf
      strLine = strLine & String$(79, "-") & vbCrLf
   End If
   
   lngI = UBound(mastrObjects)
   If lngI Then
      strLine = strLine & vbCrLf
      strLine = strLine & "Objects:" & vbCrLf
      For lngN = 1 To lngI
         strLine = strLine & String$(4, Chr(32)) & mastrObjects(lngN) & vbCrLf
         strLine = strLine & String$(6, Chr$(32)) & "Path: " & GetTypeLibInfo(mastrObjectsPath(lngN)) & vbCrLf
      Next lngN
      strLine = strLine & vbCrLf
      strLine = strLine & String$(79, "-") & vbCrLf
   End If
   
   lngI = UBound(mastrDLLs)
   If lngI Then
      strLine = strLine & vbCrLf
      strLine = strLine & "Declared DLL's:" & vbNewLine
      For lngN = 1 To lngI
         strLine = strLine & String$(4, Chr(32)) & mastrDLLs(lngN)
         strLine = strLine & String$(10, Chr(32)) & "Path: " & mastrDLLsPath(lngN) & vbNewLine
      Next lngN
      strLine = strLine & vbCrLf
      strLine = strLine & String$(79, "-") & vbCrLf
   End If
   
   strLine = strLine & "Summary:" & vbCrLf
   strLine = strLine & String$(4, Chr$(32)) & "Blank Lines" & String$(20, Chr$(32)) & _
         Format$(glngLinesOfBlank, "#,0") & vbCrLf
   strLine = strLine & String$(4, Chr$(32)) & "Lines of Code" & String$(20, Chr$(32)) & _
         Format$(glngLinesOfCode, "#,0") & vbCrLf
   strLine = strLine & String$(4, Chr$(32)) & "Lines of Comments" & String$(16, Chr$(32)) & _
         Format$(glngLinesOfCom, "#,0") & vbCrLf
   strLine = strLine & String$(4, Chr$(32)) & String$(50, "-") & vbCrLf
   strLine = strLine & String$(4, Chr$(32)) & "Total Lines (Comments & Code)" & String$(4, Chr$(32)) & _
         Format$(glngLinesOfCode + glngLinesOfCom, "#,0") & vbCrLf
   strLine = strLine & String$(79, "-") & vbCrLf
   
   If gbln_IsVBG Then
      strLine = strLine & "Group Project List:" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "Startup project" & String$(4, Chr$(32)) & mcFile.RetOnlyFilename(gastrProjectGroups(0)) & vbCrLf
      For lngI = 1 To UBound(gastrProjectGroups)
         strLine = strLine & String$(4, Chr$(32)) & "Project" & String$(4, Chr$(32)) & mcFile.RetOnlyFilename(gastrProjectGroups(lngI)) & vbCrLf
      Next lngI
      strLine = strLine & String$(79, "-") & vbCrLf
   End If
   
   For lngI = 0 To UBound(gaudtProjectInfo)
      
      If gbln_IsVBG Then strLine = strLine & mcFile.RetOnlyFilename(gastrProjectGroups(lngI)) & vbCrLf
      
      strLine = strLine & "General Information:" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "IconForm" & String$(25, Chr$(32)) & gaudtProjectInfo(lngI).IconForm & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "Startup" & String$(26, Chr$(32)) & gaudtProjectInfo(lngI).Startup & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "HelpFile" & String$(25, Chr$(32)) & gaudtProjectInfo(lngI).HelpFile & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "Title" & String$(28, Chr$(32)) & gaudtProjectInfo(lngI).Title & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "ExeName32" & String$(24, Chr$(32)) & gaudtProjectInfo(lngI).ExeName32 & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "Command32" & String$(24, Chr$(32)) & gaudtProjectInfo(lngI).Command32 & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "Name" & String$(29, Chr$(32)) & gaudtProjectInfo(lngI).PName & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "HelpContextID" & String$(20, Chr$(32)) & gaudtProjectInfo(lngI).HelpContextID & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "CompatibleMode" & String$(19, Chr$(32)) & gaudtProjectInfo(lngI).CompatibleMode & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "MajorVer" & String$(25, Chr$(32)) & gaudtProjectInfo(lngI).MajorVer & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "MinorVer" & String$(25, Chr$(32)) & gaudtProjectInfo(lngI).MinorVer & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "RevisionVer" & String$(22, Chr$(32)) & gaudtProjectInfo(lngI).RevisionVer & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "AutoIncrementVer" & String$(17, Chr$(32)) & gaudtProjectInfo(lngI).AutoIncrementVer & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "ServerSupportFiles" & String$(15, Chr$(32)) & gaudtProjectInfo(lngI).ServerSupportFiles & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionComments" & String$(18, Chr$(32)) & gaudtProjectInfo(lngI).VersionComments & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionCompanyName" & String$(15, Chr$(32)) & gaudtProjectInfo(lngI).VersionCompanyName & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionFileDescription" & String$(11, Chr$(32)) & gaudtProjectInfo(lngI).VersionFileDescription & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionLegalCopyright" & String$(12, Chr$(32)) & gaudtProjectInfo(lngI).VersionLegalCopyright & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionLegalTrademarks" & String$(11, Chr$(32)) & gaudtProjectInfo(lngI).VersionLegalTrademarks & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & "VersionProductName" & String$(15, Chr$(32)) & gaudtProjectInfo(lngI).VersionProductName & vbCrLf
      strLine = strLine & String$(79, "-") & vbCrLf
      
      strLine = strLine & "Compile Options:" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).CompilationType & " = CompilationType" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).OptimizationType & " = OptimizationType" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).FavorPentiumPro & " = FavorPentiumPro(tm)" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).CodeViewDebugInfo & " = CodeViewDebugInfo" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).NoAliasing & " = NoAliasing" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).BoundsCheck & " = BoundsCheck" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).OverflowCheck & " = OverflowCheck" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).FlPointCheck & " = FlPointCheck" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).FDIVCheck & " = FDIVCheck" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).UnroundedFP & " = UnroundedFP" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).StartMode & " = StartMode" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).Unattended & " = Unattended" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).Retained & " = Retained" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).ThreadPerObject & " = ThreadPerObject" & vbCrLf
      strLine = strLine & String$(4, Chr$(32)) & gaudtProjectInfo(lngI).MaxNumberOfThreads & " = MaxNumberOfThreads" & vbCrLf
      strLine = strLine & String$(79, "-") & vbCrLf
      
      If lngI < UBound(gaudtProjectInfo) Then
         strLine = strLine & String$(79, "-") & vbCrLf
      End If
      
   Next lngI
   
   lbl_printing.Caption = "Done Printing"
   
   If chkExport(0).Value Then
      frmView.mblnSaveRTF = True
   End If
   If chkExport(1).Value Then
      frmView.mblnSaveHTML = True
   End If
   frmView.ShowExport (strLine)
   
EndPrinting:
   
   Unload frmStatus
   gblnQuitCommand = False
   Close
   
End Sub

Private Sub Form_Initialize()
   
   '// Used for Manifest files (Win XP)
   Call ManifestWrite
   
End Sub

Private Sub Form_Load()
   
  Dim strCommandLine As String
  Dim cScreen        As clsScreenSize
   
   Set cScreen = New clsScreenSize
   cScreen.vFitScreen Me
   Set cScreen = Nothing
   
   Me.Show
   DoEvents
   
   Set mcProg1 = New clsProgressBar
   Set mcProg2 = New clsProgressBar
   
   With mcProg1
      .picBox = picProg1
      .Style = pbSolidBar
      .Max = 100
      .ShowStatus = False
      .ShowCounts = False
   End With
   With mcProg2
      .picBox = picProg2
      .Style = pbStepped2Color
      .ShowCounts = False
   End With
   
   Set gcPrint = New clsMultiPgPreview
   Set mcFile = New clsFileUtilities
   
   On Error GoTo Err_Proc
   
   With trvPrintList
      .ImageList = imglst1
      .Checkboxes = True
      .LabelEdit = tvwManual
      .Indentation = 500
      .Style = tvwTreelinesPlusMinusPictureText
      .HideSelection = False
   End With
   
   Call InitializeTreeView(trvPrintList)
   Call SetDefaultTree
   
   strCommandLine = Trim$(Command)
   If LenB(strCommandLine) > 2 Then
      If left$(strCommandLine, 1) = Chr$(34) Then
         strCommandLine = Mid$(strCommandLine, 2, Len(strCommandLine) - 2)
      End If
      If LenB(Dir$(strCommandLine)) Then
         gstrProjectName = mcFile.GetExtensionName(strCommandLine)
         If gstrProjectName = "vbg" Or gstrProjectName = "vbp" Then
            gstrProjectName = strCommandLine
            Call OpenProjectFile(False)
          Else
            gstrProjectName = vbNullString
         End If
      End If
   End If
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "Form_Load"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub Form_Resize()
   
   If Me.WindowState <> 1 Then
      trvPrintList.Height = Me.ScaleHeight - trvPrintList.top - picStatus.Height
      trvPrintList.Width = Me.ScaleWidth
      lblPageNo.left = Me.ScaleWidth - lblPageNo.Width - 25
      picProg1.Width = Me.ScaleWidth - lblPageNo.Width - 50
      picProg2.Width = Me.ScaleWidth - lblPageNo.Width - 50
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   '// Clean up
   Erase gastrFileNameList
   Erase gastrFilePathList
   Erase galngFileSizeList
   Erase gastrProjectGroups
   Erase gaudtProjectInfo
   
   Erase mastrObjects
   Erase mastrObjectsPath
   Erase mastrReference
   Erase mastrReferencePath
   Erase mastrDLLs
   Erase mastrDLLsPath
   
   Set mcFile = Nothing
   Set mcProg1 = Nothing
   Set mcProg2 = Nothing
   Set gcPrint = Nothing
   
   Call EndApp(Me)
   Set frmMain = Nothing
   
   'End
   
End Sub

Private Sub GetDllStats(ByVal LineData As String)

   Dim lngA As Integer
   Dim lngB As Integer
   Dim strTemp As String

   lngA = InStr(1, LineData, """") ' Find start of filename
   lngB = InStr(lngA + 1, LineData, """") ' Find end of filename

   strTemp = Mid$(LineData, lngA + 1, lngB - lngA - 1)
   If LCase$(right$(strTemp, 4)) <> ".dll" Then strTemp = strTemp & ".dll"
   
   Call AddNode("Dependencies", "Declared DLLs", "Declared DLLs", , "FOLDER", False, True)
   If AddNode("Declared DLLs", strTemp, strTemp, , "FOLDER", False) Then
      
      lngA = UBound(mastrDLLs) + 1
      ReDim Preserve mastrDLLs(lngA)
      ReDim Preserve mastrDLLsPath(lngA)
      
      mastrDLLs(lngA) = strTemp
      If LenB(Dir$((Environ("windir") & "\System\" & strTemp))) Then
         mastrDLLsPath(lngA) = Environ("windir") & "\System\" & strTemp
      ElseIf LenB(Dir$((Environ("windir") & "\System32\" & strTemp))) Then
         mastrDLLsPath(lngA) = Environ("windir") & "\System32\" & strTemp
      Else
         mastrDLLsPath(lngA) = C_UNKOWN
      End If
   
   End If
   
End Sub

Private Function GetGroupID(ByVal vstrTest As String) As Long
   
  Dim lngI As Long
   
   If gbln_IsVBG Then
      For lngI = 0 To UBound(gastrProjectGroups)
         If mcFile.RetOnlyFilename(gastrProjectGroups(lngI)) = mcFile.RetOnlyPath(vstrTest, False) Then Exit For
      Next lngI
      GetGroupID = lngI
    Else
      GetGroupID = 0
   End If
   
End Function

Private Sub GetGroupProjects()
   
  Dim lngFN   As Long
  Dim lngI    As Long
  Dim strLine As String
   
   lngFN = FreeFile
   Open gstrProjectName For Input As #lngFN
   Do
      Line Input #lngFN, strLine
      
      Select Case left$(strLine, 7)
       Case "Startup"
         gastrProjectGroups(lngI) = mcFile.RetOnlyPath(gstrProjectName) & Mid$(strLine, 16)
         lngI = UBound(gastrProjectGroups) + 1
         ReDim Preserve gastrProjectGroups(lngI)
         
       Case "Project"
         gastrProjectGroups(lngI) = mcFile.RetOnlyPath(gstrProjectName) & Mid$(strLine, 9)
         lngI = UBound(gastrProjectGroups) + 1
         ReDim Preserve gastrProjectGroups(lngI)
         
      End Select
      
   Loop Until EOF(lngFN)
   Close #lngFN
   
   lngI = lngI - 1
   ReDim Preserve gastrProjectGroups(lngI)
   
End Sub

Public Function GetProcName(ByVal vstrLine As String, _
                            ByVal vblnStartPoint As Long) As String
   
  Dim lngStartBr        As Long
  Dim strPropertyPrefix As String
   
   '// Purpose: parse the procedure name from the initializing line
   
   On Error GoTo Err_Proc
   
   '// If a property then add the prefix (let or get)
   If left$(vstrLine, 21) Like "Private Property [LGS]et " Then
      strPropertyPrefix = Mid$(vstrLine, 18, 4)
      
    ElseIf left$(vstrLine, 20) Like "Public Property [LGS]et " Or left$(vstrLine, 20) Like "Static" & _
            " Property [LGS]et " Or left$(vstrLine, 20) Like "Friend Property [LGS]et " Then
      
      strPropertyPrefix = Mid$(vstrLine, 17, 4)
      
    ElseIf left$(vstrLine, 13) Like "Property [LGS]et " Then
      strPropertyPrefix = Mid$(vstrLine, 10, 4)
    Else
      strPropertyPrefix = vbNullString
   End If
   
   vblnStartPoint = vblnStartPoint - 1
   
   Do Until Len(GetProcName) > 0
      lngStartBr = InStrRev(vstrLine, " ", vblnStartPoint, vbTextCompare)
      
      If vblnStartPoint - lngStartBr > 0 Then
         GetProcName = strPropertyPrefix & Trim$(Mid$(vstrLine, lngStartBr + 1, (vblnStartPoint - lngStartBr)))
      End If
      
      If LenB(GetProcName) = 0 Then
         vblnStartPoint = lngStartBr - 1
      End If
   Loop
   
Exit_Proc:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "GetProcName"
   Err.Clear
   Resume Exit_Proc
   
End Function

Private Function GetProjectFiles() As Boolean
   
  Dim lngI As Long
   
   On Error GoTo Err_Proc
   
   ReDim gastrFileNameList(0) As String
   ReDim gastrFilePathList(0) As String
   ReDim galngFileSizeList(0) As Long
   
   '// scan the files list
   With trvPrintList
      For lngI = 1 To .Nodes.Count
         If LenB(.Nodes(lngI).Tag) > 0 And (.Nodes(lngI).Checked Or mblnSumOnly) Then
            ReDim Preserve gastrFileNameList(UBound(gastrFileNameList) + 1) As String
            ReDim Preserve gastrFilePathList(UBound(gastrFilePathList) + 1) As String
            ReDim Preserve galngFileSizeList(UBound(galngFileSizeList) + 1) As Long
            gastrFilePathList(UBound(gastrFilePathList)) = .Nodes(lngI).Tag
            gastrFileNameList(UBound(gastrFileNameList)) = .Nodes(lngI).Text
            galngFileSizeList(UBound(galngFileSizeList)) = FileLen(.Nodes(lngI).Tag)
         End If
         If lngI = mlngModuleCount Then Exit For
      Next lngI
   End With
   
   DoEvents
   If UBound(gastrFileNameList) = 0 Then
      frmMsgBox.sMessage "No Files Selected To Print", vbOkButton + vbExclamation
      GetProjectFiles = False
    Else
      GetProjectFiles = True
   End If
   
Exit_Proc:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "GetProjectFiles"
   Err.Clear
   Resume Exit_Proc
   
End Function

Private Sub LoadDependencies()
   
  Dim lngI As Long
  Dim lngN As Long
   
   On Error Resume Next
   
   lngI = UBound(mastrReference)
   If lngI And Err.Number = 0 Then
      Call AddNode("Dependencies", "References", "References", , "FOLDER", False, True)
      For lngN = 1 To lngI
         Call AddNode("References", mastrReference(lngN), mastrReference(lngN), , "FOLDER", False)
      Next lngN
   End If
   
   Err.Clear
   lngI = UBound(mastrObjects)
   If lngI And Err.Number = 0 Then
      Call AddNode("Dependencies", "Objects", "Objects", , "FOLDER", False, True)
      For lngN = 1 To lngI
         Call AddNode("Objects", mastrObjects(lngN), mastrObjects(lngN), , "FOLDER", False)
      Next lngN
   End If
   
End Sub

Private Function LoadProjectProcedures(ByVal vstrFilePath As String, _
                                       ByVal vstrModuleName As String) As Boolean
   
  Dim lngSFN            As Long    '// source file number
  Dim strLine           As String  '// source read line
  Dim strProcName       As String  '// current procedure name
  Dim blnStartSub       As Boolean '// recognize sub start
  Dim blnEndSub         As Boolean '// recognize end of sub or function
  Dim blnFlag           As Boolean
  Dim blnFlag2          As Boolean
   
   On Error GoTo Err_Proc
   
   '// init algorithm flags
   blnStartSub = False
   blnEndSub = False
   strProcName = vbNullString
   'vstrModuleName = mcFile.RetOnlyFilename(vstrFilePath)
   
   frmStatus.lblInfo.Caption = vstrFilePath
   frmStatus.lblInfo.Refresh
   
   '// Open Source File for reading
   lngSFN = FreeFile
   Open vstrFilePath For Input As #lngSFN
   
   '// Remove file header
   blnFlag = False
   blnFlag2 = False
   Do
      Line Input #lngSFN, strLine
      
      blnFlag2 = CBool(InStr(Trim$(strLine), "Attribute VB_") = 1)
      If Not blnFlag Or blnFlag2 Then
         blnFlag = blnFlag2
       Else
         Exit Do
      End If
      'Call UpdateBar(sFilename)
   Loop Until EOF(lngSFN)
   If EOF(lngSFN) Then GoTo Exit_Proc
   
   '// Main scanning loop
   Do Until EOF(lngSFN)
      
      If blnFlag Then
         blnFlag = False
       Else
         '// read next line
         Line Input #lngSFN, strLine
      End If
      
      strLine = Trim$(strLine)
      If LenB(strLine) > 0 Then
         If left$(strLine, 1) = "'" Or left$(strLine, 4) = "Rem " Then
            glngLinesOfCom = glngLinesOfCom + 1
          ElseIf right$(strLine, 1) <> "_" And LenB(strLine) > 0 Then
            glngLinesOfCode = glngLinesOfCode + 1
         End If
      End If
      
      If InStrB(strLine, " Declare ") And InStrB(strLine, " Lib ") Then
         Call GetDllStats(strLine)
      End If
      
      '//  Check if it is the Begining or Ending of a procedure
      Call CheckForStartEnd(strLine, vstrModuleName, strProcName, blnStartSub, blnEndSub, True)
      
      If blnEndSub Then
         Call AddNode(vstrModuleName, GetNextKey(), strProcName, , "SUB")
         '//  Clear variables:
         blnStartSub = False
         blnEndSub = False
         glngLinesOfCode = glngLinesOfCode - 2
         strProcName = vbNullString
      End If
      
   Loop
   
Exit_Proc:
   On Error Resume Next
   Close #lngSFN
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "LoadProjectProcedures"
   Err.Clear
   Resume Exit_Proc
   
End Function

Private Sub MainPrint()
   
  Const C_sngTab1 As Single = 0.7
  Const C_sngTab2 As Single = 2.9
  Const C_sngTab3 As Single = 3
   
  Dim strLine        As String
  Dim lngI           As Long
  Dim lngN           As Long
  Dim blnFirstLook   As Boolean
  Dim blnSecondLook  As Boolean
  Dim blnEndOfSub    As Boolean
  Dim blnNewPage     As Boolean
  Dim lngIndexCount  As Long
  Dim strIndexTitle  As String
  Dim blnFlag        As Boolean
  Dim blnFlag2       As Boolean
  Dim blnOkToPrint   As Boolean
  Dim strProcName    As String
  Dim strFilename    As String
  Dim alngIndexPg()  As Long
  Dim lngIndex       As Long
   
   On Error GoTo Err_Proc
   
StartAgain:
   
   lngIndexCount = 0
   ReDim mastrIndexFile(0) As String
   ReDim mastrIndexSub(0) As String
   ReDim alngIndexPg(0) As Long
   
   If Not GetProjectFiles Then Exit Sub
   
   gcPrint.pStartDoc
   
   If chkIncludeCover.Value And Not mblnExport Then
      '// Print Cover page
      gcPrint.pFontName
      gcPrint.DrawWidth = 4
      gcPrint.pBox 0.5, 0.5, gcPrint.GetPaperWidth - 1, gcPrint.GetPaperHeight - 1, RGB(0, 0, 150)
      gcPrint.DrawWidth = 1
      gcPrint.FontSize = 24
      gcPrint.FontBold = True
      gcPrint.CurrentY = 3
      If gaudtProjectInfo(0).VersionFileDescription > vbNullString Then
         gcPrint.pCenter gcPrint.GetStripQuotes(gaudtProjectInfo(0).VersionFileDescription)
      End If
      gcPrint.pPrint
      
      If gbln_IsVBG Then
         gcPrint.FontSize = 12
         gcPrint.pCenter "Group Project Name:  " & mcFile.RetOnlyFilename(gstrProjectName)
         gcPrint.FontSize = 24
         gcPrint.pCenter "Startup Project: " & gcPrint.GetStripQuotes(gaudtProjectInfo(0).PName)
       Else
         gcPrint.pCenter "Project: " & gcPrint.GetStripQuotes(gaudtProjectInfo(0).PName)
      End If
      
      gcPrint.FontSize = 14
      gcPrint.pPrint
      gcPrint.pCenter "Revision Number:  " & gaudtProjectInfo(0).MajorVer & "." & gaudtProjectInfo(0).MinorVer _
            & "." & gaudtProjectInfo(0).RevisionVer
      gcPrint.pCenter "Revision Date:  " & gstrRevisionDate$
      gcPrint.pCenter "File Name:  " & gcPrint.GetStripQuotes(gaudtProjectInfo(0).ExeName32)
      gcPrint.pPrint
      gcPrint.pCenter gcPrint.GetStripQuotes(gaudtProjectInfo(0).VersionCompanyName)
      
      gcPrint.FontBold = False
      gcPrint.FontSize = 12
      
      If (mlngModuleCount - 6) > UBound(gastrFileNameList) Then
         gcPrint.pPrint
         gcPrint.pCenter "Partial Listing (" & CStr(UBound(gastrFileNameList)) & " of " & _
               CStr(mlngModuleCount - 6) & ")"
      End If
      
      gcPrint.CurrentY = gcPrint.GetPaperHeight - 1.5
      gcPrint.ForeColor = RGB(0, 0, 150)
      gcPrint.pPrintedDate True
      gcPrint.ForeColor = vbBlack
      
      If Printer.Duplex > 1 Then gcPrint.pNewPage
      gcPrint.pNewPage
   End If
   
   '// Begin printing document
   gcPrint.FontName = "Courier New"
   gcPrint.FontSize = 9
   blnNewPage = True
   mcProg2.Max = UBound(gastrFileNameList)
   
   For mlngFileNo = 1 To UBound(gastrFilePathList)
      
      mcProg2.Value = mlngFileNo
      strFilename = mcFile.RetOnlyFilename(gastrFilePathList(mlngFileNo))
      lngIndex = GetGroupID(gastrFileNameList(mlngFileNo))
      
      mlngFN = FreeFile
      Open gastrFilePathList(mlngFileNo) For Input As #mlngFN
      
      '// Remove file header
      blnFlag = False
      blnFlag2 = False
      Do
         Line Input #mlngFN, strLine
         
         blnFlag2 = CBool(InStr(Trim$(strLine), "Attribute VB_") = 1)
         If Not blnFlag Or blnFlag2 Then
            blnFlag = blnFlag2
          Else
            Exit Do
         End If
         Call UpdateBar(strFilename)
      Loop Until EOF(mlngFN)
      If EOF(mlngFN) Then GoTo EmptyFile
      
      blnFirstLook = True
      If chkIncludeDeclarations.Value Or mblnExport Then
         blnSecondLook = True
         blnOkToPrint = True
       Else
         blnSecondLook = False
         blnOkToPrint = False
      End If
      
      Do
         If blnFlag Then
            blnFlag = False
          Else
            '// read next line
            Line Input #mlngFN, strLine
         End If
         
         Call UpdateBar(strFilename)
         
         If blnFirstLook Then
            Call PrintHeader(lngIndex, blnNewPage)
            
            If blnOkToPrint Then
               glngLinesOfCode = glngLinesOfCode - 1
               If gcPrint.ColorMode = cmMonochrome Then gcPrint.FontBold = True
               gcPrint.ForeColor = QBColor(1) 'vbblue
               PrintLine "(Declarations)", 0.4
               If Not mblnSumOnly Then gcPrint.pPrint
               gcPrint.FontBold = False
               gcPrint.ForeColor = vbBlack
               lngIndexCount = lngIndexCount + 1
               ReDim Preserve mastrIndexFile(lngIndexCount) As String
               ReDim Preserve mastrIndexSub(lngIndexCount) As String
               ReDim Preserve alngIndexPg(lngIndexCount) As Long
               mastrIndexFile(lngIndexCount) = gastrFileNameList(mlngFileNo)
               mastrIndexSub(lngIndexCount) = "(Declarations)"
               alngIndexPg(lngIndexCount) = gcPrint.GetPage
            End If
            blnFirstLook = False
         End If
         
         If left$(LTrim$(strLine), 12) = "Private Sub " Or _
                  left$(LTrim$(strLine), 17) = "Private Function " Or _
                  left$(LTrim$(strLine), 11) = "Public Sub " Or _
                  left$(LTrim$(strLine), 16) = "Public Function " Or _
                  left$(LTrim$(strLine), 14) = "Public Static" & " " Or _
                  left$(LTrim$(strLine), 15) = "Private Static " Or _
                  left$(LTrim$(strLine), 11) = "Friend Sub " Or _
                  left$(LTrim$(strLine), 11) = "Static Sub " Or _
                  left$(LTrim$(strLine), 4) = "Sub " Or _
                  left$(LTrim$(strLine), 9) = "Function " Or _
                  left$(LTrim$(strLine), 21) Like "Private Property [LGS]et " Or _
                  left$(LTrim$(strLine), 20) Like "Public Property [LGS]et " Or _
                  left$(LTrim$(strLine), 20) Like "Static Property" & " [LGS]et " Or _
                  left$(LTrim$(strLine), 20) Like "Friend Property [LGS]et " Or _
                  left$(LTrim$(strLine), 13) Like "Property [LGS]et " Then
            
            blnEndOfSub = False
            strProcName = GetProcName(strLine, InStr(strLine, "("))
            blnOkToPrint = modTreeView.IsNodeChecked(strProcName, gastrFileNameList(mlngFileNo)) Or mblnExport Or mblnSumOnly
            strLine = Trim$(strLine)
            
            If blnOkToPrint Then
               
               If Me.chkPageBreaks.Value Then
                  If Not mblnSumOnly Then gcPrint.pNewPage
                  Call PrintHeader(lngIndex, True)
                ElseIf Not blnFirstLook And blnSecondLook Then
                  If Not mblnSumOnly Then gcPrint.pPrint
                  If Not mblnSumOnly Then gcPrint.pLine
                  blnSecondLook = False
               End If
               
               glngLinesOfCode = glngLinesOfCode - 1
               If gcPrint.ColorMode = cmMonochrome Then gcPrint.FontBold = True
               gcPrint.ForeColor = QBColor(1) 'vbblue
               lngIndexCount = lngIndexCount + 1
               ReDim Preserve mastrIndexFile(lngIndexCount) As String
               ReDim Preserve mastrIndexSub(lngIndexCount) As String
               ReDim Preserve alngIndexPg(lngIndexCount) As Long
               mastrIndexFile(lngIndexCount) = strFilename
               lngN = InStr(strLine, "(")
               
               If left$(strLine, 2) = "Pr" Then '// Private */
                  mastrIndexSub(lngIndexCount) = Mid$(strLine, 9, lngN - 9)
                ElseIf left$(strLine, 2) = "Pu" Then '// Public */
                  mastrIndexSub(lngIndexCount) = Mid$(strLine, 8, lngN - 8)
                Else
                  mastrIndexSub(lngIndexCount) = Mid$(strLine, 1, lngN - 1) '// Not Defined Public _
                        */
               End If
               alngIndexPg(lngIndexCount) = gcPrint.GetPage
            End If
         End If
         
         If left$(LTrim$(strLine), 7) = "End Sub" Or left$(LTrim$(strLine), 12) = "End" & _
                  " Function" Or left$(LTrim$(strLine), 12) = "End Property" Then
            
            blnEndOfSub = True
            If blnOkToPrint Then
               glngLinesOfCode = glngLinesOfCode - 1
               If gcPrint.ColorMode = cmMonochrome Then gcPrint.FontBold = True
               gcPrint.ForeColor = QBColor(1) 'vbBlue
               PrintLine Trim$(strLine)
               If Not mblnSumOnly Then gcPrint.pLine
               If Not mblnSumOnly Then gcPrint.pPrint
            End If
            blnOkToPrint = False
          Else
            If (blnOkToPrint And Not (blnEndOfSub And Trim$(strLine) = vbNullString)) Then PrintLine strLine
         End If
         
         blnNewPage = CheckPageLen(lngIndex)
         
         If right$(Trim$(strLine), 1) <> "_" Then
            gcPrint.FontBold = False
            gcPrint.ForeColor = vbBlack
         End If
         
         DoEvents
         
         If gblnQuitCommand Then
            gcPrint.pCancled
            lbl_printing.Caption = "Printing Aborted"
            Exit Do
         End If
         
      Loop Until EOF(mlngFN)
         
EmptyFile:
         
      Close #mlngFN
      blnFirstLook = False
      
      If Not mblnSumOnly Then
         If mlngFileNo <> UBound(gastrFileNameList) And Me.chkPageBreakFiles.Value Then gcPrint.pNewPage
      End If
      
      If gblnQuitCommand Then Exit For
      
   Next mlngFileNo
   
   If Not gblnQuitCommand And chkIncludeSummary.Value Then
      
      gcPrint.FontName = "Arial"
      
      With frmStatus
         .mcProg.Max = UBound(mastrIndexFile)
         .lblTitle.Caption = "Writing Summary.. Please wait."
         .Show , Me
      End With
      DoEvents
      
      If Not mblnSumOnly Then gcPrint.pNewPage
      mblnSumOnly = False
      
      If mblnExport Then
         Call ExportSummary(lngIndexCount, strIndexTitle)
         GoTo Exit_Proc
      End If
      
      '// Print Index pages */
      lngN = gcPrint.GetPage
      If Printer.Duplex > 1 Then
         If lngN \ 2 = lngN / 2 Then gcPrint.pNewPage
      End If
      
      Call PrintIndexHeader(True)
      For lngN = 1 To lngIndexCount
         If mcFile.RetOnlyFilename(strIndexTitle) <> mastrIndexFile(lngN) Then
            gcPrint.pPrint
            gcPrint.pPrint mastrIndexFile(lngN), 0.5, True
            gcPrint.pDots 3.4
            strIndexTitle = mastrIndexFile(lngN)
         End If
         gcPrint.pPrint mastrIndexSub(lngN), 3.5, True
         
         With frmStatus
            .lblInfo.Caption = mastrIndexFile(lngN)
            .lblInfo.Refresh
            .mcProg.Value = lngN
         End With
         
         If Not CBool(chkSummaryOnly.Value) Then
            gcPrint.pDots 7
            gcPrint.pRightJust alngIndexPg(lngN), 7.5
          Else
            gcPrint.pPrint
         End If
         
         If gcPrint.pEndOfPage Then
            gcPrint.pNewPage
            Call PrintIndexHeader(True)
            strIndexTitle = vbNullString
         End If
      Next lngN
      
      '// Print Summary page */
      On Error Resume Next
      gcPrint.pNewPage
      Call PrintIndexHeader(False)
      
      lngI = UBound(mastrReference)
      If lngI Then
         gcPrint.FontBold = True
         gcPrint.pPrint "References:", 0.5
         gcPrint.FontBold = False
         For lngN = 1 To lngI
            gcPrint.pPrint mastrReference(lngN), C_sngTab1
            gcPrint.FontSize = 8
            gcPrint.pPrint "Path: " & mastrReferencePath(lngN), C_sngTab1 + 0.25
            gcPrint.FontSize = 9
         Next lngN
         gcPrint.pLine
      End If
      
      lngI = UBound(mastrObjects)
      If lngI Then
         gcPrint.pPrint
         gcPrint.FontBold = True
         gcPrint.pPrint "Objects:", 0.5
         gcPrint.FontBold = False
         For lngN = 1 To lngI
            gcPrint.pPrint mastrObjects(lngN), C_sngTab1
            gcPrint.FontSize = 8
            gcPrint.pPrint "Path: " & GetTypeLibInfo(mastrObjectsPath(lngN)), C_sngTab1 + 0.25
            gcPrint.FontSize = 9
         Next lngN
         gcPrint.pLine
      End If
      
      lngI = UBound(mastrDLLs)
      If lngI Then
         gcPrint.pPrint
         gcPrint.FontBold = True
         gcPrint.pPrint "Declared DLL's:", 0.5
         gcPrint.FontBold = False
         For lngN = 1 To lngI
            gcPrint.pPrint mastrDLLs(lngN), C_sngTab1, True
            gcPrint.FontSize = 8
            gcPrint.pDots C_sngTab2 - 0.05
            gcPrint.pPrint "Path: " & mastrDLLsPath(lngN), C_sngTab2
            gcPrint.FontSize = 9
         Next lngN
         gcPrint.pLine
      End If
      
      gcPrint.pPrint
      gcPrint.FontBold = True
      gcPrint.pPrint "Summary:", 0.5
      gcPrint.FontBold = False
      gcPrint.pPrint "Blank Lines", C_sngTab1, True
      gcPrint.pDots 3.5
      gcPrint.pRightJust Format$(glngLinesOfBlank, "#,0"), 4.5
      gcPrint.pPrint "Lines of Code", C_sngTab1, True
      gcPrint.pDots 3.5
      gcPrint.pRightJust Format$(glngLinesOfCode, "#,0"), 4.5
      gcPrint.pPrint "Lines of Comments", C_sngTab1, True
      gcPrint.pDots 3.5
      gcPrint.pRightJust Format$(glngLinesOfCom, "#,0"), 4.5
      gcPrint.pLine C_sngTab1, 4.625
      gcPrint.pPrint "Total Lines (Comments & Code)", C_sngTab1, True
      gcPrint.pDots 3.5
      gcPrint.pRightJust Format$(glngLinesOfCode + glngLinesOfCom, "#,0"), 4.5
      gcPrint.pLine
      
      If gbln_IsVBG Then
         gcPrint.FontBold = True
         gcPrint.pPrint "Group Project List:", 0.5
         gcPrint.FontBold = False
         gcPrint.pPrint "Startup project", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint mcFile.RetOnlyFilename(gastrProjectGroups(0)), C_sngTab3
         For lngI = 1 To UBound(gastrProjectGroups)
            gcPrint.pPrint "Project", C_sngTab1, True
            gcPrint.pDots C_sngTab2
            gcPrint.pPrint mcFile.RetOnlyFilename(gastrProjectGroups(lngI)), C_sngTab3
         Next lngI
         gcPrint.pLine
      End If
      
      gcPrint.pNewPage
      Call PrintIndexHeader(False)
      
      For lngI = 0 To UBound(gaudtProjectInfo)
         
         gcPrint.FontBold = True
         If gbln_IsVBG Then gcPrint.pPrint mcFile.RetOnlyFilename(gastrProjectGroups(lngI)), 0.5
         gcPrint.pPrint "General Information:", 0.5
         gcPrint.FontBold = False
         gcPrint.pPrint "IconForm", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).IconForm), C_sngTab3
         gcPrint.pPrint "Startup", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).Startup), C_sngTab3
         gcPrint.pPrint "HelpFile", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).HelpFile), C_sngTab3
         gcPrint.pPrint "Title", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).Title), C_sngTab3
         gcPrint.pPrint "ExeName32", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).ExeName32), C_sngTab3
         gcPrint.pPrint "Command32", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).Command32), C_sngTab3
         gcPrint.pPrint "Name", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).PName), C_sngTab3
         gcPrint.pPrint "HelpContextID", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).HelpContextID), C_sngTab3
         gcPrint.pPrint "CompatibleMode", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).CompatibleMode), C_sngTab3
         gcPrint.pPrint "MajorVer", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gaudtProjectInfo(lngI).MajorVer, C_sngTab3
         gcPrint.pPrint "MinorVer", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gaudtProjectInfo(lngI).MinorVer, C_sngTab3
         gcPrint.pPrint "RevisionVer", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gaudtProjectInfo(lngI).RevisionVer, C_sngTab3
         gcPrint.pPrint "AutoIncrementVer", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gaudtProjectInfo(lngI).AutoIncrementVer, C_sngTab3
         gcPrint.pPrint "ServerSupportFiles", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).ServerSupportFiles), C_sngTab3
         gcPrint.pPrint "VersionComments", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionComments), C_sngTab3
         gcPrint.pPrint "VersionCompanyName", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionCompanyName), C_sngTab3
         gcPrint.pPrint "VersionFileDescription", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionFileDescription), C_sngTab3
         gcPrint.pPrint "VersionLegalCopyright", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionLegalCopyright), C_sngTab3
         gcPrint.pPrint "VersionLegalTrademarks", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionLegalTrademarks), C_sngTab3
         gcPrint.pPrint "VersionProductName", C_sngTab1, True
         gcPrint.pDots C_sngTab2
         gcPrint.pPrint gcPrint.GetStripQuotes(gaudtProjectInfo(lngI).VersionProductName), C_sngTab3
         gcPrint.pLine
         
         gcPrint.FontBold = True
         gcPrint.pPrint "Compile Options:", 0.5
         gcPrint.FontBold = False
         gcPrint.pPrint gaudtProjectInfo(lngI).CompilationType & " = CompilationType", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).OptimizationType & " = OptimizationType", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).FavorPentiumPro & " = FavorPentiumPro(tm)", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).CodeViewDebugInfo & " = CodeViewDebugInfo", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).NoAliasing & " = NoAliasing", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).BoundsCheck & " = BoundsCheck", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).OverflowCheck & " = OverflowCheck", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).FlPointCheck & " = FlPointCheck", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).FDIVCheck & " = FDIVCheck", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).UnroundedFP & " = UnroundedFP", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).StartMode & " = StartMode", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).Unattended & " = Unattended", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).Retained & " = Retained", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).ThreadPerObject & " = ThreadPerObject", C_sngTab1
         gcPrint.pPrint gaudtProjectInfo(lngI).MaxNumberOfThreads & " = MaxNumberOfThreads", C_sngTab1
         
         If lngI < UBound(gaudtProjectInfo) Then
            gcPrint.pNewPage
            Call PrintIndexHeader(False)
         End If
         
      Next lngI
      
      lbl_printing.Caption = "Done Printing - " & gcPrint.GetPage & " pages"
   End If
   
EndPrinting:
   
   Unload frmStatus
   gblnQuitCommand = False
   Close
   
   gcPrint.pEndDoc
   If gcPrint.SendToPrinter = True Then GoTo StartAgain
   
   Exit Sub
   
Exit_Proc:
   Close
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description & vbCrLf & gastrFileNameList(mlngFileNo), "frmMain", "MainPrint"
   Err.Clear
   Resume Next
   
End Sub

Private Sub OpenProjectFile(ByVal vblnShowDialog As Boolean)
   
  Dim strFilename  As String
  Dim lngI         As Long
   
   On Error GoTo Err_Proc
   
   If vblnShowDialog Then
      mcFile.VBGetOpenFileName strFilename, , , , , , "VB Project (*.vbp;*.vbg)|*.vbp;*.vbg", , , "Open" & _
            " a VB Project", "vbp", Me.hWnd
      If strFilename = vbNullString Then Exit Sub
      gstrProjectName = strFilename
   End If
   
   Screen.MousePointer = vbHourglass
   
   With frmStatus
      .lblTitle.Caption = "Opening Project File:"
      .lblTitle.Refresh
      .Show , Me
   End With
   Me.Enabled = False
   DoEvents
   
   glngLinesOfCom = 0
   glngLinesOfCode = 0
   glngLinesOfBlank = 0
   
   '// vb project - add all relevant files
   ReDim gastrProjectGroups(0)
   ReDim gaudtProjectInfo(0)
   
   If mcFile.GetExtensionName(gstrProjectName) = "vbg" Then
      gbln_IsVBG = True
      Call GetGroupProjects
      ReDim gaudtProjectInfo(UBound(gastrProjectGroups))
      For lngI = 0 To UBound(gastrProjectGroups)
         With frmStatus
            .lblInfo.Caption = gastrProjectGroups(lngI)
            .lblInfo.Refresh
         End With
         
         Call RefreshProject(lngI)
      Next lngI
    Else
      gbln_IsVBG = False
      gastrProjectGroups(0) = gstrProjectName
      Call RefreshProject(0)
   End If
   
   '// Save file count (node count)
   mlngModuleCount = trvPrintList.Nodes.Count
   
   If Me.trvPrintList.Nodes.Count > 0 Then
      With frmStatus
         .lblTitle.Caption = "Processing Project File:"
         .lblTitle.Refresh
         .mcProg.Max = mlngModuleCount
      End With
      '// scan the files list and add procedures
      With trvPrintList
         For lngI = 1 To .Nodes.Count
            If LenB(.Nodes(lngI).Tag) > 0 Then
               '// load project procedures into TreeView
               Call LoadProjectProcedures(.Nodes(lngI).Tag, .Nodes(lngI).Text)
            End If
            If lngI > mlngModuleCount Then Exit For
            frmStatus.mcProg.Value = lngI
         Next lngI
      End With
      
      cmd_begin.Enabled = True
   End If
   
   Call LoadDependencies
   
   Call AddNode("Dependencies", "Counts", "Counts", , "FOLDER", False, True)
   Call AddNode("Counts", "LinesOfComments", CStr(glngLinesOfCom) & "=Lines Of Comments", , "FOLDER", False)
   Call AddNode("Counts", "LinesOfCode", CStr(glngLinesOfCode) & "=Lines Of Code", , "FOLDER", False)
   
   cmd_add.Enabled = False
   Unload frmStatus
   Me.Enabled = True
   cmd_begin.Enabled = True
   cmdView.Enabled = True
   Screen.MousePointer = vbDefault
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmd_add_Click"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub PrintHeader(ByVal vlngIndex As Long, Optional _
                        ByVal vblnNewPage As Boolean = True)
   
  Dim blnFontBoldT   As Boolean
  Dim blnFontItalicT As Boolean
  Dim lngColor       As Long
  Dim strFontname    As String
   
   If mblnSumOnly Then Exit Sub
   
   On Error GoTo Err_Proc
   
   blnFontBoldT = gcPrint.FontBold
   blnFontItalicT = gcPrint.FontItalic
   lngColor = gcPrint.ForeColor
   strFontname = gcPrint.FontName
   
   gcPrint.FontName = "Arial"
   
   gcPrint.FontSize = 9
   gcPrint.FontItalic = True
   gcPrint.FontBold = False
   gcPrint.ForeColor = vbBlack
   
   If Not vblnNewPage Then gcPrint.pLine
   
   gcPrint.pPrint "Project: " & gaudtProjectInfo(vlngIndex).PName, 0.25, True
   gcPrint.pCenter "Revision Date: " & gstrRevisionDate$
   
   gcPrint.pPrint "File: ", 0.25, True
   If gcPrint.ColorMode = cmColor Then
      gcPrint.ForeColor = vbRed 'QBColor(4)
    Else
      gcPrint.FontBold = True
   End If
   gcPrint.pPrint mcFile.RetOnlyFilename(gastrFileNameList(mlngFileNo)), , True
   
   gcPrint.FontBold = False
   gcPrint.ForeColor = vbBlack
   
   gcPrint.pPrint "   " & Format(galngFileSizeList(mlngFileNo) / 1000, "#,0.0") & " KB ", , True
   gcPrint.pPrint "   Modified: " & Format$(FileDateTime(gastrFilePathList(mlngFileNo)), "m/d/yy" & _
         "  h:mm:ss ampm"), , True
   gcPrint.pRightTab "Pg." & gcPrint.GetPage, 0.25
   
   gcPrint.FontName = strFontname
   gcPrint.FontBold = blnFontBoldT
   gcPrint.FontItalic = blnFontItalicT
   
   gcPrint.pDoubleLine
   gcPrint.ForeColor = lngColor
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "PrintHeader"
   Err.Clear
   Resume Next
   
End Sub

Private Sub PrintIndexHeader(ByVal vblnIncludeSubHeading As Boolean)
   
  Dim strFontname As String
  Dim lngFontSize As Long
   
   strFontname = gcPrint.FontName
   lngFontSize = gcPrint.FontSize
   
   gcPrint.FontName = "Arial"
   gcPrint.FontSize = 9
   gcPrint.FontItalic = True
   gcPrint.ForeColor = vbBlack
   
   gcPrint.FontItalic = True
   gcPrint.pCenter "Revision Date: " & gstrRevisionDate$, True
   gcPrint.pPrint "Project: " & gaudtProjectInfo(0).PName, 0.25, True
   gcPrint.pRightTab "Index", 0.25
   gcPrint.FontItalic = False
   gcPrint.pDoubleLine
   
   gcPrint.FontName = strFontname
   gcPrint.FontSize = lngFontSize
   
   If vblnIncludeSubHeading Then
      gcPrint.FontUnderline = True
      gcPrint.FontBold = True
      gcPrint.pPrint "Module", 0.5, True
      gcPrint.pPrint "Procedure", 3.5, True
      If Not CBool(chkSummaryOnly.Value) Then
         gcPrint.pRightJust "Page#", 7.5
       Else
         gcPrint.pPrint
      End If
      gcPrint.FontUnderline = False
      gcPrint.FontBold = False
   End If
   
End Sub

Private Sub PrintLine(ByVal vstrLine As String, _
                      Optional vsngLeftMarg As Single = 0.5, _
                      Optional vblnSameLine As Boolean = False)
   
  Dim lngTxtLen     As Long
  Dim lngSpaceCount As Long
  Dim blnRemOff     As Boolean
  Dim lngN          As Long
  Dim lngColor      As Long
   
   On Error GoTo Err_Proc
   
   lngTxtLen = Len(vstrLine$)
   lngSpaceCount = Len(LTrim$(vstrLine$))
   lngSpaceCount = lngTxtLen - lngSpaceCount
   Printer.FillStyle = vbFSTransparent
   lngColor = gcPrint.ForeColor
   
   If LenB(Trim$(vstrLine$)) > 0 Then
      
      If left$(LTrim$(vstrLine$), 1) = "'" Or left$(LTrim$(vstrLine$), 4) = "Rem " Then
         glngLinesOfCom = glngLinesOfCom + 1
       ElseIf right$(Trim$(vstrLine$), 1) <> "_" Then
         glngLinesOfCode = glngLinesOfCode + 1
      End If
      
      If mblnSumOnly Then Exit Sub
      
      gcPrint.CurrentX = vsngLeftMarg
      For lngN = 1 To lngTxtLen
         If Mid$(vstrLine$, lngN, 1) = Chr$(34) Then blnRemOff = Not blnRemOff   '// Look for "
         '// Look for ' or Rem
         If (Mid$(vstrLine$, lngN, 1) = "'" Or Mid$(vstrLine$, lngN, 4) = "Rem ") And Not blnRemOff Then
            gcPrint.ForeColor = QBColor(2) 'vbGreen
            gcPrint.FontItalic = True
            gcPrint.pFontName
            gcPrint.FontSize = 9
         End If
         
         gcPrint.pPrint Mid$(vstrLine$, lngN, 1), , True
         If gcPrint.CurrentX > gcPrint.GetPaperWidth - 1 And lngN < lngTxtLen Then
            '// Break Line as space
            If Mid$(vstrLine$, lngN, 1) = " " Or gcPrint.CurrentX > gcPrint.GetPaperWidth - 0.3 Then
               gcPrint.pPrint "¬"
               gcPrint.CurrentX = vsngLeftMarg
               gcPrint.pPrint String$(lngSpaceCount + 1, 32), , True
            End If
         End If
      Next lngN
      
      gcPrint.FontItalic = False
      gcPrint.ForeColor = lngColor
      gcPrint.FontName = "Courier New"
      gcPrint.FontSize = 9
      
    Else
      glngLinesOfBlank = glngLinesOfBlank + 1
   End If
   
   If Not vblnSameLine And Not mblnSumOnly Then gcPrint.pPrint
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "PrintLine"
   Err.Clear
   Resume Next
   
End Sub

Private Sub RefreshProject(ByVal vlngIndex As Long)
   
  Dim lngObj               As Long
  Dim lngFN                As Long
  Dim lngX                 As Long
  Dim lngI                 As Long
  Dim strLine              As String
  Dim strTemp              As String
  Dim lngRef               As Long
  Dim strExt               As String
  Dim strSPath             As String
  Dim strObjectName        As String
  Dim strFilename          As String
  Dim blnAutoIncrementVer  As Boolean
  Dim blnGetRefPath        As Long
   
   On Error GoTo Err_Proc
   
   If vlngIndex = 0 Then
      ReDim mastrObjects(0) As String
      ReDim mastrObjectsPath(0) As String
      ReDim mastrReference(0) As String
      ReDim mastrReferencePath(0) As String
      ReDim mastrDLLs(0) As String
      ReDim mastrDLLsPath(0) As String
      
      '// Set Project Name
      trvPrintList.Nodes(1).Text = mcFile.RetOnlyFilename(gstrProjectName)
      trvPrintList.Nodes(1).Checked = True
      
      gstrRevisionDate$ = FileDateTime(gstrProjectName)
   End If
   
   lngFN = FreeFile
   
   '// ensures backslash is exists
   If gbln_IsVBG Then
      strSPath = mcFile.RetOnlyPath(gastrProjectGroups(vlngIndex))
      Open gastrProjectGroups(vlngIndex) For Input As #lngFN
    Else
      strSPath = mcFile.RetOnlyPath(gstrProjectName)
      Open gstrProjectName For Input As #lngFN
   End If
   
   Do
      Line Input #lngFN, strLine
      
      Select Case left$(strLine, 5)
       Case "Prope" 'PropertyPage=
         strObjectName = Mid$(strLine, 14)
         GoSub AddFileToList
       
       Case "UserC" 'UserControl=
         strObjectName = Mid$(strLine, 13)
         GoSub AddFileToList
       
       Case "UserD" 'UserDocument=
         strObjectName = Mid$(strLine, 14)
         GoSub AddFileToList
       
       Case "Desig" 'Designer=
         strObjectName = Mid$(strLine, 10)
         GoSub AddFileToList
       
       Case "Refer" 'Reference=
         lngI = InStrRev(strLine, "#")
         If lngI Then
            lngRef = lngRef + 1
            ReDim Preserve mastrReference(lngRef) As String
            ReDim Preserve mastrReferencePath(lngRef) As String
            
            lngX = InStrRev(strLine, "#", lngI - 1)
            mastrReferencePath(lngRef) = Mid$(strLine, 14, lngX - 14)
            mastrReferencePath(lngRef) = Replace(mastrReferencePath(lngRef), "#", "\")
            
            strTemp = GetTypeLibInfo(mastrReferencePath(lngRef), True)
            If InStrB(strTemp, C_UNKOWN) Then
               lngX = InStrRev(strLine, "#")
               mastrReference(lngRef) = Mid$(strLine, lngX + 1)
               lngX = InStr(mastrReferencePath(lngRef), "\")
               strTemp = Mid$(mastrReferencePath(lngRef), lngX + 1)
               lngX = InStrRev(strTemp, "\")
               strTemp = left$(strTemp, lngX - 1)
               mastrReference(lngRef) = mastrReference(lngRef) & "  (v" & strTemp & ")"
            Else
               mastrReference(lngRef) = strTemp
            End If
            
            strTemp = GetTypeLibInfo(mastrReferencePath(lngRef))
            If InStrB(strTemp, C_UNKOWN) Then
               lngX = InStrRev(strLine, "#", lngI - 1)
               mastrReferencePath(lngRef) = Mid$(strLine, lngX + 1, lngI - lngX - 1)
               If left$(mastrReferencePath(lngRef), 1) = "." Then
                  mastrReferencePath(lngRef) = FixedPath(mastrReferencePath(lngRef))
               End If
            Else
               mastrReferencePath(lngRef) = strTemp
            End If
            
         End If
         
       Case "Objec" 'Object=
         lngX = InStr(strLine, ";")
         If lngX Then
            lngObj = lngObj + 1
            ReDim Preserve mastrObjects(lngObj) As String
            ReDim Preserve mastrObjectsPath(lngObj) As String
            
            mastrObjectsPath(lngObj) = Mid$(strLine, 8, lngX - 8)
            mastrObjectsPath(lngObj) = Replace(mastrObjectsPath(lngObj), "#", "\")
            
            strTemp = GetTypeLibInfo(mastrObjectsPath(lngObj), True)
            If InStrB(strTemp, C_UNKOWN) Then
               mastrObjects(lngObj) = Mid$(strLine, lngX + 2)
               lngX = InStr(mastrObjectsPath(lngObj), "\")
               strTemp = Mid$(mastrObjectsPath(lngObj), lngX + 1)
               lngX = InStrRev(strTemp, "\")
               strTemp = left$(strTemp, lngX - 1)
               mastrObjects(lngObj) = mastrObjects(lngObj) & "  (v" & strTemp & ")"
            Else
               mastrObjects(lngObj) = strTemp
            End If
         End If
         
       Case "Class" 'Class=xxxxx; xxxxx.cls
         lngX = InStr(strLine, " ")
         If lngX Then
            strObjectName = Mid$(strLine, lngX + 1)
            GoSub AddFileToList
          Else
            strObjectName = Mid$(strLine, 7)
            GoSub AddFileToList
         End If
       
       Case "Form=" 'Form=
         strObjectName = Mid$(strLine, 6)
         GoSub AddFileToList
       
       Case "Modul" 'Module=xxxxx; xxxxx.bas
         lngX = InStr(strLine, " ")
         If lngX Then
            strObjectName = Mid$(strLine, lngX + 1)
            GoSub AddFileToList
          Else
            strObjectName = Mid$(strLine, 8)
            GoSub AddFileToList
         End If
       
       Case "IconF" 'IconForm=
         gaudtProjectInfo(vlngIndex).IconForm = Mid$(strLine, 10)
       
       Case "Title" 'Title=
         gaudtProjectInfo(vlngIndex).Title = Mid$(strLine, 7)
       
       Case "ExeNa" 'ExeName32=
         gaudtProjectInfo(vlngIndex).ExeName32 = Mid$(strLine, 11)
       
       Case "Comma" 'Command32=
         gaudtProjectInfo(vlngIndex).Command32 = Mid$(strLine, 11)
       
       Case "Major" 'MajorVer=
         gaudtProjectInfo(vlngIndex).MajorVer = Mid$(strLine, 10)
       
       Case "Minor" 'MinorVer=
         gaudtProjectInfo(vlngIndex).MinorVer = Mid$(strLine, 10)
       
       Case "Revis" 'RevisionVer=
         gaudtProjectInfo(vlngIndex).RevisionVer = Mid$(strLine, 13)
         If Val(gaudtProjectInfo(vlngIndex).RevisionVer) > 0 Then gaudtProjectInfo(vlngIndex).RevisionVer = CStr(Val(gaudtProjectInfo(vlngIndex).RevisionVer) - 1)
       
       Case "Name=" 'Name=
         gaudtProjectInfo(vlngIndex).PName = Mid$(strLine, 6)
       
       Case "MaxNu" 'MaxNumberOfThreads
         gaudtProjectInfo(vlngIndex).MaxNumberOfThreads = CStr(Val(right$(strLine, 1)))
       
       Case "Threa" 'ThreadPerObject
         gaudtProjectInfo(vlngIndex).ThreadPerObject = CStr(Val(right$(strLine, 1)))
       
       Case "Retai" 'Retained
         gaudtProjectInfo(vlngIndex).Retained = CStr(Val(right$(strLine, 1)))
       
       Case "Unatt" 'Unattended
         gaudtProjectInfo(vlngIndex).Unattended = CStr(Val(right$(strLine, 1)))
       
       Case "Unrou" 'UnroundedFP
         gaudtProjectInfo(vlngIndex).UnroundedFP = CStr(Val(right$(strLine, 1)))
       
       Case "FDIVC" 'FDIVCheck
         gaudtProjectInfo(vlngIndex).FDIVCheck = CStr(Val(right$(strLine, 1)))
       
       Case "FlPoi" 'FlPointCheck
         gaudtProjectInfo(vlngIndex).FlPointCheck = CStr(Val(right$(strLine, 1)))
       
       Case "Overf" 'OverflowCheck
         gaudtProjectInfo(vlngIndex).OverflowCheck = CStr(Val(right$(strLine, 1)))
       
       Case "Bound" 'BoundsCheck
         gaudtProjectInfo(vlngIndex).BoundsCheck = CStr(Val(right$(strLine, 1)))
       
       Case "NoAli" 'NoAliasing
         gaudtProjectInfo(vlngIndex).NoAliasing = CStr(Val(right$(strLine, 1)))
       
       Case "CodeV" 'CodeViewDebugInfo
         gaudtProjectInfo(vlngIndex).CodeViewDebugInfo = CStr(Val(right$(strLine, 1)))
       
       Case "Favor" 'FavorPentiumPro(tm)
         gaudtProjectInfo(vlngIndex).FavorPentiumPro = CStr(Val(right$(strLine, 1)))
       
       Case "Optim" 'OptimizationType
         gaudtProjectInfo(vlngIndex).OptimizationType = CStr(Val(right$(strLine, 1)))
       
       Case "Serve" 'ServerSupportFiles
         gaudtProjectInfo(vlngIndex).ServerSupportFiles = IIf(right$(strLine, 1) = "1", "Yes", "No")
       
       Case "AutoI" 'AutoIncrementVer=
         blnAutoIncrementVer = Not CBool(right$(strLine, 1))
         gaudtProjectInfo(vlngIndex).AutoIncrementVer = CStr(blnAutoIncrementVer)
       
       Case Else
         If left$(strLine, 22) = "VersionFileDescription" Then
            gaudtProjectInfo(vlngIndex).VersionFileDescription = Mid$(strLine, 24)
          ElseIf left$(strLine, 15) = "VersionComments" Then
            gaudtProjectInfo(vlngIndex).VersionComments = Mid$(strLine, 17)
          ElseIf left$(strLine, 18) = "VersionCompanyName" Then
            gaudtProjectInfo(vlngIndex).VersionCompanyName = Mid$(strLine, 20)
          ElseIf left$(strLine, 21) = "VersionLegalCopyright" Then
            gaudtProjectInfo(vlngIndex).VersionLegalCopyright = Mid$(strLine, 23)
          ElseIf left$(strLine, 22) = "VersionLegalTrademarks" Then
            gaudtProjectInfo(vlngIndex).VersionLegalTrademarks = Mid$(strLine, 24)
          ElseIf left$(strLine, 18) = "VersionProductName" Then
            gaudtProjectInfo(vlngIndex).VersionProductName = Mid$(strLine, 20)
          ElseIf left$(strLine, 6) = "StartM" Then
            gaudtProjectInfo(vlngIndex).StartMode = right$(strLine, 1)
          ElseIf left$(strLine, 6) = "Startu" Then
            gaudtProjectInfo(vlngIndex).Startup = Mid$(strLine, 9)
          ElseIf left$(strLine, 14) = "CompatibleMode" Then
            gaudtProjectInfo(vlngIndex).CompatibleMode = Mid$(strLine, 16)
          ElseIf left$(strLine, 15) = "CompilationType" Then
            gaudtProjectInfo(vlngIndex).CompilationType = right$(strLine, 1)
          ElseIf left$(strLine, 5) = "HelpC" Then
            gaudtProjectInfo(vlngIndex).HelpContextID = Mid$(strLine, 15)
          ElseIf left$(strLine, 5) = "HelpF" Then
            gaudtProjectInfo(vlngIndex).HelpFile = Mid$(strLine, 10)
         End If
      End Select
   Loop Until EOF(lngFN)
   
   If blnAutoIncrementVer Then
      If Val(gaudtProjectInfo(vlngIndex).RevisionVer) > 0 Then
         gaudtProjectInfo(vlngIndex).RevisionVer = Format$(Val(gaudtProjectInfo(vlngIndex).RevisionVer) - 1, "00")
      End If
   End If
   
   If vlngIndex = 0 Then
      lblProject.Caption = gaudtProjectInfo(vlngIndex).PName & "  v" & gaudtProjectInfo(vlngIndex).MajorVer & _
            "." & gaudtProjectInfo(vlngIndex).MinorVer & "." & gaudtProjectInfo(vlngIndex).RevisionVer
      
      If gbln_IsVBG Then lblProject.Caption = lblProject.Caption & " (Group Project)"
   End If
   
Exit_Proc:
   On Error Resume Next
   Close #lngFN
   Me.Show
   
   Exit Sub
   
AddFileToList:
   '// Purpose: Adds new file to the TreeView list
   If Dir$(strSPath & strObjectName) > vbNullString Then
      With frmStatus
         .lblInfo.Caption = strObjectName
         .lblInfo.Refresh
      End With
      
      strFilename = strSPath & strObjectName
      
      strExt = mcFile.GetExtensionName(strFilename)
      Select Case strExt
       Case "frm"
         strExt = "Forms"
       Case "cls"
         strExt = "Classes"
       Case "bas"
         strExt = "Modules"
       Case "dsr"
         strExt = "Designer"
       Case "ctl"
         strExt = "User controls"
       Case "dob"
         strExt = "User documents"
      End Select
      
      If LenB(strFilename) > 0 Then
         If modTreeView.FindTag(strFilename) = 0 Then
            If gbln_IsVBG Then
               modTreeView.AddNode strExt, _
                     mcFile.RetOnlyFilename(gastrProjectGroups(vlngIndex)) & "\" & mcFile.RetOnlyFilename(strFilename), _
                     mcFile.RetOnlyFilename(gastrProjectGroups(vlngIndex)) & "\" & mcFile.RetOnlyFilename(strFilename), _
                     strFilename, strExt
             Else
               modTreeView.AddNode strExt, _
                     mcFile.RetOnlyFilename(strFilename), _
                     mcFile.RetOnlyFilename(strFilename), _
                     strFilename, strExt
            End If
         End If
      End If
      
    Else
      frmMsgBox.SMessageModal strObjectName & " file is missing.", vbInformation + vbOkButton
   End If
   Return
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "RefreshProject"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Private Sub trvPrintList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Call DisableCheck
   
End Sub

Private Sub trvPrintList_NodeCheck(ByVal Node As MSComctlLib.Node)
   
   Call NodeCheckedEvent(Node)
   
End Sub

Private Sub trvPrintList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, _
                                     Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   '// File or directory?
   If (GetAttr(Data.Files(1)) And vbDirectory) = vbDirectory Then
      '// do nothing
    Else
      If mcFile.GetExtensionName(Data.Files(1)) = "vbp" Then
         gstrProjectName = Data.Files(1)
         Call SetDefaultTree
         Call OpenProjectFile(False)
      End If
   End If
   
End Sub

Private Sub UpdateBar(ByVal vstrFileName As String)
   
  Dim lngPresentLoc As Long
   
   '// Update Amount Of File Printed
   
   On Error GoTo Err_Proc
   
   lngPresentLoc = Loc(mlngFN) * 128  '// Get location within file.
   lngPresentLoc = Int((lngPresentLoc / galngFileSizeList(mlngFileNo)) * 100)
   If lngPresentLoc > 100 Then lngPresentLoc = 100
   
   mcProg1.TextAfter = vstrFileName
   mcProg1.Value = lngPresentLoc
   
   lblPageNo = "Page" & vbCrLf & CStr(gcPrint.GetPage)
   lblPageNo.Refresh
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "UpdateBar"
   Err.Clear
   Resume Exit_Proc
   
End Sub

