VERSION 5.00
Begin VB.Form frmMultiPgPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   ClientHeight    =   6885
   ClientLeft      =   1665
   ClientTop       =   1725
   ClientWidth     =   8610
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMultiPgPreview_BMPJPG4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6885
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGoto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   3090
      TabIndex        =   25
      Top             =   5700
      Visible         =   0   'False
      Width           =   3120
      Begin VB.TextBox txtGoto 
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
         Left            =   1305
         TabIndex        =   26
         Text            =   "1"
         Top             =   105
         Width           =   1590
      End
      Begin VBProjectDocumenter.chameleonButton cmdGotoOK 
         Height          =   255
         Left            =   2130
         TabIndex        =   35
         Top             =   465
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Ok"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000F&
         Height          =   750
         Left            =   15
         Top             =   15
         Width           =   3045
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Goto Page#"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   465
         Left            =   120
         TabIndex        =   27
         Top             =   165
         Width           =   1080
      End
   End
   Begin VB.PictureBox picFullPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3840
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   2340
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picPrintPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3885
      ScaleHeight     =   435
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   8055
      ScaleHeight     =   6615
      ScaleWidth      =   555
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   555
      Begin VB.VScrollBar VScroll1 
         Height          =   1260
         Left            =   150
         Max             =   100
         Min             =   -20
         TabIndex        =   23
         Top             =   3495
         Width           =   270
      End
      Begin VBProjectDocumenter.chameleonButton cmd_quit 
         Cancel          =   -1  'True
         Height          =   630
         Left            =   15
         TabIndex        =   28
         Top             =   15
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1111
         BTYPE           =   9
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
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":0028
         PICN            =   "frmMultiPgPreview_BMPJPG4.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton cmd_print 
         Height          =   630
         Left            =   15
         TabIndex        =   29
         Top             =   660
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1111
         BTYPE           =   9
         TX              =   "Print"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":073E
         PICN            =   "frmMultiPgPreview_BMPJPG4.frx":075A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton cmdFullPage 
         Height          =   510
         Left            =   15
         TabIndex        =   30
         Top             =   1305
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "Fit"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":0E64
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton Command1 
         Height          =   345
         Index           =   0
         Left            =   45
         TabIndex        =   31
         Top             =   2595
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16507600
         BCOLO           =   16507600
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":0E80
         PICN            =   "frmMultiPgPreview_BMPJPG4.frx":0E9C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton Command1 
         Height          =   345
         Index           =   1
         Left            =   270
         TabIndex        =   32
         Top             =   2595
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   ""
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":1226
         PICN            =   "frmMultiPgPreview_BMPJPG4.frx":1242
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton cmdGoTo 
         Height          =   510
         Left            =   45
         TabIndex        =   33
         Top             =   2970
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "&Goto"
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
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":15CD
         PICN            =   "frmMultiPgPreview_BMPJPG4.frx":15E9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   45
         TabIndex        =   24
         Top             =   1830
         Width           =   465
      End
   End
   Begin VB.PictureBox picHScroll 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6615
      Visible         =   0   'False
      Width           =   8610
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   0
         Max             =   100
         TabIndex        =   1
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.PictureBox picPrintOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H000000FF&
      Height          =   2640
      Left            =   555
      ScaleHeight     =   2610
      ScaleWidth      =   3150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   3180
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1695
         TabIndex        =   5
         Text            =   "1"
         Top             =   1350
         Width           =   420
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2475
         TabIndex        =   6
         Text            =   "1"
         Top             =   1350
         Width           =   420
      End
      Begin VBProjectDocumenter.chameleonButton cmdPrint 
         Height          =   360
         Left            =   2265
         TabIndex        =   34
         Top             =   2070
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Ok"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":1FFB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy all pages to a Folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   4
         Left            =   585
         TabIndex        =   16
         Top             =   420
         Width           =   2250
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   4
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2017
         Top             =   390
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":20B4
         Top             =   705
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy page to clipboard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   0
         Left            =   585
         TabIndex        =   2
         Top             =   735
         Width           =   2250
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Current Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   1
         Left            =   585
         TabIndex        =   3
         Top             =   1065
         Width           =   1965
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2151
         Top             =   1035
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   2
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":21EE
         Top             =   1335
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   3
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":228B
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   3
         Left            =   585
         TabIndex        =   7
         Top             =   1695
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2175
         TabIndex        =   13
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label lblPrintingPg 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Height          =   210
         Left            =   255
         TabIndex        =   12
         Top             =   2250
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Left            =   135
         TabIndex        =   10
         Top             =   30
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000F&
         Height          =   2535
         Left            =   30
         Top             =   30
         Width           =   3090
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Pages"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   2
         Left            =   585
         TabIndex        =   4
         Top             =   1365
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3765
   End
   Begin VB.PictureBox picGetFolder 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4440
      Left            =   1245
      ScaleHeight     =   4410
      ScaleWidth      =   6375
      TabIndex        =   17
      Top             =   615
      Visible         =   0   'False
      Width           =   6405
      Begin VB.DriveListBox Drive1 
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
         Left            =   1530
         TabIndex        =   21
         Top             =   45
         Width           =   3930
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3465
         Left            =   30
         TabIndex        =   20
         Top             =   450
         Width           =   6315
      End
      Begin VB.CommandButton cmdNewFolder 
         Height          =   345
         Left            =   5955
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "New Folder"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUpOne 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2676
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Back Up"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VBProjectDocumenter.chameleonButton cmdQuit 
         Height          =   375
         Left            =   3270
         TabIndex        =   36
         Top             =   3975
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":2928
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBProjectDocumenter.chameleonButton cmdOpen 
         Height          =   375
         Left            =   4755
         TabIndex        =   37
         Top             =   3975
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ok"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMultiPgPreview_BMPJPG4.frx":2944
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Select a Directory: "
         Height          =   195
         Left            =   75
         TabIndex        =   22
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.Image imgFit 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2960
      Top             =   5370
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFit 
      Height          =   240
      Index           =   1
      Left            =   360
      Picture         =   "frmMultiPgPreview_BMPJPG4.frx":2EEA
      Top             =   5385
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   0
      Picture         =   "frmMultiPgPreview_BMPJPG4.frx":3474
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   555
      Picture         =   "frmMultiPgPreview_BMPJPG4.frx":3521
      Top             =   4875
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 1999-2003
'//************************************/
Option Explicit

Public PageNumber As Integer

Private ViewPage  As Integer
Private TempDir   As String
Private OptionV   As Integer
Private FitToPage As Boolean

Private Type PanState
   X As Long
   Y As Long
End Type
Private PanSet As PanState

Private Declare Function StretchBlt Lib "gdi32" ( _
   ByVal hdc As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, _
   ByVal xSrc As Long, _
   ByVal ySrc As Long, _
   ByVal nSrcWidth As Long, _
   ByVal nSrcHeight As Long, _
   ByVal dwRop As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
   ByRef lpVersionInformation As OSVersionInfo) As Long

Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" ( _
   ByVal lpPathName As String, _
   ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Type SECURITY_ATTRIBUTES
   nLength              As Long
   lpSecurityDescriptor As Long
   bInheritHandle       As Long
End Type

Private Type OSVersionInfo
   OSVSize       As Long
   dwVerMajor    As Long
   dwVerMinor    As Long
   dwBuildNumber As Long
   PlatformID    As Long
   szCSDVersion  As String * 128
End Type
Private UseStretchBit As Boolean


Private Sub cmdFullPage_Click()
   
  Dim xmin   As Single
  Dim ymin   As Single
  Dim wid    As Single
  Dim hgt    As Single
  Dim aspect As Single
   
   On Error GoTo Err_Proc
   
   '// If already here then restore original
   If cmdFullPage.Value = 0 Then
      Picture1.Visible = True
      Picture1.SetFocus
      picFullPage.Visible = False
      Set cmdFullPage.PictureNormal = imgFit(0).Picture
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   DoEvents
   Set cmdFullPage.PictureNormal = imgFit(1).Picture
   
   '// Clear any picture and set the size and loaction
   Set picFullPage.Picture = Nothing
   If Not picHScroll.Visible Then
      picFullPage.Height = Me.Height - 100
      picFullPage.Width = picFullPage.Height * 0.773
      picFullPage.Move ((Me.Width - Picture2.Width) - picFullPage.Width) \ 2, 0
   Else
      picFullPage.top = 50
      picFullPage.left = 50
      picFullPage.Width = Me.Width - Picture2.Width - 100
      picFullPage.Height = picFullPage.Width * 0.773
   End If
   
   '// Get the scale values
   aspect = Picture1.ScaleHeight / Picture1.ScaleWidth
   wid = picFullPage.ScaleWidth
   hgt = picFullPage.ScaleHeight
   
   '// MaintainRatio
   If hgt / wid > aspect Then
      hgt = aspect * wid
      xmin = picFullPage.ScaleLeft
      ymin = (picFullPage.ScaleHeight - hgt) / 2
   Else
      wid = hgt / aspect
      xmin = (picFullPage.ScaleWidth - wid) / 2
      ymin = picFullPage.ScaleTop
   End If
   
   If UseStretchBit Then '// NT platform
      StretchBlt picFullPage.hdc, xmin, ymin, wid, hgt, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
   Else
      picFullPage.PaintPicture Picture1.Picture, xmin, ymin, wid, hgt, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
   End If
   
   picGoto.Visible = False
   Picture1.Visible = False
   picFullPage.Visible = True
   picFullPage.SetFocus
   
   Screen.MousePointer = vbDefault
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "cmdFullPage_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub cmdGoToOK_Click()
   
  Dim NewPageNo As Integer
   
   On Local Error Resume Next
   
   txtGoto.SetFocus
   NewPageNo = Val(txtGoto)
   If NewPageNo = 0 Then Exit Sub
   
   NewPageNo = NewPageNo - 1
   If NewPageNo > PageNumber Then NewPageNo = PageNumber
   ViewPage = NewPageNo
   
   Picture1.SetFocus
   Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
   
   picPrintOptions.Visible = False
   picGetFolder.Visible = False
   picGoto.Visible = False
   
   VScroll1.Value = 0
   HScroll1.Value = 0
   Call DisplayPages
   
End Sub

Private Sub cmdGoTo_Click()
   
   On Error GoTo Err_Proc
   
   picGoto.top = cmdGoTo.top
   picGoto.left = Me.Width - (Picture2.Width + picGoto.Width + 50)
   picGoto.Visible = True
   picGoto.ZOrder
   txtGoto = CStr(ViewPage + 1)
   txtGoto.SelStart = 0
   txtGoto.SelLength = Len(txtGoto)
   txtGoto.SetFocus
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "cmdGoTo_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub cmdNewFolder_Click()
   
  Dim NewFolderName As String
  Dim Security      As SECURITY_ATTRIBUTES
   
   On Error GoTo Err_Proc
   
   NewFolderName = InputBox("Enter Folder Name", , "New Folder")
   NewFolderName = Trim$(NewFolderName)
   If NewFolderName > vbNullString Then
      CreateDirectory Dir1.Path & "\" & NewFolderName, Security
      NewFolderName = Dir1.Path & "\" & NewFolderName
      Dir1.Refresh
      Dir1.Path = NewFolderName
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "cmdNewFolder_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub cmdOpen_Click()
   
  Dim FolderName As String
  Dim ReportTitle As String
  Dim i As Integer
  Dim cJPG As clsJPEG
   
   FolderName = Dir1.Path & "\"
   picGetFolder.Visible = False
   
   picPrintOptions.Visible = True
   picPrintOptions.Enabled = False
   lblPrintingPg.Visible = True
   cmdPrint.Visible = False
   
   On Local Error GoTo CopyError:
   
   DoEvents
   ReportTitle = Trim$(gcPrint.ReportTitle)
   If ReportTitle = vbNullString Or InStr(ReportTitle, "\") Then
      ReportTitle = "PPview"
   End If
   
   Set cJPG = New clsJPEG
   
   For i = 0 To PageNumber
      lblPrintingPg.Caption = "Copying page " & i + 1
      lblPrintingPg.Refresh
      
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(i) & ".bmp")
      With cJPG
         '// (/ 15 - 4) Change to pixels and remove border
         .SampleHDC Picture1.hdc, Picture1.Width / 15 - 4, Picture1.Height / 15 - 4
         .SaveFile FolderName & ReportTitle & CStr(i + 1) & ".jpg"
      End With
      
   Next
   
   '// Restore normal view
   picPrintOptions.Enabled = True
   cmdPrint.Visible = True
   picPrintOptions.Visible = False
   lblPrintingPg.Visible = False
   
   Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
   Set cJPG = Nothing
   
   Exit Sub
   
CopyError:
   If Err.Number = 76 Then
      ReportTitle = "PPview"
      Resume
   End If
   
End Sub

Private Sub cmdPrint_Click()
   
  Dim i As Long
   
   On Error GoTo Err_Proc
   
   '// Prevent frmPrinting again until done
   Picture1.SetFocus
   picPrintOptions.Enabled = False
   lblPrintingPg.Visible = True
   cmdPrint.Visible = False
   
   Select Case OptionV
   Case 0 '// Copy to clipboard
      Clipboard.Clear
      Clipboard.SetData Picture1.Picture, vbCFBitmap
   Case 1 '// Print current page
      lblPrintingPg.Caption = "Printing page " & ViewPage + 1
      lblPrintingPg.Refresh
      Call PrintPictureBox(Picture1, True, False)
   Case 2 '// Print range
      For i = Val(txtFrom) - 1 To Val(txtTo) - 1
         lblPrintingPg.Caption = "Printing page " & CStr(i + 1) & " of " & txtTo
         lblPrintingPg.Refresh
         Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(i) & ".bmp")
         Call PrintPictureBox(Picture1, True, False)
      Next i
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
   Case 4
      picGetFolder.Visible = True
      picGetFolder.ZOrder
   Case Else '// Print all
      gcPrint.SendToPrinter = True '// Send to Printer */
      Unload Me
   End Select
   
   '// Restore normal view
   picPrintOptions.Enabled = True
   cmdPrint.Visible = True
   picPrintOptions.Visible = False
   lblPrintingPg.Visible = False
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "cmdPrint_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub cmdQuit_Click()
   
   picGetFolder.Visible = False
   '// Restore normal view
   picPrintOptions.Enabled = True
   cmdPrint.Visible = True
   picPrintOptions.Visible = False
   lblPrintingPg.Visible = False
   
End Sub

Private Sub cmdUpOne_Click()
   
   Dir1.Path = Dir1.List(-2)
   
End Sub

Private Sub cmd_Print_Click()
   
   On Error GoTo Err_Proc
   
   txtTo.Text = PageNumber + 1
   OptionV = 3
   Call optText_Click(OptionV)
   picGoto.Visible = False
   
   picPrintOptions.left = Me.Width - (Picture2.Width + picPrintOptions.Width + 50)
   picPrintOptions.top = cmd_Print.top
   
   picGetFolder.left = Me.Width - (Picture2.Width + picGetFolder.Width + 50)
   picGetFolder.top = cmd_Print.top
   
   picPrintOptions.Visible = True
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "cmd_print_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub cmd_quit_Click()
   
   gcPrint.SendToPrinter = False
   Unload Me
   
End Sub

Private Sub Command1_Click(Index As Integer)
   
   On Local Error Resume Next
   
   If Index = 0 Then
      ViewPage = ViewPage - 1
      If ViewPage < 0 Then ViewPage = 0
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
   Else
      ViewPage = ViewPage + 1
      If ViewPage > PageNumber Then ViewPage = PageNumber
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
   End If
   
   Picture1.top = 0
   picPrintOptions.Visible = False
   picGoto.Visible = False
   VScroll1.Value = 0
   HScroll1.Value = 0
   Call DisplayPages
   
End Sub

Private Sub Decode_KeyUp(KeyCode As Integer, Shift As Integer)
   
   On Local Error Resume Next
   
   Select Case KeyCode
   Case 38 '// Arrow up
      VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
   Case 40 '// Arrow down
      VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
   Case 37 '// Arrow left
      If HScroll1.Visible = False Then
         Call Command1_Click(0)
      Else
         HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
      End If
   Case 39 '// Arrow right
      If HScroll1.Visible = False Then
         Call Command1_Click(1)
      Else
         HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
      End If
   Case 33 '// PageUp
      Call Command1_Click(0)
   Case 34 '// PageDown
      Call Command1_Click(1)
   Case 71 '// G
      Call cmdGoTo_Click
   Case 35, 36 '// Home, End
      Dim NewPageNo As Long
      If KeyCode = 36 Then
         NewPageNo = 0
      Else
         NewPageNo = PageNumber
      End If
      ViewPage = NewPageNo
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
      picPrintOptions.Visible = False
      picGetFolder.Visible = False
      picGoto.Visible = False
      VScroll1.Value = 0
      HScroll1.Value = 0
      Call DisplayPages
   End Select
   
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Dir1.Path = Dir1.List(Dir1.ListIndex)
   
End Sub

Private Sub DisplayPages()
   
   On Error GoTo Err_Proc
   
   Label1 = CStr(ViewPage + 1) & vbNewLine & "- of -" & vbNewLine & CStr(PageNumber + 1)
   
   If Picture1.Width > Me.Width - Picture2.Width Then
      picHScroll.Visible = True
   Else
      picHScroll.Visible = False
   End If
   
   If Picture1.Height >= Me.Height Then
      VScroll1.Visible = True
   Else
      VScroll1.Visible = False
   End If
   
   If picFullPage.Visible Then cmdFullPage_Click
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "DisplayPages"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub Drive1_Change()
   
   Dir1.Path = Drive1.Drive
   
End Sub

Private Sub Form_Activate()
   
   On Error GoTo Err_Proc
   
   Screen.MousePointer = vbDefault
   Call DisplayPages
   If Picture1.Width < Me.Width - Picture2.Width Then
      Picture1.Move ((Me.Width - Picture2.Width) - Picture1.Width) \ 2, 0
   End If
   Set cmdFullPage.PictureNormal = imgFit(0).Picture
   Label5 = "Goto Page#" & vbCrLf & "(1 to " & CStr(PageNumber + 1) & ")"
   Picture1.SetFocus
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "Form_Activate"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub Form_Click()
   
   picPrintOptions.Visible = False
   picGetFolder.Visible = False
   picGoto.Visible = False
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 71 Or KeyAscii = 103 Then cmdGoTo_Click
   
End Sub

Private Sub Form_Load()
   
  Dim osv As OSVersionInfo
  Const VER_PLATFORM_WIN32_NT = 2
   
   On Error GoTo Err_Proc
   
   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      If osv.PlatformID = VER_PLATFORM_WIN32_NT Then
         UseStretchBit = True
      Else
         UseStretchBit = False
      End If
   End If
   
   Me.Move 0, 0, Screen.Width, Screen.Height
   Picture1.Move 0, 0
   
   VScroll1.Height = Me.Height - cmdGoTo.top - cmdGoTo.Height - 500
   HScroll1.Width = Me.Width - Picture2.Width - 500
   
   TempDir = Environ("TEMP") & "\"
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "Form_Load"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
  Dim tFilename As String
   
   On Error GoTo Err_Proc
   
   '// Remove preview pages
   tFilename = Dir$(TempDir & "PPview*.bmp")
   If tFilename > vbNullString Then
      Do
         Kill TempDir & tFilename
         tFilename = Dir$(TempDir & "PPview*.bmp")
      Loop Until tFilename = vbNullString
   End If
   
   PageNumber = 0
   ViewPage = 0
   Set frmMultiPgPreview = Nothing
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "Form_Unload"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub HScroll1_Change()
   
   On Local Error Resume Next
   Picture1.left = -(HScroll1.Value)
   Picture1.SetFocus
   On Local Error GoTo 0
   
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   On Local Error Resume Next
   Select Case KeyCode
   Case 38 '// Arrow up
      VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
   Case 40 '// Arrow down
      VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
   Case 33 '// PageUp
      Call Command1_Click(0)
   Case 34 '// PageDown
      Call Command1_Click(1)
   Case 71 '// G
      Call cmdGoTo_Click
   Case 35, 36 '// Home, End
      Dim NewPageNo As Long
      If KeyCode = 36 Then
         NewPageNo = 0
      Else
         NewPageNo = PageNumber
      End If
      ViewPage = NewPageNo
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
      picPrintOptions.Visible = False
      picGetFolder.Visible = False
      picGoto.Visible = False
      VScroll1.Value = 0
      HScroll1.Value = 0
      Call DisplayPages
   End Select
   
End Sub

Private Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint _
                          As Boolean = False, Optional AllowNegative As Boolean = False) As Boolean
   
   On Error GoTo Err_Proc
   
   If KeyAscii > 0 And KeyAscii <> 8 Then
      If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
      If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
      If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
         KeyAscii = False
         IsNumber = False
      Else
         IsNumber = True
      End If
   Else
      IsNumber = IsNumeric(CheckString)
   End If
   
Exit_Here:
   Exit Function
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "IsNumber"
   Err.Clear
   Resume Exit_Here
   
End Function

Private Sub optPrint_Click(Index As Integer)
   
  Dim i As Byte
   
   On Error GoTo Err_Proc
   
   OptionV = Index
   For i = 0 To 4
      If Index = i Then
         optPrint(i).Picture = optArt(1).Picture
      Else
         optPrint(i).Picture = optArt(0).Picture
      End If
   Next i
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "optPrint_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub optText_Click(Index As Integer)
   
  Dim i As Byte
   
   On Error GoTo Err_Proc
   
   OptionV = Index
   For i = 0 To 4
      If Index = i Then
         optPrint(i).Picture = optArt(1).Picture
      Else
         optPrint(i).Picture = optArt(0).Picture
      End If
   Next i
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "optText_Click"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub picFullPage_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Call Decode_KeyUp(KeyCode, Shift)
   
End Sub

Private Sub Picture1_Click()
   
   picPrintOptions.Visible = False
   picGetFolder.Visible = False
   picGoto.Visible = False
   
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Call Decode_KeyUp(KeyCode, Shift)
   
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   On Local Error Resume Next
   
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.X = X
      PanSet.Y = Y
      MousePointer = vbSizePointer
   End If
   
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
  Dim nTop  As Integer
  Dim nLeft As Integer
   
   On Local Error Resume Next
   
   If Button = vbLeftButton And Shift = 0 Then
      
      '// new coordinates?
      With Picture1
         nTop = -(.top + (Y - PanSet.Y))
         nLeft = -(.left + (X - PanSet.X))
      End With
      
      '// Check limits
      With VScroll1
         If .Visible Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -Picture1.top
         End If
      End With
      
      With HScroll1
         If .Visible Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -Picture1.left
         End If
      End With
      
      Picture1.Move -nLeft, -nTop
      
   End If
   
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   On Local Error Resume Next
   
   If Button = vbLeftButton And Shift = 0 Then
      If VScroll1.Visible Then VScroll1.Value = -(Picture1.top)
      If HScroll1.Visible Then HScroll1.Value = -(Picture1.left)
   End If
   
   MousePointer = vbDefault
   
End Sub

Private Sub PrintPictureBox(pBox As PictureBox, Optional ScaleToFit As Boolean = True, Optional MaintainRatio As Boolean = True)
   
  Dim xmin As Single
  Dim ymin As Single
  Dim wid As Single
  Dim hgt As Single
  Dim aspect As Single
   
   On Error GoTo Err_Proc
   
   Screen.MousePointer = vbHourglass
   
   If Not ScaleToFit Then
      wid = Printer.ScaleX(pBox.ScaleWidth, pBox.ScaleMode, Printer.ScaleMode)
      hgt = Printer.ScaleY(pBox.ScaleHeight, pBox.ScaleMode, Printer.ScaleMode)
      xmin = (Printer.ScaleWidth - wid) / 2
      ymin = (Printer.ScaleHeight - hgt) / 2
   Else
      aspect = pBox.ScaleHeight / pBox.ScaleWidth
      wid = Printer.ScaleWidth
      hgt = Printer.ScaleHeight
      
      If MaintainRatio Then
         If hgt / wid > aspect Then
            hgt = aspect * wid
            xmin = Printer.ScaleLeft
            ymin = (Printer.ScaleHeight - hgt) / 2
         Else
            wid = hgt / aspect
            xmin = (Printer.ScaleWidth - wid) / 2
            ymin = Printer.ScaleTop
         End If
      End If
   End If
   
   Printer.PaintPicture pBox.Picture, xmin, ymin, wid, hgt, , , , , vbSrcCopy
   Printer.EndDoc
   
   Printer.Orientation = gcPrint.Orientation
   
   Screen.MousePointer = vbDefault
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "PrintPictureBox"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub txtFrom_Change()
   
   If Val(txtFrom) < 1 Then txtFrom = 1
   If Val(txtFrom) > Val(txtTo) Then txtFrom = txtTo
   
End Sub

Private Sub txtFrom_GotFocus()
   
   txtFrom.SelStart = 0
   txtFrom.SelLength = Len(txtFrom)
   
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
   
   On Error GoTo Err_Proc
   
   Select Case KeyCode
   Case 38  '// "+"
      txtFrom = txtFrom + 1
      KeyCode = False
   Case 40  '// "-"
      txtFrom = txtFrom - 1
      KeyCode = False
   End Select
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "txtFrom_KeyDown"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
   
   IsNumber txtFrom, KeyAscii, False, False
   
End Sub

Private Sub txtGoto_Change()
   
   If Val(txtGoto) > PageNumber + 1 Then txtGoto = PageNumber + 1
   
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
   
   On Error GoTo Err_Proc
   
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdGoToOK_Click
   ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "txtGoto_KeyPress"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub txtTo_Change()
   
   If Val(txtTo) > PageNumber + 1 Then txtTo = PageNumber + 1
   If Val(txtTo) < Val(txtFrom) Then txtTo = txtFrom
   
End Sub

Private Sub txtTo_GotFocus()
   
   txtTo.SelStart = 0
   txtTo.SelLength = Len(txtTo)
   
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
   
   On Error GoTo Err_Proc
   
   Select Case KeyCode
   Case 38  '// "+"
      txtTo = txtTo + 1
      KeyCode = False
   Case 40  '// "-"
      txtTo = txtTo - 1
      KeyCode = False
   End Select
   
Exit_Here:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMultiPgPreview", "txtTo_KeyDown"
   Err.Clear
   Resume Exit_Here
   
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
   
   IsNumber txtTo, KeyAscii, False, False
   
End Sub

Private Sub VScroll1_Change()
   
   On Local Error Resume Next
   Picture1.top = -(VScroll1.Value)
   Picture1.SetFocus
   On Local Error GoTo 0
   
End Sub

Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   On Local Error Resume Next
   
   Select Case KeyCode
   Case 37, 33 '// Arrow left, PageUp
      If HScroll1.Visible = False Then
         Call Command1_Click(0)
      Else
         HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
      End If
   Case 39, 34 '// Arrow right, PageDown
      If HScroll1.Visible = False Then
         Call Command1_Click(1)
      Else
         HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
      End If
   Case 71 '// G
      Call cmdGoTo_Click
   Case 35, 36 '// Home, End
      Dim NewPageNo As Long
      If KeyCode = 36 Then
         NewPageNo = 0
      Else
         NewPageNo = PageNumber
      End If
      ViewPage = NewPageNo
      Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
      picPrintOptions.Visible = False
      picGetFolder.Visible = False
      picGoto.Visible = False
      VScroll1.Value = 0
      HScroll1.Value = 0
      Call DisplayPages
   End Select
   
End Sub

