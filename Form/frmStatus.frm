VERSION 5.00
Begin VB.Form frmStatus 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   2625
   ClientTop       =   2280
   ClientWidth     =   7305
   ControlBox      =   0   'False
   ForeColor       =   &H80000015&
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7305
   Begin VB.PictureBox ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7245
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3585
      Width           =   7305
   End
   Begin VB.TextBox lblWorking 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Working... Please Wait"
      Top             =   0
      Width           =   7320
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   2565
      Left            =   210
      TabIndex        =   2
      Top             =   855
      Width           =   6900
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   7305
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mcProg As clsProgressBar

Private Sub Form_Initialize()
   Set mcProg = New clsProgressBar
   With mcProg
      .Style = pbSolid2Color
      .picBox = ProgressBar1
      .ShowCounts = False
      .ShowStatus = False
   End With
End Sub

Private Sub Form_Load()
   
  Dim cScreen As clsScreenSize
   
   Set cScreen = New clsScreenSize
   cScreen.CenterForm Me
   Set cScreen = Nothing
   
   Me.Show
   DoEvents
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmStatus = Nothing
End Sub

