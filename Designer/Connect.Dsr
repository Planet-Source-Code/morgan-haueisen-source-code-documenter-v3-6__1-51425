VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6585
   ClientLeft      =   2235
   ClientTop       =   1785
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   11615
   _Version        =   393216
   Description     =   "VB Project Documenter"
   DisplayName     =   "VB Project Documenter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents MenuEvents      As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1
Private MenuItem                   As CommandBarControl
Private CommandBarMenu             As CommandBar
Private Const MenuName             As String = "Add-Ins"

Private Sub AddinInstance_Initialize()
   
   '/* According the MSDN help file 'Add-In Essentials' Designer needs this even if it is empty.
   
End Sub

Private Sub AddinInstance_OnAddinsUpdate(custom() As Variant)
   
   '/* According the MSDN help file 'Add-In Essentials' Designer needs this even if it is empty.
   
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
   
   On Error Resume Next
   '/* If using a form
   Unload frmMain
   On Error GoTo 0
   
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)
   
  Dim i                   As Long
  Dim tString          As String
  Dim PreserveClipGraphic As Variant
   
   Set VBInstance = Application
   
   If ConnectMode = ext_cm_External Then
      '/* do nothing
    Else
      If AddInMenuAvailable Then
         On Error Resume Next
         Set CommandBarMenu = VBInstance.CommandBars(MenuName)
         On Error GoTo 0
         If CommandBarMenu Is Nothing Then
            MsgBox App.Title & " was loaded but could not be connected to the " & MenuName & " menu.", vbCritical
          Else
            DoEvents
            With CommandBarMenu
               Set MenuItem = .Controls.Add(msoControlButton)
               i = .Controls.Count - 1
               If .Controls(i).BeginGroup Then
                  If Not .Controls(i - 1).BeginGroup Then
                     '/* menu separator required
                     MenuItem.BeginGroup = True
                  End If
               End If
            End With
            
            '/* Set Menu Caption
            '/* The "&" in next line sets the First letter in the application's
            '/* ProductName as the hotkey you may need to change this for your language/font
            MenuItem.Caption = "&" & App.ProductName & " V" & App.Major & "." & App.Minor & "." & App.Revision & "..."
            
            '/* Add an icon to the menu bar
            On Error Resume Next
            With Clipboard
               '/* set menu picture
               tString = .GetText
               PreserveClipGraphic = .GetData
               
               .SetData LoadResPicture(101, vbResBitmap)
               MenuItem.PasteFace
               .Clear
               If IsObject(PreserveClipGraphic) Then
                  .SetData PreserveClipGraphic
               End If
               If LenB(tString) Then
                  .SetText tString
               End If
            End With
            
            Set MenuEvents = VBInstance.Events.CommandBarEvents(MenuItem)
            
         End If
      End If
   End If
   
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)
   
   On Error Resume Next
   Select Case RemoveMode
    Case vbext_dm_HostShutdown
      '/* If using a form
      Unload frmMain
    Case vbext_dm_UserClosed
      '/* If using a form
      Set frmMain = Nothing
   End Select
   MenuItem.Delete
   On Error GoTo 0
   
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
   
   '/* According the MSDN help file 'Add-In Essentials' Designer needs this even if it is empty.
   
End Sub

Private Sub AddinInstance_Terminate()
   
   '/* According the MSDN help file 'Add-In Essentials' Designer needs this even if it is empty.
   
   '/* If using a form
   Set frmMain = Nothing
   
End Sub

Private Function AddInMenuAvailable() As Boolean
   
   AddInMenuAvailable = Not VBInstance.CommandBars("Add-Ins") Is Nothing
   If Not AddInMenuAvailable Then
      MsgBox "'Add-Ins' Menu is unavailable.", vbCritical
   End If
   
End Function

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, _
                             handled As Boolean, _
                             CancelDefault As Boolean)
   
   '/* Start the application
   frmMain.Show
   
End Sub

