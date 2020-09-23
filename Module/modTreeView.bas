Attribute VB_Name = "modTreeView"
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

Option Explicit

Private mtvwObjTrv         As TreeView
Private mlngNodeCheckIndex As Long
Private mlngNodeCheckColor As Long
Private Const C_lngColorG  As Long = &H80000011

Public Function AddNode(ByVal vstrParent As String, _
                        ByVal vstrKey As String, _
                        ByVal vstrText As String, _
                        Optional ByVal vstrTag As String = vbNullString, _
                        Optional ByVal vstrImage As String = "ROOT", _
                        Optional ByVal vblnIsChecked As Boolean = True, _
                        Optional ByVal vblnIsExpanded As Boolean = False, _
                        Optional ByVal vblnIsSorted As Boolean = True) As Boolean
   
  Dim NodX As MSComctlLib.Node
   
   On Error GoTo Err_Proc
   
   With mtvwObjTrv
      Set NodX = .Nodes.Add(vstrParent, tvwChild, vstrKey, vstrText)
      NodX.Tag = vstrTag
      NodX.Image = vstrImage
      NodX.Checked = vblnIsChecked
      NodX.Expanded = vblnIsExpanded
      NodX.Sorted = vblnIsSorted
      If Not vblnIsChecked Then NodX.ForeColor = C_lngColorG '// disabled text color
      AddNode = NodX.Index
   End With
   AddNode = True
   
Exit_Proc:
   Exit Function
   
Err_Proc:
   AddNode = False
   'Err_Handler True, Err.Number, Err.Description, "modTreeView", "AddNode"
   Err.Clear
   Resume Exit_Proc
   
End Function

Public Sub ClearNodes()
   
   mtvwObjTrv.Nodes.Clear
   
End Sub

Public Sub CollapseAllNodes(Optional ByVal vlngStartNode As Long = 1)
   
  Dim lngI As Long
   
   '// Collapse all nodes
   With mtvwObjTrv
      For lngI = vlngStartNode To .Nodes.Count
         .Nodes.Item(lngI).Expanded = False
      Next lngI
   End With
   
End Sub

Public Sub DisableCheck()
   
   '// Put this in the MouseUp event of the TreeView control
   '// to prevent a node from being checked
   
   If mlngNodeCheckIndex Then
      If mlngNodeCheckColor = C_lngColorG Then
         mtvwObjTrv.Nodes.Item(mlngNodeCheckIndex).Checked = False
      End If
   End If
   
End Sub

Public Sub ExpandAllNodes(Optional ByVal vlngStartNode As Long = 1)
   
  Dim lngI As Long
   
   '// Expanded all nodes
   With mtvwObjTrv
      For lngI = vlngStartNode To .Nodes.Count
         .Nodes.Item(lngI).Expanded = True
      Next lngI
   End With
   
End Sub

Public Function FindTag(ByVal vstrNodeTag As String, _
                        Optional ByVal vstrNodeParent As String = vbNullString, _
                        Optional ByVal vblnMakeVisible As Boolean = False) As Long
   
  Dim lngI     As Long
  Dim strPNode As String
   
   On Error Resume Next
   
   vstrNodeTag = LCase$(vstrNodeTag)
   vstrNodeParent = LCase$(vstrNodeParent)
   
   With mtvwObjTrv
      For lngI = 1 To .Nodes.Count
         strPNode = LCase$(.Nodes(lngI).Parent & vbNullString)
         If LCase$(.Nodes(lngI).Tag) = vstrNodeTag And strPNode = vstrNodeParent Then
            If vblnMakeVisible Then
               .Nodes(lngI).EnsureVisible
               .Nodes(lngI).Selected = True
            End If
            FindTag = lngI
            Exit For
         End If
      Next lngI
   End With
   
End Function

Public Function FindText(ByVal vstrNodeText As String, _
                         Optional ByVal vstrNodeParent As String = vbNullString, _
                         Optional ByVal vblnMakeVisible As Boolean = False) As Long
   
  Dim lngI     As Long
  Dim strPNode As String
   
   On Error Resume Next
   
   vstrNodeText = LCase$(vstrNodeText)
   vstrNodeParent = LCase$(vstrNodeParent)
   
   With mtvwObjTrv
      For lngI = 1 To .Nodes.Count
         strPNode = LCase$(.Nodes(lngI).Parent & vbNullString)
         If LCase$(.Nodes(lngI).Text) = vstrNodeText And strPNode = vstrNodeParent Then
            If vblnMakeVisible Then
               .Nodes(lngI).EnsureVisible
               .Nodes(lngI).Selected = True
            End If
            FindText = lngI
            Exit For
         End If
      Next lngI
   End With
   
End Function

Public Function GetNextKey() As String
   
   GetNextKey = "K" & CStr(mtvwObjTrv.Nodes.Count + 1)
   
End Function

Public Sub InitializeTreeView(ByRef rtvwObjTrv As TreeView)
   
   Set mtvwObjTrv = rtvwObjTrv
   
End Sub

Public Function IsNodeChecked(ByVal vstrNodeText As String, _
                              Optional ByVal vstrNodeParent As String, _
                              Optional ByRef rlngNodeIndex As Long) As Boolean
   
  Dim lngI     As Long
  Dim strPNode As String
   
   On Error Resume Next
   
   vstrNodeText = LCase$(vstrNodeText)
   vstrNodeParent = LCase$(vstrNodeParent)
   
   With mtvwObjTrv
      For lngI = 1 To .Nodes.Count
         strPNode = LCase$(.Nodes(lngI).Parent & vbNullString)
         If LCase$(.Nodes(lngI).Text) = vstrNodeText And strPNode = vstrNodeParent Then
            IsNodeChecked = .Nodes(lngI).Checked
            rlngNodeIndex = lngI
            Exit For
         End If
      Next lngI
   End With
   
End Function

Public Function MakeUniqueKey(ByVal vstrNodeKey As String) As String
   
  Dim lngI As Long
   
   MakeUniqueKey = vstrNodeKey
   
   With mtvwObjTrv
      For lngI = 1 To .Nodes.Count
         If .Nodes(lngI).Key = vstrNodeKey Then
            MakeUniqueKey = MakeUniqueKey & CStr(mtvwObjTrv.Nodes.Count + 1)
            Exit For
         End If
      Next lngI
   End With
   
End Function

Public Sub NodeCheckedEvent(ByRef Node As MSComctlLib.Node)
   
  Dim lngI                 As Long
  Dim blnFlag              As Boolean
  Dim ParentNode           As Node
  Dim ChildNode            As Node
  Dim ChildNode2           As Node
  Dim CurrentNode          As Node
  Dim blnCurrentNodeStatus As Boolean
   
   On Error GoTo Err_Proc
   
   Set CurrentNode = Node
   
   blnFlag = False
   lngI = CurrentNode.Index
   '// Set variables for DisableCheck Sub
   mlngNodeCheckIndex = lngI
   mlngNodeCheckColor = CurrentNode.ForeColor
   
   blnCurrentNodeStatus = CurrentNode.Checked
   
   '// Look from top down
   If Not mtvwObjTrv.Nodes.Item(lngI).Child Is Nothing Then
      Set ParentNode = mtvwObjTrv.Nodes.Item(lngI).Child.FirstSibling
      
      Do While Not ParentNode Is Nothing
         '// Ignore nodes that have been grayed out
         If ParentNode.ForeColor <> C_lngColorG Then
            ParentNode.Checked = CurrentNode.Checked
          Else
            ParentNode.Checked = False
         End If
         
         If Not ParentNode.Child Is Nothing Then
            Set ChildNode = ParentNode.Child
            Do While Not ChildNode Is Nothing
               '// Ignore nodes that have been grayed out
               If ChildNode.ForeColor <> C_lngColorG Then
                  ChildNode.Checked = CurrentNode.Checked
                Else
                  ChildNode.Checked = False
               End If
               
               If Not ChildNode.Child Is Nothing Then
                  Set ChildNode2 = ChildNode.Child
                  Do While Not ChildNode2 Is Nothing
                     '// Ignore nodes that have been grayed out
                     If ChildNode2.ForeColor <> C_lngColorG Then
                        ChildNode2.Checked = CurrentNode.Checked
                      Else
                        ChildNode2.Checked = False
                     End If
                     If Not ChildNode2.Next Is Nothing Then
                        Set ChildNode2 = ChildNode2.Next
                      Else
                        Set ChildNode2 = ChildNode2.Child
                     End If
                  Loop
               End If
               
               If Not ChildNode.Next Is Nothing Then
                  Set ChildNode = ChildNode.Next
                Else
                  Set ChildNode = ChildNode.Child
               End If
            Loop
         End If
         
         Set ParentNode = ParentNode.Next
      Loop
   End If
   
   '// Look from Bottom up
   If blnCurrentNodeStatus = True Then '// Checked
      '// Check all parent nodes
      Do While Not mtvwObjTrv.Nodes.Item(lngI).Parent Is Nothing
         If mtvwObjTrv.Nodes.Item(lngI).Parent.ForeColor <> C_lngColorG Then
            mtvwObjTrv.Nodes.Item(lngI).Parent.Checked = CurrentNode.Checked
         End If
         lngI = mtvwObjTrv.Nodes.Item(lngI).Parent.Index
      Loop
      
    Else '// Unchecked
      
      If Not CurrentNode.Parent Is Nothing Then
         Set ParentNode = CurrentNode.Parent.Child
         
         Do While Not ParentNode Is Nothing
            Set ChildNode = ParentNode.FirstSibling
            
            Do While Not ChildNode Is Nothing
               If ChildNode.Checked = True And ChildNode.ForeColor <> C_lngColorG Then
                  blnFlag = True
                  Exit Do
               End If
               Set ChildNode = ChildNode.Next
            Loop
            
            If blnFlag = False Then
               If Not ParentNode.Parent Is Nothing Then
                  ParentNode.Parent.Checked = False
                  blnFlag = False
               End If
             Else
               Exit Do
            End If
            
            If Not ParentNode.Parent Is Nothing Then
               Set ParentNode = ParentNode.Parent
             Else
               Set ParentNode = ParentNode.Parent
            End If
         Loop
      End If
   End If
   
   '// Clean up
   Set CurrentNode = Nothing
   Set ParentNode = Nothing
   Set ChildNode = Nothing
   Set ChildNode2 = Nothing
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "modTreeView", "NodeCheckedEvent"
   Err.Clear
   Resume Exit_Proc
   
End Sub

Public Sub SetDefaultTree()
   
   '// Load default Tree structure
  Dim NodX As MSComctlLib.Node
   
   On Error GoTo Err_Proc
   
   With mtvwObjTrv
      
      .Nodes.Clear
      
      Set NodX = .Nodes.Add(, , "Project", "Project")
      NodX.Tag = vbNullString
      NodX.Image = "ROOT"
      NodX.Expanded = True
      NodX.Checked = True
      
      Call AddNode("Project", "Dependencies", "Dependencies", , "FOLDER", False, False)
      Call AddNode("Project", "Forms", "Forms", , "FOLDER", , True)
      Call AddNode("Project", "Modules", "Modules", , "FOLDER", , True)
      Call AddNode("Project", "Classes", "Classes", , "FOLDER", , True)
      Call AddNode("Project", "User controls", "User controls", , "FOLDER", , True)
      Call AddNode("Project", "User documents", "User documents", , "FOLDER", , True)
      Call AddNode("Project", "Designer", "Designer", , "FOLDER", , True)
      
   End With
   
Exit_Proc:
   Exit Sub
   
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "modTreeView", "SetDefaultTree"
   Err.Clear
   Resume Exit_Proc
   
End Sub

