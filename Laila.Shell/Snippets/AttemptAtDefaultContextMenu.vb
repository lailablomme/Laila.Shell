'    Case "copy", "cut", "paste", "delete"
'        Dim button As Button = New Button()
'        Dim image As Image = New Image()
'        image.Width = 16
'        image.Height = 16
'        image.Margin = New Thickness(4)
'        Select Case item.Tag.ToString().Split(vbTab)(1)
'            Case "copy" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/copy16.png")
'            Case "cut" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/cut16.png")
'            Case "paste" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/paste16.png")
'            Case "delete" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/delete16.png")
'        End Select
'        button.Content = image
'        button.ToolTip = CType(item, MenuItem).Header.ToString().Replace("&", "")
'        button.Margin = New Thickness(0, 0, 4, 0)
'        _menu.ButtonsTop.Add(button)
'Case Else
'    _menu.Items.Add(item)
'Dim cmcb As ContextMenuCB = New ContextMenuCB()
'Dim itemArray As PCUITEMID_CHILD_ARRAY = marshalItemArray(lastpidls)
'Dim dcm As New DEFCONTEXTMENU With {
'        .psf = ShellFolder,
'        .pcmcb = cmcb,
'        .pidlFolder = folderpidl, ' folderpidl,
'        .cidl = itemArray.cItems,
'        .apidl = itemArray.pItems,
'        .punkAssociationInfo = IntPtr.Zero,
'        .cKeys = 0,
'        .aKeys = IntPtr.Zero
'    }
''If IntPtr.Zero.Equals(ptrContextMenu) Then
'Dim hr As Integer = Functions.SHCreateDefaultContextMenu(dcm, GetType(IContextMenu).GUID, ptrContextMenu)

'Imports System.Runtime.InteropServices

'Private Function marshalItemArray(itemPtrs As IntPtr()) As PCUITEMID_CHILD_ARRAY
'    Dim size As Integer = Marshal.SizeOf(GetType(IntPtr)) * itemPtrs.Length
'    Dim pArray As IntPtr = Marshal.AllocHGlobal(size)
'    Marshal.Copy(itemPtrs, 0, pArray, itemPtrs.Length)

'    Dim itemArray As New PCUITEMID_CHILD_ARRAY With {
'    .cItems = CUInt(itemPtrs.Length),
'    .pItems = pArray
'}

'    Return itemArray
'End Function

'Dim contextMenu As IContextMenu, ptrContextMenu As IntPtr
'Dim folderpidl As IntPtr, shellItemPtr As IntPtr, shellFolderPtr As IntPtr
'shellItemPtr = Marshal.GetIUnknownForObject(_shellItem2)
'Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
'Dim desktoppidl As IntPtr, shellItemPtr2 As IntPtr ', shellFolderPtr As IntPtr
'shellItemPtr2 = Marshal.GetIUnknownForObject(Shell.Desktop._shellItem2)
'Functions.SHGetIDListFromObject(shellItemPtr2, desktoppidl)
'Dim isSpecial As Boolean
'Dim it As Item
'If Not items Is Nothing AndAlso items.Count > 0 Then
'it = items(0)
'Else
'it = Me
'End If
'If it.FullPath.StartsWith("shell:::") Then isSpecial = True
'        'Do While Not Me.Parent Is Nothing AndAlso Not it.FullPath = Shell.Desktop.FullPath AndAlso Not isSpecial
'        '    it = Me.Parent
'        '    If it.FullPath.StartsWith("shell:::") OrElse it.FullPath.StartsWith("\\") Then isSpecial = True
'        'Loop

'        Dim sf As IShellFolder
''shellFolderPtr = Marshal.GetIUnknownForObject(sf)

''If Not items Is Nothing AndAlso items.Count > 0 Then
'Dim pidls(If(items Is Nothing OrElse items.Count = 0, 0, items.Count - 1)) As IntPtr
'Dim lastpidls(If(items Is Nothing OrElse items.Count = 0, 0, items.Count - 1)) As IntPtr
'Dim sp() As IntPtr
'Dim custom As IntPtr, pf As IntPtr
'Dim itemArray As PCUITEMID_CHILD_ARRAY
'If items Is Nothing OrElse items.Count = 0 Then
'' right-click on folder background...
'pidls(0) = folderpidl
'lastpidls(0) = Functions.ILFindLastID(pidls(0))

'If isSpecial Then
'_shellFolder.CreateViewObject(IntPtr.Zero, GetType(IContextMenu).GUID, ptrContextMenu)
'contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
'Else
'If Me.Parent Is Nothing Then
'' ...of desktop
'Functions.SHGetDesktopFolder(sf)
'itemArray = marshalItemArray(pidls)
'sp = Nothing
'pf = desktoppidl
'Else
'' of any other folder
'sf = Me.Parent._shellFolder
'itemArray = marshalItemArray(lastpidls)
'Dim sip As IntPtr ', shellFolderPtr As IntPtr
'sip = Marshal.GetIUnknownForObject(Me.Parent._shellItem2)
'Functions.SHGetIDListFromObject(sip, pf)
'sp = lastpidls
'End If
'End If
'Else
'' right-clock on folder/item
'For i = 0 To items.Count - 1
'shellItemPtr = Marshal.GetIUnknownForObject(items(i)._shellItem2)
'Functions.SHGetIDListFromObject(shellItemPtr, pidls(i))
'lastpidls(i) = Functions.ILFindLastID(pidls(i))
'Next

'If isSpecial Then
'_shellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptrContextMenu)
''pf = folderpidl
'sp = Nothing
'Else
'If (Me.FullPath = Shell.Desktop.FullPath AndAlso items.Count = 1 AndAlso items(0).FullPath = Shell.Desktop.FullPath) Then
'' right-click on desktop folder
'Functions.SHGetDesktopFolder(sf)
'itemArray = marshalItemArray(pidls)
'sp = Nothing
'pf = desktoppidl
'Else
'' right-click on any other folder/item
'sf = _shellFolder
'itemArray = marshalItemArray(lastpidls)
'pf = folderpidl
'sp = lastpidls
'End If
'End If
'End If


'Dim dcm As New DEFCONTEXTMENU With {
'        .psf = sf,
'        .pidlFolder = pf, ' folderpidl,
'        .cidl = itemArray.cItems,
'        .apidl = itemArray.pItems,
'        .punkAssociationInfo = IntPtr.Zero,
'        .cKeys = 0,
'        .aKeys = IntPtr.Zero
'    }
'If IntPtr.Zero.Equals(ptrContextMenu) Then
'Dim hr As Integer = Functions.SHCreateDefaultContextMenu(dcm, GetType(IContextMenu).GUID, ptrContextMenu)
'End If

''_shellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptrContextMenu)
'contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
















''Dim apidl As IntPtr = Marshal.AllocCoTaskMem(IntPtr.Size * pidls.Count)
''For i = 0 To pidls.Count - 1
''    Marshal.WriteIntPtr(apidl, i * IntPtr.Size, pidls(i))
''Next
'Imports System.Runtime.InteropServices

'Dim apidl(pidls.Count - 1) As ITEMIDLIST
'For i = 0 To pidls.Count - 1
'apidl(i) = Marshal.PtrToStructure(Of ITEMIDLIST)(pidls(i))
'Next
''_shellFolder.CreateViewObject(IntPtr.Zero, GetType(IContextMenu).GUID, ptrContextMenu)
''contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
''Dim psf As IntPtr = Marshal.GetIUnknownForObject(_shellFolder)
''Dim p(0) As IntPtr : p(0) = folderpidl
'Dim dcm As DEFCONTEXTMENU
'dcm.hwnd = Shell._hwnd
'dcm.pidlFolder = IntPtr.Zero
'dcm.psf = Marshal.GetIUnknownForObject(_shellFolder)
'dcm.cidl = pidls.Count
'Dim size As Integer = 1000
'dcm.apidl = Marshal.AllocHGlobal(size)
'For i As Integer = 0 To apidl.Length - 1
'Dim offset As IntPtr = IntPtr.Add(dcm.apidl, i * Marshal.SizeOf(GetType(ITEMIDLIST)))
'Marshal.StructureToPtr(apidl(i), offset, False)
'Next
''Marshal.StructureToPtr(apidl, dcm.apidl, False)

'Dim h = Functions.SHCreateDefaultContextMenu(dcm, GetType(IContextMenu3).GUID, ptrContextMenu)
'Debug.WriteLine("SHCreateDefaultContextMenu returns " & h.ToString())
'contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
