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
