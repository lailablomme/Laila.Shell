﻿Imports System.ComponentModel
Imports System.Reflection.Metadata
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class ContextMenu
    Implements IDisposable

    Public Event Click(id As Integer, verb As String, ByRef isHandled As Boolean)

    Public Property DefaultId As String

    Private _contextMenu As IContextMenu
    Private _contextMenu2 As IContextMenu2
    Private _contextMenu3 As IContextMenu3
    Private _hMenu As IntPtr
    Private _firstContextMenuCall As Boolean = True
    Private _menu As Controls.ContextMenu
    Private _parent As Folder
    Private _invokedId As String = Nothing
    Private disposedValue As Boolean

    Public Function GetContextMenu(parent As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean) As System.Windows.Controls.ContextMenu
        If Not _menu Is Nothing Then
            Return _menu
        End If

        Dim hasPaste As Boolean, pasteHeader As String
        If Not isDefaultOnly AndAlso (items Is Nothing OrElse items.Count = 0) Then
            ' check for paste
            _parent = parent.Parent
            makeContextMenu({parent}, isDefaultOnly)
            Dim pasteMenuItem As MenuItem = getMenuItems(_hMenu, -1).FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "paste")
            If Not pasteMenuItem Is Nothing Then
                hasPaste = True
                pasteHeader = pasteMenuItem.Header
            End If
            Me.ReleaseContextMenu()
        End If

        _parent = parent
        makeContextMenu(items, isDefaultOnly)

        Dim makeButton As Func(Of Object, String, Windows.Controls.Button) =
            Function(tag As Object, toolTip As String) As Windows.Controls.Button
                Dim button As System.Windows.Controls.Button = New System.Windows.Controls.Button()
                Dim image As Image = New Image()
                image.Width = 16
                image.Height = 16
                image.Margin = New Thickness(2)
                Select Case tag.ToString().Split(vbTab)(1)
                    Case "copy" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/copy16.png")
                    Case "cut" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/cut16.png")
                    Case "paste" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/paste16.png")
                    Case "delete" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/delete16.png")
                End Select
                button.Content = image
                button.ToolTip = toolTip
                'button.Margin = New Thickness(0, 0, 4, 0)
                button.Tag = tag
                Return button
            End Function

        ' make our own menu
        _menu = New Controls.ContextMenu()
        Dim menuItems As List(Of Control) = getMenuItems(_hMenu, -1)
        Dim lastMenuItem As Control
        For Each item In menuItems
            Select Case item.Tag?.ToString().Split(vbTab)(1)
                Case "copy", "cut", "paste", "delete", "pintohome"
                    ' don't add these
                Case Else
                    If Not (TypeOf item Is Separator AndAlso Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator) Then
                        _menu.Items.Add(item)
                        lastMenuItem = item
                    End If
            End Select
        Next
        Dim menuItem As MenuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "cut")
        If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "copy")
        If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "paste")
        If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", ""))) _
            Else If hasPaste Then _menu.Buttons.Add(makeButton("-1" & vbTab & "paste", pasteHeader.Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "delete")
        If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))

        AddHandler _menu.Closed,
            Sub(s As Object, e As EventArgs)
                If Not _invokedId Is Nothing Then
                    Me.InvokeCommand(_invokedId)
                Else
                    Me.ReleaseContextMenu()
                End If
            End Sub

        Dim wireMenuItems As Action(Of List(Of Control)) =
            Sub(wireItems As List(Of Control))
                For Each c As Control In wireItems
                    If (TypeOf c Is MenuItem AndAlso CType(c, MenuItem).Items.Count = 0) _
                        OrElse TypeOf c Is Windows.Controls.Button Then
                        If TypeOf c Is Windows.Controls.Button Then
                            AddHandler CType(c, Windows.Controls.Button).Click, AddressOf menuItem_Click
                        ElseIf TypeOf c Is MenuItem Then
                            AddHandler CType(c, MenuItem).Click, AddressOf menuItem_Click
                        End If
                    ElseIf TypeOf c Is MenuItem Then
                        wireMenuItems(CType(c, MenuItem).Items.Cast(Of Control).ToList())
                    End If
                Next
            End Sub
        wireMenuItems(_menu.Items.Cast(Of Control).ToList())
        wireMenuItems(_menu.Buttons.Cast(Of Control).ToList())

        If _menu.Items.Count = 0 Then
            _menu = Nothing
        End If

        'Dim idTPM As Integer = Functions.TrackPopupMenuEx(hMenu, &H100, 0, 0, Shell._hwnd, IntPtr.Zero) - 1
        'Debug.WriteLine("TrackPopupMenuEx returns " & idTPM)
        'InvokeCommand(contextMenu, Nothing, idTPM & vbTab)

        Return _menu ' New ContextMenu() 
    End Function

    Private Sub makeContextMenu(items As IEnumerable(Of Item), isDefaultOnly As Boolean)
        Dim ptrContextMenu As IntPtr
        Dim folderpidl As IntPtr, folderpidl2 As IntPtr, shellItemPtr As IntPtr
        Try
            shellItemPtr = Marshal.GetIUnknownForObject(_parent._shellItem2)
            Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
        Finally
            If Not IntPtr.Zero.Equals(shellItemPtr) Then
                Marshal.Release(shellItemPtr)
            End If
        End Try

        Dim pidls(If(items Is Nothing OrElse items.Count = 0, 0, items.Count - 1)) As IntPtr
        Dim lastpidls(If(items Is Nothing OrElse items.Count = 0, 0, items.Count - 1)) As IntPtr

        Try
            If Not items Is Nothing AndAlso items.Count > 0 Then
                ' user clicked on an item
                Dim shellFolder As IShellFolder

                If Not (_parent.FullPath = Shell.Desktop.FullPath AndAlso Not items Is Nothing AndAlso items.Count = 1 _
                AndAlso items(0).FullPath = Shell.Desktop.FullPath) Then
                    shellFolder = _parent._shellFolder
                    For i = 0 To items.Count - 1
                        Try
                            shellItemPtr = Marshal.GetIUnknownForObject(items(i)._shellItem2)
                            Functions.SHGetIDListFromObject(shellItemPtr, pidls(i))
                        Finally
                            If Not IntPtr.Zero.Equals(shellItemPtr) Then
                                Marshal.Release(shellItemPtr)
                            End If
                        End Try
                        lastpidls(i) = Functions.ILFindLastID(pidls(i))
                    Next
                Else
                    Dim eaten As Integer
                    Functions.SHGetDesktopFolder(shellFolder)
                    shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, IO.Path.GetDirectoryName(_parent.FullPath), eaten, folderpidl, 0)
                    For i = 0 To items.Count - 1
                        shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, items(i).FullPath, eaten, pidls(i), 0)
                        lastpidls(i) = Functions.ILFindLastID(pidls(i))
                    Next
                End If

                shellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptrContextMenu)
                _contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))

                folderpidl2 = folderpidl
            Else
                ' user clicked on the background
                Dim shellFolder As IShellFolder = _parent._shellFolder
                If _parent.Parent Is Nothing Then
                    ' this is the desktop
                    Dim eaten As Integer
                    Functions.SHGetDesktopFolder(shellFolder)
                    shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, IO.Path.GetDirectoryName(_parent.FullPath), eaten, folderpidl, 0)
                    shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, _parent.FullPath, eaten, pidls(0), 0)
                    lastpidls(0) = Functions.ILFindLastID(pidls(0))
                    folderpidl2 = pidls(0)
                Else
                    ' this is any other folder
                    folderpidl2 = folderpidl
                    Try
                        shellItemPtr = Marshal.GetIUnknownForObject(_parent.Parent._shellItem2)
                        Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
                    Finally
                        If Not IntPtr.Zero.Equals(shellItemPtr) Then
                            Marshal.Release(shellItemPtr)
                        End If
                    End Try
                    lastpidls(0) = Functions.ILFindLastID(folderpidl2)
                End If

                shellFolder.CreateViewObject(IntPtr.Zero, GetType(IContextMenu).GUID, ptrContextMenu)
                _contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
            End If

            Dim ptr2 As IntPtr, ptr3 As IntPtr
            Try
                Marshal.QueryInterface(ptrContextMenu, GetType(IContextMenu2).GUID, ptr2)
                If Not IntPtr.Zero.Equals(ptr2) Then
                    _contextMenu2 = Marshal.GetObjectForIUnknown(ptr2)
                End If
                Marshal.QueryInterface(ptrContextMenu, GetType(IContextMenu3).GUID, ptr3)
                If Not IntPtr.Zero.Equals(ptr3) Then
                    _contextMenu3 = Marshal.GetObjectForIUnknown(ptr3)
                End If
            Finally
                If Not IntPtr.Zero.Equals(ptr2) Then
                    Marshal.Release(ptr2)
                End If
                If Not IntPtr.Zero.Equals(ptr3) Then
                    Marshal.Release(ptr3)
                End If
            End Try

            _hMenu = Functions.CreatePopupMenu()
            Dim flags As Integer = CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS Or CMF.CMF_EXPLORE
            If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
            _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)

            If _firstContextMenuCall Then
                ' somehow very first call doesn't return all items
                Functions.DestroyMenu(_hMenu)
                _hMenu = Functions.CreatePopupMenu()
                _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)
                _firstContextMenuCall = False
            End If

            'Dim shellExtInitPtr As IntPtr, shellExtInit As IShellExtInit, dataObject As ComTypes.IDataObject
            'Try
            '    Marshal.QueryInterface(ptrContextMenu, GetType(IShellExtInit).GUID, shellExtInitPtr)
            '    If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
            '        shellExtInit = Marshal.GetObjectForIUnknown(shellExtInitPtr)
            '        Functions.SHCreateDataObject(folderpidl, lastpidls.Count, lastpidls, IntPtr.Zero, GetType(ComTypes.IDataObject).GUID, dataObject)
            '        shellExtInit.Initialize(folderpidl2, dataObject, IntPtr.Zero)
            '    End If
            'Finally
            '    If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
            '        Marshal.Release(shellExtInitPtr)
            '    End If
            '    If Not shellExtInit Is Nothing Then
            '        Marshal.ReleaseComObject(shellExtInit)
            '    End If
            '    If Not dataObject Is Nothing Then
            '        Marshal.ReleaseComObject(dataObject)
            '    End If
            'End Try
        Finally
            If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                Marshal.Release(ptrContextMenu)
            End If
            For i = 0 To pidls.Count - 1
                Marshal.FreeCoTaskMem(pidls(i))
            Next
            Marshal.FreeCoTaskMem(folderpidl)
        End Try
    End Sub

    Private Function getMenuItems(hMenu2 As IntPtr, parentIndex As Integer) As List(Of Control)
        If parentIndex >= 0 Then
            Dim lParam As Integer = (&HFFFF0000) Or (parentIndex And &HFFFF)
            If Not _contextMenu3 Is Nothing Then
                Dim ptr3 As IntPtr
                Dim h As HRESULT = _contextMenu3.HandleMenuMsg2(WM.INITMENUPOPUP, hMenu2, lParam, ptr3)
                Debug.WriteLine("contextMenu3 returned" & h.ToString())
                If Not IntPtr.Zero.Equals(ptr3) Then
                    Marshal.Release(ptr3)
                End If
            ElseIf Not _contextMenu2 Is Nothing Then
                _contextMenu2.HandleMenuMsg(WM.INITMENUPOPUP, hMenu2, lParam)
            End If
        End If

        Dim result As List(Of Control) = New List(Of Control)()

        For i = 0 To Functions.GetMenuItemCount(hMenu2) - 1
            Dim mii As MENUITEMINFO
            mii.cbSize = CUInt(Marshal.SizeOf(mii))
            mii.fMask = MIIM.MIIM_STRING
            mii.dwTypeData = New String(" "c, 2050)
            mii.cch = mii.dwTypeData.Length - 2
            Functions.GetMenuItemInfo(hMenu2, i, True, mii)
            Dim header As String = mii.dwTypeData.Substring(0, mii.cch)

            mii = New MENUITEMINFO()
            mii.cbSize = CUInt(Marshal.SizeOf(mii))
            mii.fMask = MIIM.MIIM_BITMAP Or MIIM.MIIM_FTYPE
            Functions.GetMenuItemInfo(hMenu2, i, True, mii)

            Dim bitmapSource As BitmapSource
            If Not IntPtr.Zero.Equals(mii.hbmpItem) Then
                bitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpItem, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            Else
                bitmapSource = Nothing
            End If

            If mii.fType = MFT.SEPARATOR Then
                ' refuse initial and double separators
                If Not result.Count = 0 AndAlso Not TypeOf result(result.Count - 1) Is Separator Then
                    result.Add(New Separator())
                End If
            Else
                mii = New MENUITEMINFO()
                mii.cbSize = CUInt(Marshal.SizeOf(mii))
                mii.fMask = MIIM.MIIM_ID
                Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                Dim cmd As StringBuilder = New StringBuilder(), id As Integer
                cmd.Append(New String(" ", 2050))
                If mii.wID >= 1 AndAlso mii.wID <= 99999 Then
                    id = mii.wID - 1
                    _contextMenu.GetCommandString(id, GCS.VERBW, 0, cmd, 2048)
                    If cmd.Length = 0 Then
                        _contextMenu.GetCommandString(id, GCS.VERBA, 0, cmd, 2048)
                    End If
                End If

                Debug.WriteLine(header & "  " & id)

                mii = New MENUITEMINFO()
                mii.cbSize = CUInt(Marshal.SizeOf(mii))
                mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = header.Replace("&", "_"),
                    .Icon = New Image() With {.Source = bitmapSource},
                    .Tag = id & vbTab & cmd.ToString(),
                    .IsEnabled = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DISABLED), False, True),
                    .FontWeight = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
                }

                If CBool(mii.fState And MFS.MFS_DEFAULT) Then
                    DefaultId = menuItem.Tag
                End If

                If Not IntPtr.Zero.Equals(mii.hSubMenu) Then
                    Dim subMenu As List(Of Control) = getMenuItems(mii.hSubMenu, i)
                    For Each subMenuItem In subMenu
                        menuItem.Items.Add(subMenuItem)
                    Next
                End If

                result.Add(menuItem)
            End If
        Next

        ' remove trailing separators
        While result.Count > 0 AndAlso TypeOf result(result.Count - 1) Is Separator
            result.RemoveAt(result.Count - 1)
        End While

        Return result
    End Function


    Private Sub menuItem_Click(c As Control, e2 As EventArgs)
        Dim isHandled As Boolean = False

        _menu.IsOpen = False
        RaiseEvent Click(c.Tag.ToString().Split(vbTab)(0), c.Tag.ToString().Split(vbTab)(1), isHandled)

        If Not isHandled Then
            invokeCommandDelayed(c.Tag)
        End If
    End Sub

    Protected Sub invokeCommandDelayed(id As String)
        _invokedId = id
    End Sub

    Public Sub InvokeCommand(id As String)
        Dim cmi As New CMInvokeCommandInfoEx
        Debug.WriteLine("InvokeCommand " & id)
        If Convert.ToInt32(id.Split(vbTab)(0)) >= 0 Then
            cmi.lpVerb = New IntPtr(Convert.ToUInt32(id.Split(vbTab)(0)))
            cmi.lpVerbW = New IntPtr(Convert.ToUInt32(id.Split(vbTab)(0)))
        Else
            cmi.lpVerb = Marshal.StringToHGlobalAnsi(id.Split(vbTab)(1))
            cmi.lpVerbW = Marshal.StringToHGlobalUni(id.Split(vbTab)(1))
        End If
        'cmi.lpDirectory = _parent.FullPath
        'cmi.lpDirectoryW = _parent.FullPath
        cmi.fMask = CMIC.UNICODE Or CMIC.ASYNCOK
        If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then cmi.fMask = cmi.fMask Or CMIC.CONTROL_DOWN
        If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then cmi.fMask = cmi.fMask Or CMIC.SHIFT_DOWN
        cmi.nShow = SW.SHOWNORMAL
        'cmi.hwnd = Shell._hwnd
        cmi.cbSize = CUInt(Marshal.SizeOf(cmi))

        Dim h As HRESULT = _contextMenu.InvokeCommand(cmi)
        Debug.WriteLine("InvokeCommand returned " & h.ToString())

        Me.ReleaseContextMenu()
    End Sub

    Private Sub ReleaseContextMenu()
        If Not IntPtr.Zero.Equals(_hMenu) Then
            Functions.DestroyMenu(_hMenu)
        End If
        If Not _contextMenu Is Nothing Then
            Marshal.ReleaseComObject(_contextMenu)
        End If
        If Not _contextMenu2 Is Nothing Then
            Marshal.ReleaseComObject(_contextMenu2)
        End If
        If Not _contextMenu3 Is Nothing Then
            Marshal.ReleaseComObject(_contextMenu3)
        End If
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' dispose managed state (managed objects)
            End If

            ' free unmanaged resources (unmanaged objects) and override finalizer
            ' set large fields to null
            'Me.ReleaseContextMenu()
            disposedValue = True
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    Protected Overrides Sub Finalize()
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=False)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
