Imports System.ComponentModel
Imports System.IO
Imports System.Reflection.Metadata
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Public Class ContextMenu
    Implements IDisposable

    Public Event CommandInvoked(sender As Object, e As CommandInvokedEventArgs)

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
            If Not IntPtr.Zero.Equals(_hMenu) Then
                Dim pasteMenuItem As MenuItem = getMenuItems(_hMenu, -1).FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "paste")
                If Not pasteMenuItem Is Nothing Then
                    hasPaste = True
                    pasteHeader = pasteMenuItem.Header
                End If
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
                    Case "rename" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/rename16.png")
                    Case "delete" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/delete16.png")
                End Select
                button.Content = image
                button.ToolTip = toolTip
                button.Tag = tag
                Return button
            End Function

        Dim makeToggleButton As Func(Of Object, String, Boolean, Windows.Controls.Primitives.ToggleButton) =
            Function(tag As Object, toolTip As String, isChecked As Boolean) As Windows.Controls.Primitives.ToggleButton
                Dim button As System.Windows.Controls.Primitives.ToggleButton = New System.Windows.Controls.Primitives.ToggleButton()
                Dim image As Image = New Image()
                image.Width = 16
                image.Height = 16
                image.Margin = New Thickness(2)
                Select Case tag.ToString().Split(vbTab)(1)
                    Case "laila.shell.(un)pin" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/pin16.png")
                End Select
                button.Content = image
                button.ToolTip = toolTip
                button.Tag = tag
                button.IsChecked = isChecked
                Return button
            End Function

        Dim osver As Version = Environment.OSVersion.Version
        Dim isWindows11 As Boolean = osver.Major = 10 AndAlso osver.Minor = 0 AndAlso osver.Build >= 22000

        ' make our own menu
        _menu = New Controls.ContextMenu()
        If Not IntPtr.Zero.Equals(_hMenu) Then
            Dim menuItems As List(Of Control) = getMenuItems(_hMenu, -1)
            Dim lastMenuItem As Control
            For Each item In menuItems
                Dim verb As String = item.Tag?.ToString().Split(vbTab)(1)
                Select Case verb
                    Case "copy", "cut", "paste", "delete", "pintohome", "rename"
                        ' don't add these
                    Case Else
                        Dim isNotDoubleSeparator As Boolean = Not (TypeOf item Is Separator AndAlso
                            Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator)
                        Dim isNotDoubleOneDriveItem As Boolean = verb Is Nothing OrElse
                            Not (isWindows11 AndAlso
                                (verb.StartsWith("{5250E46F-BB09-D602-5891-F476DC89B70") _
                                 OrElse verb = "MakeAvailableOffline" _
                                 OrElse verb = "MakeAvailableOnline"))
                        If isNotDoubleSeparator AndAlso isNotDoubleOneDriveItem Then
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
            menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "rename")
            If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
            menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "delete")
            If Not menuItem Is Nothing Then _menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
            If items.Count = 1 Then
                Dim test As Item = Item.FromParsingName(items(0).FullPath, Nothing)
                If Not test Is Nothing Then ' this won't work for all items 
                    test.Dispose()
                    Dim isPinned As Boolean = PinnedItems.GetIsPinned(items(0).FullPath)
                    _menu.Buttons.Add(makeToggleButton("-1" & vbTab & "laila.shell.(un)pin", If(isPinned, "Unpin item", "Pin item"), isPinned))
                End If
            End If
        End If

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
                        OrElse TypeOf c Is ButtonBase Then
                        If TypeOf c Is Windows.Controls.Button Then
                            AddHandler CType(c, Windows.Controls.Button).Click, AddressOf menuItem_Click
                        ElseIf TypeOf c Is Windows.Controls.Primitives.ToggleButton Then
                            AddHandler CType(c, Windows.Controls.Primitives.ToggleButton).Click, AddressOf menuItem_Click
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

                If Not (_parent.FullPath = Shell.Desktop.FullPath AndAlso items.Count = 1 _
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
                If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                    _contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
                End If

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
                If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                    _contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
                End If
            End If

            If Not IntPtr.Zero.Equals(ptrContextMenu) Then
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
            End If

            If Not _contextMenu Is Nothing Then
                _hMenu = Functions.CreatePopupMenu()
                Dim flags As Integer = CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS Or CMF.CMF_EXPLORE Or CMF.CMF_CANRENAME
                If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
                _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)

                If _firstContextMenuCall Then
                    ' somehow very first call doesn't return all items
                    Functions.DestroyMenu(_hMenu)
                    _hMenu = Functions.CreatePopupMenu()
                    _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)
                    _firstContextMenuCall = False
                End If
            End If

            If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                Dim shellExtInitPtr As IntPtr, shellExtInit As IShellExtInit, dataObject As ComTypes.IDataObject
                Try
                    Marshal.QueryInterface(ptrContextMenu, GetType(IShellExtInit).GUID, shellExtInitPtr)
                    If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
                        shellExtInit = Marshal.GetObjectForIUnknown(shellExtInitPtr)
                        Functions.SHCreateDataObject(folderpidl, lastpidls.Count, lastpidls, IntPtr.Zero, GetType(ComTypes.IDataObject).GUID, dataObject)
                        shellExtInit.Initialize(folderpidl2, dataObject, IntPtr.Zero)
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
                        Marshal.Release(shellExtInitPtr)
                    End If
                    If Not shellExtInit Is Nothing Then
                        Marshal.ReleaseComObject(shellExtInit)
                    End If
                    If Not dataObject Is Nothing Then
                        Marshal.ReleaseComObject(dataObject)
                    End If
                End Try
            End If
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
            mii.fMask = MIIM.MIIM_BITMAP Or MIIM.MIIM_FTYPE Or MIIM.MIIM_CHECKMARKS
            Functions.GetMenuItemInfo(hMenu2, i, True, mii)

            Dim bitmapSource As BitmapSource
            If Not IntPtr.Zero.Equals(mii.hbmpItem) Then
                bitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpItem, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            ElseIf Not IntPtr.Zero.Equals(mii.hbmpUnchecked) Then
                bitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpUnchecked, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
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

                Debug.WriteLine(header & "  " & id & vbTab & cmd.ToString())

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
        _menu.IsOpen = False

        Dim e As CommandInvokedEventArgs = New CommandInvokedEventArgs() With {
            .Id = c.Tag.ToString().Split(vbTab)(0),
            .Verb = c.Tag.ToString().Split(vbTab)(1)
        }
        If TypeOf c Is ToggleButton Then e.IsChecked = CType(c, ToggleButton).IsChecked
        RaiseEvent CommandInvoked(Me, e)

        If Not e.IsHandled Then
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
            _contextMenu = Nothing
        End If
        If Not _contextMenu2 Is Nothing Then
            Marshal.ReleaseComObject(_contextMenu2)
        End If
        If Not _contextMenu3 Is Nothing Then
            Marshal.ReleaseComObject(_contextMenu3)
        End If
    End Sub

    Public Sub DoRename(point As Point, width As Double, item As Item, grid As Grid)
        Dim originalName As String, isDrive As Boolean

        Dim doRename As Action(Of String) =
            Sub(newName As String)
                If Not originalName = newName Then
                    ' rename item
                    If isDrive Then
                        Functions.SetVolumeLabelW(item.FullPath, newName)
                        Shell.RaiseNotificationEvent(Me, New NotificationEventArgs() With {
                                                             .Item1Path = item.FullPath, .[Event] = SHCNE.UPDATEITEM})
                    Else
                        Dim fileOperation As IFileOperation
                        Dim h As HRESULT = Functions.CoCreateInstance(Guids.CLSID_FileOperation, IntPtr.Zero, 1, GetType(IFileOperation).GUID, fileOperation)
                        fileOperation.RenameItem(item._shellItem2, newName, Nothing)
                        fileOperation.PerformOperations()
                        Marshal.ReleaseComObject(fileOperation)

                        ' notify pinned items & frequent folders
                        Dim newFullPath As String =
                        item.FullPath.Substring(0, item.FullPath.LastIndexOf(IO.Path.DirectorySeparatorChar) + 1) + newName
                        PinnedItems.RenameItem(item.FullPath, newFullPath)
                        FrequentFolders.RenameItem(item.FullPath, newFullPath)
                    End If
                End If
            End Sub

        ' make textbox
        Dim textBox As System.Windows.Controls.TextBox
        textBox = New System.Windows.Controls.TextBox() With {
            .Margin = New Thickness(point.X, point.Y + 1, 0, 0),
            .HorizontalAlignment = HorizontalAlignment.Left,
            .VerticalAlignment = VerticalAlignment.Top,
            .Width = width
        }
        textBox.MaxLength = 260
        If item.FullPath.Equals(IO.Path.GetPathRoot(item.FullPath)) Then
            isDrive = True
            item._shellItem2.GetDisplayName(SIGDN.PARENTRELATIVEEDITING, originalName)
        Else
            originalName = item.FullPath.Substring(item.FullPath.LastIndexOf(IO.Path.DirectorySeparatorChar) + 1)
        End If
        textBox.Text = originalName
        grid.Children.Add(textBox)
        textBox.Focus()

        ' select filename without extension
        textBox.SelectionStart = 0
        If textBox.Text.Contains(".") Then
            textBox.SelectionLength = textBox.Text.IndexOf(".")
        Else
            textBox.SelectionLength = textBox.Text.Length
        End If

        ' hook textbox
        AddHandler textBox.LostFocus,
            Sub(s2 As Object, e2 As RoutedEventArgs)
                grid.Children.Remove(textBox)
                If Not textBox.Tag = "cancel" AndAlso Not String.IsNullOrWhiteSpace(textBox.Text) Then
                    doRename(textBox.Text)
                End If
            End Sub
        AddHandler textBox.PreviewKeyDown,
            Sub(s2 As Object, e2 As KeyEventArgs)
                Select Case e2.Key
                    Case Key.Enter
                        grid.Children.Remove(textBox)
                    Case Key.Escape
                        textBox.Tag = "cancel"
                        grid.Children.Remove(textBox)
                End Select
            End Sub
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
