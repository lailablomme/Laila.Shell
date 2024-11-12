Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Channels
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class Menus
        Inherits FrameworkElement
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged, AddressOf OnSelectedItemsCoerce))
        Public Shared ReadOnly ItemContextMenuProperty As DependencyProperty = DependencyProperty.Register("ItemContextMenu", GetType(Laila.Shell.Controls.ContextMenu), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly NewItemMenuProperty As DependencyProperty = DependencyProperty.Register("NewItemMenu", GetType(Laila.Shell.Controls.ContextMenu), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SortMenuProperty As DependencyProperty = DependencyProperty.Register("SortMenu", GetType(Laila.Shell.Controls.ContextMenu), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ViewMenuProperty As DependencyProperty = DependencyProperty.Register("ViewMenu", GetType(Laila.Shell.Controls.ContextMenu), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDefaultOnlyProperty As DependencyProperty = DependencyProperty.Register("IsDefaultOnly", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoAutoDisposeProperty As DependencyProperty = DependencyProperty.Register("DoAutoDispose", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanCutProperty As DependencyProperty = DependencyProperty.Register("CanCut", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanCopyProperty As DependencyProperty = DependencyProperty.Register("CanCopy", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanPasteProperty As DependencyProperty = DependencyProperty.Register("CanPaste", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanRenameProperty As DependencyProperty = DependencyProperty.Register("CanRename", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanDeleteProperty As DependencyProperty = DependencyProperty.Register("CanDelete", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanShareProperty As DependencyProperty = DependencyProperty.Register("CanShare", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly UpdateDelayProperty As DependencyProperty = DependencyProperty.Register("UpdateDelay", GetType(Integer?), GetType(Menus), New FrameworkPropertyMetadata(0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Public Event CommandInvoked(sender As Object, e As CommandInvokedEventArgs)

        Public Property DefaultId As String

        Private _contextMenu As IContextMenu
        Private _newItemContextMenu As IContextMenu
        Private _selectedItemsContextMenu As IContextMenu
        Private _contextMenu2 As IContextMenu2
        Private _contextMenu3 As IContextMenu3
        Private _hMenu As IntPtr
        Private _newItemhMenu As IntPtr
        Private _selectedItemshMenu As IntPtr
        Private Shared _firstContextMenuCall As Boolean = True
        Private _parent As Folder
        Private _invokedId As String = Nothing
        Private disposedValue As Boolean
        Private _lastItems As IEnumerable(Of Item)
        Private _lastFolder As Folder
        Private _updateTimer As Timer
        Private _isCheckingInternally As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(Menus), New FrameworkPropertyMetadata(GetType(Menus)))
        End Sub

        Public Sub New()
            Dim viewMenu As ContextMenu = New ContextMenu()
            For Each item In Shell.FolderViews
                Dim viewSubMenuItem As MenuItem = New MenuItem() With {
                    .Header = item.Key,
                    .Icon = New Image() With {.Source = New BitmapImage(New Uri(item.Value.Item1, UriKind.Absolute))},
                    .Tag = "View:" & item.Key,
                    .IsCheckable = True
                }
                AddHandler viewSubMenuItem.Checked,
                    Sub(s2 As Object, e2 As EventArgs)
                        If Not _isCheckingInternally Then
                            Me.Folder.View = item.Key
                        End If
                    End Sub
                AddHandler viewSubMenuItem.Unchecked,
                    Sub(s2 As Object, e2 As EventArgs)
                        If Not _isCheckingInternally Then
                            _isCheckingInternally = True
                            viewSubMenuItem.IsChecked = True
                            _isCheckingInternally = False
                        End If
                    End Sub
                viewMenu.Items.Add(viewSubMenuItem)
            Next
            Me.ViewMenu = viewMenu
            initializeViewMenu()
        End Sub

        Private Function getContextMenu(parent As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean) As Laila.Shell.Controls.ContextMenu
            _lastItems = items

            Dim hasPaste As Boolean
            If Not isDefaultOnly AndAlso (items Is Nothing OrElse items.Count = 0) Then
                ' check for paste by checking if it would accept a drop
                Dim dataObject As IDataObject
                Functions.OleGetClipboard(dataObject)

                Dim dropTarget As IDropTarget, pidl As IntPtr, shellItemPtr As IntPtr, dropTargetPtr As IntPtr
                Try
                    shellItemPtr = Marshal.GetIUnknownForObject(parent.ShellItem2)
                    Functions.SHGetIDListFromObject(shellItemPtr, pidl)
                    Using parent2 As Folder = parent.GetParent()
                        If Not parent2 Is Nothing Then
                            Dim lastpidl As IntPtr = Functions.ILFindLastID(pidl), shellFolder As IShellFolder = parent2.ShellFolder
                            Try
                                shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {lastpidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                            Finally
                                If Not shellFolder Is Nothing Then
                                    Marshal.ReleaseComObject(shellFolder)
                                End If
                            End Try
                        Else
                            Dim shellFolder As IShellFolder
                            Functions.SHGetDesktopFolder(shellFolder)
                            ' desktop
                            shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {pidl}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
                        End If
                    End Using
                    If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                        dropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
                    Else
                        dropTarget = Nothing
                    End If

                    If Not dropTarget Is Nothing Then
                        Dim effect As DROPEFFECT = Laila.Shell.DROPEFFECT.DROPEFFECT_COPY
                        Dim hr As HRESULT = dropTarget.DragEnter(dataObject, 0, New WIN32POINT(), effect)
                        dropTarget.DragLeave()

                        hasPaste = hr = HRESULT.Ok AndAlso effect <> DROPEFFECT.DROPEFFECT_NONE
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(shellItemPtr) Then
                        Marshal.Release(shellItemPtr)
                    End If
                    If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                        Marshal.Release(dropTargetPtr)
                    End If
                    If Not IntPtr.Zero.Equals(pidl) Then
                        Marshal.FreeCoTaskMem(pidl)
                    End If
                    If Not dropTarget Is Nothing Then
                        Marshal.ReleaseComObject(dropTarget)
                    End If
                End Try
            End If

            _parent = parent
            makeContextMenu(items, isDefaultOnly)

            Dim makeButton As Func(Of Object, String, System.Windows.Controls.Button) =
                Function(tag As Object, toolTip As String) As System.Windows.Controls.Button
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

            Dim makeToggleButton As Func(Of Object, String, Boolean, ToggleButton) =
                Function(tag As Object, toolTip As String, isChecked As Boolean) As ToggleButton
                    Dim button As ToggleButton = New ToggleButton()
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
            Dim menu As Controls.ContextMenu = New Controls.ContextMenu()
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
                            (Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator))
                            Dim isNotInitialSeparator As Boolean = Not (TypeOf item Is Separator AndAlso menu.Items.Count = 0)
                            Dim isNotDoubleOneDriveItem As Boolean = verb Is Nothing OrElse
                            Not (isWindows11 AndAlso
                                (verb.StartsWith("{5250E46F-BB09-D602-5891-F476DC89B70") _
                                 OrElse verb.StartsWith("{1FA0E654-C9F2-4A1F-9800-B9A75D744B0") _
                                 OrElse verb = "MakeAvailableOffline" _
                                 OrElse verb = "MakeAvailableOnline"))
                            If isNotDoubleSeparator AndAlso isNotInitialSeparator AndAlso isNotDoubleOneDriveItem Then
                                menu.Items.Add(item)
                                lastMenuItem = item
                            End If
                    End Select
                Next
                Dim menuItem As MenuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "cut")
                If Not menuItem Is Nothing Then menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
                menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "copy")
                If Not menuItem Is Nothing Then menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
                menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "paste")
                If Not menuItem Is Nothing Then menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", ""))) _
                Else If hasPaste Then menu.Buttons.Add(makeButton("-1" & vbTab & "paste", "Paste"))
                menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "rename")
                If Not menuItem Is Nothing Then menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
                menuItem = menuItems.FirstOrDefault(Function(i) i.Tag?.ToString().Split(vbTab)(1) = "delete")
                If Not menuItem Is Nothing Then menu.Buttons.Add(makeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
                'If Not items Is Nothing AndAlso items.Count = 1 Then
                '    Dim test As Item = Item.FromParsingName(items(0).FullPath, Nothing)
                '    If Not test Is Nothing Then ' this won't work for all items 
                '        test.Dispose()
                '        Dim isPinned As Boolean = PinnedItems.GetIsPinned(items(0).FullPath)
                '        menu.Buttons.Add(makeToggleButton("-1" & vbTab & "laila.shell.(un)pin", If(isPinned, "Unpin item", "Pin item"), isPinned))
                '    End If
                'End If
            End If

            AddHandler menu.Closed,
                Sub(s As Object, e As EventArgs)
                    If Not _invokedId Is Nothing Then
                        Me.InvokeCommand(_selectedItemsContextMenu, _invokedId)
                    ElseIf Me.DoAutoDispose Then
                        Me.Dispose()
                    Else
                        _lastFolder = Nothing
                        _lastItems = Nothing
                        Me.Update()
                    End If
                End Sub

            Dim wireMenuItems As Action(Of List(Of Control)) =
                Sub(wireItems As List(Of Control))
                    For Each c As Control In wireItems
                        If (TypeOf c Is MenuItem AndAlso CType(c, MenuItem).Items.Count = 0) _
                        OrElse TypeOf c Is ButtonBase Then
                            If TypeOf c Is System.Windows.Controls.Button Then
                                AddHandler CType(c, System.Windows.Controls.Button).Click, AddressOf menuItem_Click
                            ElseIf TypeOf c Is ToggleButton Then
                                AddHandler CType(c, ToggleButton).Click, AddressOf menuItem_Click
                            ElseIf TypeOf c Is MenuItem Then
                                AddHandler CType(c, MenuItem).Click, AddressOf menuItem_Click
                            End If
                        ElseIf TypeOf c Is MenuItem Then
                            wireMenuItems(CType(c, MenuItem).Items.Cast(Of Control).ToList())
                        End If
                    Next
                End Sub
            wireMenuItems(menu.Items.Cast(Of Control).ToList())
            wireMenuItems(menu.Buttons.Cast(Of Control).ToList())

            'Dim idTPM As Integer = Functions.TrackPopupMenuEx(hMenu, &H100, 0, 0, Shell._hwnd, IntPtr.Zero) - 1
            'Debug.WriteLine("TrackPopupMenuEx returns " & idTPM)
            'InvokeCommand(contextMenu, Nothing, idTPM & vbTab)

            Return menu ' New ContextMenu() 
        End Function

        Private Function getNewItemMenu(contextMenu As Laila.Shell.Controls.ContextMenu)
            If Not contextMenu Is Nothing Then
                Dim newMenuItem As MenuItem = contextMenu.Items.Cast(Of Control) _
                    .FirstOrDefault(Function(c) TypeOf c Is MenuItem _
                        AndAlso Not c.Tag Is Nothing _
                        AndAlso c.Tag.ToString().Split(vbTab)(1) = "New")

                If Not newMenuItem Is Nothing Then
                    Dim menu As Laila.Shell.Controls.ContextMenu = New Controls.ContextMenu()
                    For Each control In newMenuItem.Items
                        If TypeOf control Is Separator Then
                            menu.Items.Add(New Separator())
                        Else
                            Dim menuItem As MenuItem = control
                            Dim copiedMenuItem As MenuItem = New MenuItem() With {
                                .Header = menuItem.Header,
                                .Icon = menuItem.Icon,
                                .Tag = menuItem.Tag
                            }
                            AddHandler copiedMenuItem.Click, AddressOf newMenuItem_Click
                            menu.Items.Add(copiedMenuItem)
                        End If
                    Next

                    AddHandler menu.Closed,
                        Sub(s As Object, e As EventArgs)
                            If Not _invokedId Is Nothing Then
                                Me.InvokeCommand(_newItemContextMenu, _invokedId)
                            ElseIf Me.DoAutoDispose Then
                                Me.Dispose()
                            Else
                                _lastFolder = Nothing
                                _lastItems = Nothing
                                Me.Update()
                            End If
                        End Sub

                    Return menu
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Function

        Private Sub makeContextMenu(items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            Dim ptrContextMenu As IntPtr
            Dim folderpidl As IntPtr, folderpidl2 As IntPtr, shellItemPtr As IntPtr
            Try
                shellItemPtr = Marshal.GetIUnknownForObject(_parent.ShellItem2)
                Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
            Finally
                If Not IntPtr.Zero.Equals(shellItemPtr) Then
                    Marshal.Release(shellItemPtr)
                End If
            End Try

            Dim pidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr
            Dim lastpidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr
            Dim flags As Integer = CMF.CMF_NORMAL

            Try
                Dim shellFolder As IShellFolder
                Try
                    If Not items Is Nothing AndAlso items.Count > 0 Then
                        ' user clicked on an item
                        flags = flags Or CMF.CMF_ITEMMENU

                        If Not (_parent.FullPath = Shell.Desktop.FullPath AndAlso items.Count = 1 _
                AndAlso items(0).FullPath = Shell.Desktop.FullPath) Then
                            shellFolder = _parent.ShellFolder
                            For i = 0 To items.Count - 1
                                Try
                                    shellItemPtr = Marshal.GetIUnknownForObject(items(i).ShellItem2)
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
                        If _parent.FullPath.Equals(Shell.Desktop.FullPath) Then
                            ' this is the desktop
                            Dim eaten As Integer
                            Functions.SHGetDesktopFolder(shellFolder)
                            shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, IO.Path.GetDirectoryName(_parent.FullPath), eaten, folderpidl, 0)
                            shellFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, _parent.FullPath, eaten, folderpidl2, 0)
                        Else
                            shellFolder = _parent.ShellFolder
                            ' this is any other folder
                            folderpidl2 = folderpidl
                            Try
                                Using parent2 = _parent.GetParent()
                                    shellItemPtr = Marshal.GetIUnknownForObject(parent2.ShellItem2)
                                    Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)
                                End Using
                            Finally
                                If Not IntPtr.Zero.Equals(shellItemPtr) Then
                                    Marshal.Release(shellItemPtr)
                                End If
                            End Try
                            'lastpidls(0) = Functions.ILFindLastID(folderpidl2)
                        End If

                        shellFolder.CreateViewObject(IntPtr.Zero, GetType(IContextMenu).GUID, ptrContextMenu)
                        If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                            _contextMenu = Marshal.GetTypedObjectForIUnknown(ptrContextMenu, GetType(IContextMenu))
                        End If
                    End If
                Finally
                    If Not shellFolder Is Nothing Then
                        Marshal.ReleaseComObject(shellFolder)
                    End If
                End Try

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

                If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                    Dim shellExtInitPtr As IntPtr, shellExtInit As IShellExtInit, dataObject As ComTypes.IDataObject
                    Try
                        Marshal.QueryInterface(ptrContextMenu, GetType(IShellExtInit).GUID, shellExtInitPtr)
                        If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
                            shellExtInit = Marshal.GetObjectForIUnknown(shellExtInitPtr)
                            If Not (items Is Nothing OrElse items.Count = 0) Then _
                                Functions.SHCreateDataObject(folderpidl, lastpidls.Count, lastpidls, IntPtr.Zero, GetType(ComTypes.IDataObject).GUID, dataObject)
                            shellExtInit.Initialize(folderpidl2, If(items Is Nothing OrElse items.Count = 0, Nothing, dataObject), IntPtr.Zero)
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

                If Not _contextMenu Is Nothing Then
                    _hMenu = Functions.CreatePopupMenu()
                    flags = flags Or CMF.CMF_EXTENDEDVERBS Or CMF.CMF_EXPLORE Or CMF.CMF_CANRENAME
                    If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
                    _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)

                    'If _firstContextMenuCall Then
                    '    ' somehow very first call doesn't return all items
                    '    Functions.DestroyMenu(_hMenu)
                    '    _hMenu = Functions.CreatePopupMenu()
                    '    _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)
                    '    _firstContextMenuCall = False
                    'End If
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

        Private Sub newMenuItem_Click(c As Control, e2 As EventArgs)
            Me.NewItemMenu.IsOpen = False

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

        Private Sub menuItem_Click(c As Control, e2 As EventArgs)
            Me.ItemContextMenu.IsOpen = False

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
            Me.InvokeCommand(_selectedItemsContextMenu, id)
        End Sub

        Public Sub InvokeCommand(contextMenu As IContextMenu, id As String)
            Select Case id.Split(vbTab)(1)
                Case "Windows.ModernShare"
                    Dim assembly As Assembly = Assembly.LoadFrom("Laila.Shell.WinRT.dll")
                    Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ModernShare")
                    Dim methodInfo As MethodInfo = type.GetMethod("ShowShareUI")
                    Dim instance As Object = Activator.CreateInstance(type)
                    methodInfo.Invoke(instance, {_lastItems.ToList().Select(Function(i) i.FullPath).ToList()})
                Case Else
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

                    Dim h As HRESULT = contextMenu.InvokeCommand(cmi)
                    Debug.WriteLine("InvokeCommand returned " & h.ToString())
            End Select

            If Me.DoAutoDispose Then
                Me.Dispose()
            Else
                _lastFolder = Nothing
                _lastItems = Nothing
                Me.Update()
            End If
        End Sub

        Private Sub releaseContextMenuFull()
            releaseContextMenu()
            If Not IntPtr.Zero.Equals(_newItemhMenu) Then
                Functions.DestroyMenu(_newItemhMenu)
            End If
            If Not IntPtr.Zero.Equals(_selectedItemshMenu) Then
                Functions.DestroyMenu(_selectedItemshMenu)
            End If
            If Not _newItemContextMenu Is Nothing Then
                Marshal.ReleaseComObject(_newItemContextMenu)
                _newItemContextMenu = Nothing
            End If
            If Not _selectedItemsContextMenu Is Nothing Then
                Marshal.ReleaseComObject(_selectedItemsContextMenu)
                _selectedItemsContextMenu = Nothing
            End If
        End Sub

        Private Sub releaseContextMenu()
            Debug.WriteLine("releaseContextMenu()")
            If Not IntPtr.Zero.Equals(_hMenu) Then
                Functions.DestroyMenu(_hMenu)
            End If
            If Not _contextMenu Is Nothing Then
                Marshal.ReleaseComObject(_contextMenu)
                _contextMenu = Nothing
            End If
            If Not _contextMenu2 Is Nothing Then
                Marshal.ReleaseComObject(_contextMenu2)
                _contextMenu2 = Nothing
            End If
            If Not _contextMenu3 Is Nothing Then
                Marshal.ReleaseComObject(_contextMenu3)
                _contextMenu3 = Nothing
            End If
        End Sub

        Public Shared Sub DoRename(point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double, item As Item, grid As Grid)
            Dim originalName As String, isDrive As Boolean

            Dim doRename As Action(Of String) =
            Sub(newName As String)
                If Not originalName = newName Then
                    ' rename item
                    If isDrive Then
                        Functions.SetVolumeLabelW(item.FullPath, newName)
                        Shell.RaiseNotificationEvent(Nothing, New NotificationEventArgs() With {
                                                     .Item1Path = item.FullPath, .[Event] = SHCNE.UPDATEITEM})
                    Else
                        Dim fileOperation As IFileOperation
                        Dim h As HRESULT = Functions.CoCreateInstance(Guids.CLSID_FileOperation, IntPtr.Zero, 1, GetType(IFileOperation).GUID, fileOperation)
                        fileOperation.RenameItem(item.ShellItem2, newName, Nothing)
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
                .Margin = New Thickness(point.X, point.Y, 0, 0),
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top,
                .Width = size.Width,
                .Height = size.Height,
                .MaxLength = 260,
                .TextWrapping = TextWrapping.Wrap,
                .TextAlignment = textAlignment,
                .UseLayoutRounding = True,
                .SnapsToDevicePixels = True,
                .FontSize = fontSize
            }
            textBox.SetValue(Panel.ZIndexProperty, 100)
            If item.FullPath.Equals(IO.Path.GetPathRoot(item.FullPath)) Then
                isDrive = True
                item.ShellItem2.GetDisplayName(SIGDN.PARENTRELATIVEEDITING, originalName)
            Else
                originalName = item.DisplayName
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

        Public Property CanCut As Boolean
            Get
                Return GetValue(CanCutProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanCutProperty, value)
            End Set
        End Property

        Public Property CanCopy As Boolean
            Get
                Return GetValue(CanCopyProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanCopyProperty, value)
            End Set
        End Property

        Public Property CanPaste As Boolean
            Get
                Return GetValue(CanPasteProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanPasteProperty, value)
            End Set
        End Property

        Public Property CanRename As Boolean
            Get
                Return GetValue(CanRenameProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanRenameProperty, value)
            End Set
        End Property

        Public Property CanDelete As Boolean
            Get
                Return GetValue(CanDeleteProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanDeleteProperty, value)
            End Set
        End Property

        Public Property CanShare As Boolean
            Get
                Return GetValue(CanShareProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanShareProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property

        Public Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property IsDefaultOnly As Boolean
            Get
                Return GetValue(IsDefaultOnlyProperty)
            End Get
            Set(value As Boolean)
                SetValue(IsDefaultOnlyProperty, value)
            End Set
        End Property

        Public Property DoAutoDispose As Boolean
            Get
                Return GetValue(DoAutoDisposeProperty)
            End Get
            Set(value As Boolean)
                SetValue(DoAutoDisposeProperty, value)
            End Set
        End Property

        Public Property UpdateDelay As Integer?
            Get
                Return GetValue(UpdateDelayProperty)
            End Get
            Set(value As Integer?)
                SetValue(UpdateDelayProperty, value)
            End Set
        End Property

        Public Property ItemContextMenu As Laila.Shell.Controls.ContextMenu
            Get
                Return GetValue(ItemContextMenuProperty)
            End Get
            Set(value As Laila.Shell.Controls.ContextMenu)
                SetValue(ItemContextMenuProperty, value)
            End Set
        End Property

        Public Property NewItemMenu As Laila.Shell.Controls.ContextMenu
            Get
                Return GetValue(NewItemMenuProperty)
            End Get
            Set(value As Laila.Shell.Controls.ContextMenu)
                SetValue(NewItemMenuProperty, value)
            End Set
        End Property

        Public Property SortMenu As Laila.Shell.Controls.ContextMenu
            Get
                Return GetValue(SortMenuProperty)
            End Get
            Set(value As Laila.Shell.Controls.ContextMenu)
                SetValue(SortMenuProperty, value)
            End Set
        End Property

        Public Property ViewMenu As Laila.Shell.Controls.ContextMenu
            Get
                Return GetValue(ViewMenuProperty)
            End Get
            Set(value As Laila.Shell.Controls.ContextMenu)
                SetValue(ViewMenuProperty, value)
            End Set
        End Property

        Public Sub Update()
            Dim didGetMenu As Boolean

            If Not Me.IsDefaultOnly AndAlso Not EqualityComparer(Of Folder).Default.Equals(_lastFolder, Me.Folder) Then
                releaseContextMenu()
                If Not IntPtr.Zero.Equals(_newItemhMenu) Then
                    Functions.DestroyMenu(_newItemhMenu)
                End If
                If Not _newItemContextMenu Is Nothing Then
                    Marshal.ReleaseComObject(_newItemContextMenu)
                    _newItemContextMenu = Nothing
                End If

                Me.ItemContextMenu = getContextMenu(Me.Folder, Nothing, False)
                Me.NewItemMenu = getNewItemMenu(Me.ItemContextMenu)
                _newItemContextMenu = _contextMenu
                _newItemhMenu = _hMenu
                _selectedItemsContextMenu = _contextMenu
                _contextMenu = Nothing
                _hMenu = IntPtr.Zero

                If Not _contextMenu2 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu2)
                    _contextMenu2 = Nothing
                End If
                If Not _contextMenu3 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu3)
                    _contextMenu3 = Nothing
                End If

                _lastFolder = Me.Folder
                didGetMenu = True

                If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                    Dim viewMenuItem As MenuItem = New MenuItem() With {
                        .Header = "View",
                        .Icon = New Image() With {.Source = New BitmapImage(New Uri("pack://application:,,,/Laila.Shell;component/Images/view16.png", UriKind.Absolute))}
                    }
                    For Each item In Shell.FolderViews
                        Dim viewSubMenuItem As MenuItem = New MenuItem() With {
                            .Header = item.Key,
                            .Icon = New Image() With {.Source = New BitmapImage(New Uri(item.Value.Item1, UriKind.Absolute))},
                            .Tag = "View:" & item.Key,
                            .IsCheckable = True,
                            .IsChecked = item.Key = Me.Folder.View
                        }
                        AddHandler viewSubMenuItem.Checked,
                            Sub(s2 As Object, e2 As EventArgs)
                                If Not _isCheckingInternally Then
                                    Me.Folder.View = item.Key
                                End If
                            End Sub
                        AddHandler viewSubMenuItem.Unchecked,
                            Sub(s2 As Object, e2 As EventArgs)
                                If Not _isCheckingInternally Then
                                    _isCheckingInternally = True
                                    viewSubMenuItem.IsChecked = True
                                    _isCheckingInternally = False
                                End If
                            End Sub
                        viewMenuItem.Items.Add(viewSubMenuItem)
                    Next
                    Me.ItemContextMenu.Items.Insert(0, viewMenuItem)
                    If Me.ItemContextMenu.Items.Count > 1 Then Me.ItemContextMenu.Items.Insert(1, New Separator())
                End If
                initializeViewMenu()

                Task.Run(AddressOf makeSortMenu)
            End If

            If Not Me.SelectedItems Is Nothing _
                OrElse Not EqualityComparer(Of IEnumerable(Of Item)).Default.Equals(_lastItems, Me.SelectedItems) Then
                releaseContextMenu()
                If Not IntPtr.Zero.Equals(_selectedItemshMenu) Then
                    Functions.DestroyMenu(_selectedItemshMenu)
                End If
                If Not _selectedItemsContextMenu Is Nothing Then
                    Marshal.ReleaseComObject(_selectedItemsContextMenu)
                    _selectedItemsContextMenu = Nothing
                End If

                Me.ItemContextMenu = getContextMenu(Me.Folder, Me.SelectedItems, Me.IsDefaultOnly)
                _selectedItemsContextMenu = _contextMenu
                _selectedItemshMenu = _hMenu
                _contextMenu = Nothing
                _hMenu = IntPtr.Zero

                If Not _contextMenu2 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu2)
                    _contextMenu2 = Nothing
                End If
                If Not _contextMenu3 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu3)
                    _contextMenu3 = Nothing
                End If

                _lastItems = Me.SelectedItems
                didGetMenu = True
            End If

            If didGetMenu Then
                Me.CanCut = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Me.ItemContextMenu.Buttons.Exists(
                        Function(b) b.Tag.ToString().Split(vbTab)(1) = "cut")
                Me.CanCopy = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Me.ItemContextMenu.Buttons.Exists(
                        Function(b) b.Tag.ToString().Split(vbTab)(1) = "copy")
                Me.CanPaste = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Me.ItemContextMenu.Buttons.Exists(
                        Function(b) b.Tag.ToString().Split(vbTab)(1) = "paste")
                Me.CanRename = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Me.ItemContextMenu.Buttons.Exists(
                        Function(b) b.Tag.ToString().Split(vbTab)(1) = "rename")
                Me.CanDelete = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Me.ItemContextMenu.Buttons.Exists(
                        Function(b) b.Tag.ToString().Split(vbTab)(1) = "delete")
                Me.CanShare = Not Me.ItemContextMenu Is Nothing _
                    AndAlso Not Me.ItemContextMenu.Items.Cast(Of Control).FirstOrDefault(
                        Function(c) TypeOf c Is MenuItem _
                            AndAlso Not CType(c, MenuItem).Tag Is Nothing _
                            AndAlso CType(c, MenuItem).Tag.ToString().Split(vbTab)(1) = "Windows.ModernShare") Is Nothing
            End If
        End Sub

        Private Async Function makeSortMenu() As Task
            Dim primaryProperties As List(Of String) = New List(Of String)()
            Dim additionalProperties As List(Of String) = New List(Of String)()
            Dim descriptions As Dictionary(Of String, String) = New Dictionary(Of String, String)()
            Dim folder As Folder

            UIHelper.OnUIThread(
                Sub()
                    folder = Me.Folder

                    Dim viewState As FolderViewState = FolderViewState.FromViewName(Me.Folder.FullPath)
                    If Not viewState Is Nothing AndAlso Not viewState.Columns Is Nothing AndAlso Not viewState.Columns.Count = 0 Then
                        For Each viewStateColumn In viewState.Columns.Where(Function(c) c.IsVisible)
                            Dim pKey As String = viewStateColumn.Name.Substring(viewStateColumn.Name.IndexOf("[") + 1)
                            pKey = pKey.Substring(0, pKey.IndexOf("]"))
                            Dim column As Column = Me.Folder.Columns.FirstOrDefault(Function(c) c.PROPERTYKEY.ToString() = pKey)
                            If Not column Is Nothing Then
                                primaryProperties.Add(column.PROPERTYKEY.ToString())
                                descriptions.Add(column.PROPERTYKEY.ToString(), column.DisplayName)
                            End If
                        Next
                    Else
                        For Each column In Me.Folder.Columns.Where(Function(c) c.IsVisible)
                            primaryProperties.Add(column.PROPERTYKEY.ToString())
                            descriptions.Add(column.PROPERTYKEY.ToString(), column.DisplayName)
                        Next
                    End If
                End Sub)

            Dim getProperties As Action(Of Item, PROPERTYKEY) =
                Sub(item As Item, pKey As PROPERTYKEY)
                    Dim text As String = item.PropertiesByKey(pKey).Text
                    If String.IsNullOrWhiteSpace(text) Then text = "prop:System.ItemTypeText;System.Size"
                    For Each propName In text.Substring(5).Split(";")
                        propName = propName.TrimStart("*")
                        Dim prop As [Property] = item.PropertiesByCanonicalName(propName)
                        If Not prop Is Nothing _
                                AndAlso Not primaryProperties.Contains(prop.Key.ToString()) _
                                AndAlso Not additionalProperties.Contains(prop.Key.ToString()) Then
                            additionalProperties.Add(prop.Key.ToString())
                            descriptions.Add(prop.Key.ToString(), prop.DescriptionDisplayName)
                        End If
                    Next
                End Sub
            Dim PKEY_System_InfoTipText As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 4
            }
            Dim PKEY_System_PropList_TileInfo As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 3
            }
            Dim uniqueTypes As List(Of String) = New List(Of String)()
            Dim items As List(Of Item) = Await folder.GetItemsAsync()
            UIHelper.OnUIThread(
                Sub()
                    For Each item In items
                        Dim itemTypeText As String = item.PropertiesByCanonicalName("System.ItemTypeText").Text
                        If Not uniqueTypes.Contains(itemTypeText) Then
                            uniqueTypes.Add(itemTypeText)
                            getProperties(item, PKEY_System_InfoTipText)
                            getProperties(item, PKEY_System_PropList_TileInfo)
                        End If
                    Next

                    Dim sortMenu As ContextMenu = New ContextMenu()
                    Dim moreMenuItem As MenuItem
                    Dim PKEY_System_ItemNameDisplay As String = Me.Folder.PropertiesByCanonicalName("System.ItemNameDisplay").Key.ToString()
                    Dim sortMenuItemCheckedAction As RoutedEventHandler = New RoutedEventHandler(
                        Sub(s2 As Object, e2 As RoutedEventArgs)
                            If Not _isCheckingInternally Then
                                If CType(s2, MenuItem).Tag.ToString().Substring(5).IndexOf(":") >= 0 Then
                                    Me.Folder.ItemsSortPropertyName =
                                    String.Format("PropertiesByKeyAsText[{0}].Text",
                                                  CType(s2, MenuItem).Tag.ToString().Substring(5))
                                Else
                                    Me.Folder.ItemsSortPropertyName = CType(s2, MenuItem).Tag.ToString().Substring(5)
                                End If
                            End If
                        End Sub)
                    Dim menuItemUncheckedAction As RoutedEventHandler = New RoutedEventHandler(
                        Sub(s2 As Object, e2 As RoutedEventArgs)
                            If Not _isCheckingInternally Then
                                _isCheckingInternally = True
                                CType(s2, MenuItem).IsChecked = True
                                _isCheckingInternally = False
                            End If
                        End Sub)
                    For Each propName In primaryProperties
                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = descriptions(propName),
                            .Tag = "Sort:" & If(propName = PKEY_System_ItemNameDisplay, "ItemNameDisplaySortValue", propName),
                            .IsCheckable = True
                        }
                        AddHandler menuItem.Checked, sortMenuItemCheckedAction
                        AddHandler menuItem.Unchecked, menuItemUncheckedAction
                        sortMenu.Items.Add(menuItem)
                    Next
                    If additionalProperties.Count > 0 Then
                        moreMenuItem = New MenuItem() With {.Header = "More", .Tag = "SortMore"}
                        sortMenu.Items.Add(moreMenuItem)
                        For Each propName In additionalProperties
                            Dim menuItem As MenuItem = New MenuItem() With {
                                .Header = descriptions(propName),
                                .Tag = "Sort:" & If(propName = PKEY_System_ItemNameDisplay, "ItemNameDisplaySortValue", propName),
                                .IsCheckable = True
                            }
                            AddHandler menuItem.Checked, sortMenuItemCheckedAction
                            AddHandler menuItem.Unchecked, menuItemUncheckedAction
                            moreMenuItem.Items.Add(menuItem)
                        Next
                    End If
                    sortMenu.Items.Add(New Separator())
                    Dim sortAscendingMenuItem As MenuItem = New MenuItem() With {
                        .Header = "Ascending",
                        .IsCheckable = True,
                        .Tag = "SortAscending"
                    }
                    sortMenu.Items.Add(sortAscendingMenuItem)
                    AddHandler sortAscendingMenuItem.Checked,
                        Sub(s2 As Object, e2 As EventArgs)
                            If Not _isCheckingInternally Then
                                Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
                            End If
                        End Sub
                    AddHandler sortAscendingMenuItem.Unchecked, menuItemUncheckedAction
                    Dim sortDescendingMenuItem As MenuItem = New MenuItem() With {
                        .Header = "Descending",
                        .IsCheckable = True,
                        .Tag = "SortDescending"
                    }
                    sortMenu.Items.Add(sortDescendingMenuItem)
                    AddHandler sortDescendingMenuItem.Checked,
                        Sub(s2 As Object, e2 As EventArgs)
                            If Not _isCheckingInternally Then
                                Me.Folder.ItemsSortDirection = ListSortDirection.Descending
                            End If
                        End Sub
                    AddHandler sortDescendingMenuItem.Unchecked, menuItemUncheckedAction
                    sortMenu.Items.Add(New Separator())
                    Dim groupByMenuItemCheckedAction As RoutedEventHandler = New RoutedEventHandler(
                        Sub(s2 As Object, e2 As RoutedEventArgs)
                            If Not _isCheckingInternally Then
                                Me.Folder.ItemsGroupByPropertyName =
                                    String.Format("PropertiesByKeyAsText[{0}].Text",
                                                  CType(s2, MenuItem).Tag.ToString().Substring(6))
                            End If
                        End Sub)
                    Dim groupByMoreMenuItem As MenuItem = New MenuItem() With {
                        .Header = "Group by",
                        .Tag = "GroupByMore"
                    }
                    sortMenu.Items.Add(groupByMoreMenuItem)
                    For Each propName In primaryProperties
                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = descriptions(propName),
                            .Tag = "Group:" & propName,
                            .IsCheckable = True
                        }
                        AddHandler menuItem.Checked, groupByMenuItemCheckedAction
                        AddHandler menuItem.Unchecked, menuItemUncheckedAction
                        groupByMoreMenuItem.Items.Add(menuItem)
                    Next
                    For Each propName In additionalProperties
                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = descriptions(propName),
                            .Tag = "Group:" & propName,
                            .IsCheckable = True
                        }
                        AddHandler menuItem.Checked, groupByMenuItemCheckedAction
                        AddHandler menuItem.Unchecked, menuItemUncheckedAction
                        groupByMoreMenuItem.Items.Add(menuItem)
                    Next
                    Dim groupByNoneMenuItem As MenuItem = New MenuItem() With {
                        .Header = "(None)",
                        .Tag = "Group:",
                        .IsCheckable = True
                    }
                    AddHandler groupByNoneMenuItem.Checked,
                        Sub(s2 As Object, e2 As EventArgs)
                            Me.Folder.ItemsGroupByPropertyName = Nothing
                        End Sub
                    AddHandler groupByNoneMenuItem.Unchecked, menuItemUncheckedAction
                    groupByMoreMenuItem.Items.Add(groupByNoneMenuItem)

                    Me.SortMenu = sortMenu
                    initializeSortMenu()
                End Sub)
        End Function

        Private Sub initializeSortMenu()
            If Not Me.SortMenu Is Nothing Then
                _isCheckingInternally = True
                Dim sortPropertyName As String = Me.Folder.ItemsSortPropertyName
                Dim groupByPropertyName As String = Me.Folder.ItemsGroupByPropertyName
                If Not String.IsNullOrWhiteSpace(sortPropertyName) Then
                    If sortPropertyName.IndexOf("[") >= 0 Then
                        sortPropertyName = sortPropertyName.Substring(sortPropertyName.IndexOf("[") + 1)
                        sortPropertyName = sortPropertyName.Substring(0, sortPropertyName.IndexOf("]"))
                    End If
                End If
                If Not String.IsNullOrWhiteSpace(groupByPropertyName) Then
                    If groupByPropertyName.IndexOf("[") >= 0 Then
                        groupByPropertyName = groupByPropertyName.Substring(groupByPropertyName.IndexOf("[") + 1)
                        groupByPropertyName = groupByPropertyName.Substring(0, groupByPropertyName.IndexOf("]"))
                    End If
                End If
                Dim sortMoreMenuItem As MenuItem = Nothing
                Dim groupByMoreMenuItem As MenuItem = Nothing
                For Each item In Me.SortMenu.Items
                    If TypeOf item Is MenuItem Then
                        If CType(item, MenuItem).Tag.ToString().StartsWith("Sort:") Then
                            CType(item, MenuItem).IsChecked = CType(item, MenuItem).Tag.ToString() = "Sort:" & sortPropertyName
                        ElseIf CType(item, MenuItem).Tag.ToString() = "SortMore" Then
                            sortMoreMenuItem = item
                        ElseIf CType(item, MenuItem).Tag.ToString() = "GroupByMore" Then
                            groupByMoreMenuItem = item
                        ElseIf CType(item, MenuItem).Tag.ToString() = "SortAscending" Then
                            CType(item, MenuItem).IsChecked = Me.Folder.ItemsSortDirection = ListSortDirection.Ascending
                        ElseIf CType(item, MenuItem).Tag.ToString() = "SortDescending" Then
                            CType(item, MenuItem).IsChecked = Me.Folder.ItemsSortDirection = ListSortDirection.Descending
                        End If
                    End If
                Next
                If Not sortMoreMenuItem Is Nothing Then
                    For Each item In sortMoreMenuItem.Items
                        If TypeOf item Is MenuItem AndAlso CType(item, MenuItem).Tag.ToString().StartsWith("Sort:") Then
                            CType(item, MenuItem).IsChecked = CType(item, MenuItem).Tag.ToString() = "Sort:" & sortPropertyName
                        End If
                    Next
                End If
                If Not groupByMoreMenuItem Is Nothing Then
                    For Each item In groupByMoreMenuItem.Items
                        If TypeOf item Is MenuItem AndAlso CType(item, MenuItem).Tag.ToString().StartsWith("Group:") Then
                            If CType(item, MenuItem).Tag.ToString() = "Group:" Then
                                CType(item, MenuItem).Visibility = If(Not String.IsNullOrWhiteSpace(groupByPropertyName), Visibility.Visible, Visibility.Collapsed)
                                CType(item, MenuItem).IsChecked = False
                            Else
                                CType(item, MenuItem).IsChecked = CType(item, MenuItem).Tag.ToString() = "Group:" & groupByPropertyName
                            End If
                        End If
                    Next
                End If
                _isCheckingInternally = False
            End If
        End Sub

        Private Sub initializeViewMenu()
            If Not Me.ViewMenu Is Nothing AndAlso Not Me.Folder Is Nothing Then
                _isCheckingInternally = True
                For Each item In Me.ViewMenu.Items
                    If TypeOf item Is MenuItem AndAlso CType(item, MenuItem).Tag.ToString().StartsWith("View:") Then
                        CType(item, MenuItem).IsChecked = CType(item, MenuItem).Tag.ToString().Substring(5) = Me.Folder.View
                    End If
                Next
                _isCheckingInternally = False
            End If
        End Sub

        Private Sub folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "ItemsSortPropertyName", "ItemsSortDirection", "ItemsGroupByPropertyName"
                    initializeSortMenu()
                Case "View"
                    initializeViewMenu()
            End Select
        End Sub

        Private Sub onSelectionChanged()
            If Not _updateTimer Is Nothing Then
                _updateTimer.Dispose()
            End If

            If Me.UpdateDelay.HasValue AndAlso Me.UpdateDelay > 0 Then
                _updateTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                'Me.Update()

                                If Not _updateTimer Is Nothing Then
                                    _updateTimer.Dispose()
                                    _updateTimer = Nothing
                                End If
                            End Sub)
                    End Sub), Nothing, Me.UpdateDelay.Value, Timeout.Infinite)
            ElseIf Me.UpdateDelay.HasValue Then
                Me.Update()
            End If
        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim icm As Menus = TryCast(d, Menus)
            If Not e.OldValue Is Nothing Then
                RemoveHandler CType(e.OldValue, Folder).PropertyChanged, AddressOf icm.folder_PropertyChanged
            End If
            If Not e.NewValue Is Nothing Then
                AddHandler CType(e.NewValue, Folder).PropertyChanged, AddressOf icm.folder_PropertyChanged
            End If
            icm.onSelectionChanged()
        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim icm As Menus = TryCast(d, Menus)
            icm.onSelectionChanged()
        End Sub

        Shared Function OnSelectedItemsCoerce(ByVal d As DependencyObject, baseValue As Object) As Object
            If Not baseValue Is Nothing Then
                Dim items As IEnumerable(Of Item) = baseValue
                If items.Count = 0 Then
                    Return Nothing
                Else
                    Return baseValue
                End If
            Else
                Return baseValue
            End If
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                releaseContextMenuFull()
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
End Namespace