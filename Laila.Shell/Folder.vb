Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media.Imaging
Imports Accessibility
Imports Microsoft.Xaml.Behaviors

Public Class Folder
    Inherits Item

    Private _columns As List(Of Column)
    Private _items As ObservableCollection(Of Item)
    Private _firstContextMenuCall As Boolean = True
    Private _itemsLock As Object = New Object()
    Friend _shellFolder As IShellFolder
    Private _columnManager As IColumnManager

    Public Shared Function FromKnownFolderGuid(knownFolderGuid As Guid, setIsLoadingAction As Action(Of Boolean)) As Folder
        Return FromParsingName("shell:::" & knownFolderGuid.ToString("B"), Nothing, setIsLoadingAction)
    End Function

    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim ptr2 As IntPtr
        shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, ptr2)
        Return Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IShellFolder))
    End Function

    Friend Shared Function GetIShellFolderFromPidl(pidl As IntPtr, bindingParent As Folder) As IShellFolder
        Dim ptr As IntPtr
        Try
            bindingParent._shellFolder.BindToObject(pidl, IntPtr.Zero, Guids.IID_IShellFolder, ptr)
            Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellFolder))
        Finally
            Marshal.Release(ptr)
        End Try
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(shellItem2, logicalParent, setIsLoadingAction)

        If Not shellItem2 Is Nothing Then
            _shellFolder = Folder.GetIShellFolderFromIShellItem2(shellItem2)
        End If
    End Sub

    Public ReadOnly Property Columns(canonicalName As String) As Column
        Get
            Return Me.Columns.SingleOrDefault(Function(c) canonicalName.Equals(c.CanonicalName))
        End Get
    End Property

    Public ReadOnly Property Columns As IEnumerable(Of Column)
        Get
            If _columns Is Nothing Then
                _columns = New List(Of Column)()

                Dim ptr As IntPtr
                _shellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
                Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
                _columnManager = shellView
                Dim count As Integer
                _columnManager.GetColumnCount(CM_ENUM_FLAGS.CM_ENUM_ALL, count)
                Dim propertyKeys(count - 1) As PROPERTYKEY
                _columnManager.GetColumns(CM_ENUM_FLAGS.CM_ENUM_ALL, propertyKeys, count)
                Dim index As Integer = 0
                For Each propertyKey In propertyKeys
                    Dim col As Column = New Column(propertyKey, _columnManager, index)
                    If Not col._propertyDescription Is Nothing Then
                        _columns.Add(col)
                        index += 1
                    End If
                Next
            End If

            Return _columns
        End Get
    End Property

    Public Overridable ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
        Get
            If _items Is Nothing Then
                If Not _setIsLoadingAction Is Nothing Then
                    _setIsLoadingAction(True)
                End If

                Dim result As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()

                Dim t As Thread = New Thread(New ThreadStart(
                    Sub()
                        SyncLock _itemsLock
                            If _items Is Nothing Then
                                updateItems(result)
                                Me.Items = result
                            End If
                        End SyncLock

                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(False)
                        End If
                    End Sub))

                t.Start()

                Return New ObservableCollection(Of Item)()
            Else
                Return _items
            End If
        End Get
    End Property

    Public Overridable Property Items As ObservableCollection(Of Item)
        Get
            SyncLock _itemsLock
                If _items Is Nothing Then
                    Dim result As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()
                    updateItems(result)
                    Me.Items = result
                End If

                Return _items
            End SyncLock
        End Get
        Set(value As ObservableCollection(Of Item))
            SetValue(_items, value)
            Me.NotifyOfPropertyChange("ItemsThreaded")
        End Set
    End Property

    Protected Sub updateItems(items As ObservableCollection(Of Item))
        updateItems(SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, True,
                    Function(item As Item) As Boolean
                        Return Not items.FirstOrDefault(Function(i) i.FullPath = item.FullPath) Is Nothing
                    End Function,
                    Sub(item As Item)
                        items.Add(item)
                    End Sub,
                    Sub(item As Item)
                        items.Remove(item)
                    End Sub,
                    Function(paths As List(Of String)) As List(Of Item)
                        Return items.Where(Function(i) Not paths.Contains(i.FullPath)).ToList()
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Folder(shellItem2, Me, _setIsLoadingAction)
                    End Function,
                    0)
    End Sub

    Protected Sub updateItems(flags As UInt32, condition As Boolean,
                              exists As Func(Of Item, Boolean), add As Action(Of Item), remove As Action(Of Item),
                              getToBeRemoved As Func(Of List(Of String), List(Of Item)),
                              makeNewFolder As Func(Of IShellItem2, Item), uiHelp As Integer)
        Dim paths As List(Of String) = New List(Of String)

        Dim attr As SFGAO = SFGAO.HASSUBFOLDER Or SFGAO.ISSLOW
        _shellItem2.GetAttributes(attr, attr)

        If Not condition AndAlso (attr.HasFlag(SFGAO.ISSLOW) OrElse Me.FullPath.StartsWith("\\")) AndAlso attr.HasFlag(SFGAO.HASSUBFOLDER) Then
            Application.Current.Dispatcher.Invoke(
                Sub()
                    add(New DummyTreeViewFolder("Loading..."))
                End Sub)
        Else
            If Not isWindows7OrLower() Then
                Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
                Functions.CreateBindCtx(0, bindCtxPtr)
                bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

                If Not bindCtx Is Nothing AndAlso Not IntPtr.Zero.Equals(bindCtxPtr) Then
                    Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
                    Try
                        Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
                        propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))
                    Finally
                        If Not IntPtr.Zero.Equals(propertyBagPtr) Then
                            Marshal.Release(bindCtxPtr)
                        End If
                    End Try

                    If Not propertyBag Is Nothing Then
                        Dim var As PROPVARIANT
                        Dim enumShellItems As IEnumShellItems
                        Try
                            var.vt = VarEnum.VT_UI4
                            var.union.uintVal = flags
                            propertyBag.Write("SHCONTF", var)

                            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag)

                            Dim ptr2 As IntPtr
                            Try
                                _shellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                                enumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))
                            Finally
                                If Not IntPtr.Zero.Equals(ptr2) Then Marshal.Release(ptr2)
                            End Try

                            If Not enumShellItems Is Nothing Then
                                Dim shellItemArray(0) As IShellItem, fetched As UInt32 = 1
                                enumShellItems.Next(1, shellItemArray, fetched)
                                Dim isOnce As Boolean = False
                                While fetched = 1 And (Not isOnce OrElse condition)
                                    Dim attr2 As Integer = SFGAO.FOLDER
                                    shellItemArray(0).GetAttributes(attr2, attr2)
                                    Dim fullPath As String = Item.GetFullPathFromShellItem2(shellItemArray(0))
                                    Application.Current.Dispatcher.Invoke(
                                        Sub()
                                            Dim newItem As Item
                                            If CBool(attr2 And SFGAO.FOLDER) Then
                                                newItem = makeNewFolder(shellItemArray(0))
                                            Else
                                                newItem = New Item(shellItemArray(0), Nothing, Nothing)
                                            End If
                                            paths.Add(newItem.FullPath)
                                            If Not exists(newItem) Then
                                                add(newItem)
                                            Else
                                                newItem.Dispose()
                                            End If
                                        End Sub)
                                    enumShellItems.Next(1, shellItemArray, fetched)
                                    isOnce = True
                                    If uiHelp > 0 Then Thread.Sleep(uiHelp)
                                End While
                            End If
                        Catch ex As Exception
                            Application.Current.Dispatcher.Invoke(
                                Sub()
                                    add(New DummyTreeViewFolder(ex.Message))
                                End Sub)
                        Finally
                            If Not enumShellItems Is Nothing Then
                                Marshal.ReleaseComObject(enumShellItems)
                            End If
                            Marshal.ReleaseComObject(bindCtx)
                            Marshal.Release(bindCtxPtr)
                            Marshal.ReleaseComObject(propertyBag)
                            var.Dispose()
                        End Try
                    End If
                End If
            Else
                Dim list As IEnumIDList
                _shellFolder.EnumObjects(Nothing, flags, list)
                If Not list Is Nothing Then
                    Dim pidl(0) As IntPtr, fetched As Integer
                    While list.Next(1, pidl, fetched) = 0
                        Dim attr2 As Integer = SFGAO.FOLDER
                        _shellFolder.GetAttributesOf(1, pidl, attr2)
                        Dim shellItem2 As IShellItem2 = Item.GetIShellItem2FromPidl(pidl(0), Me)
                        Dim path As String = Item.GetFullPathFromShellItem2(shellItem2)
                        Dim newItem As Item
                        If CBool(attr2 And SFGAO.FOLDER) Then
                            newItem = makeNewFolder(shellItem2)
                        Else
                            newItem = New Item(shellItem2, Me, _setIsLoadingAction)
                        End If
                        paths.Add(newItem.FullPath)
                        If Not exists(newItem) Then
                            add(newItem)
                        Else
                            newItem.Dispose()
                        End If
                    End While
                End If
            End If

            Application.Current.Dispatcher.Invoke(
              Sub()
                  For Each item In getToBeRemoved(paths)
                      remove(item)
                  Next
              End Sub)
        End If
    End Sub

    Public Function GetContextMenu(items As IEnumerable(Of Item), ByRef contextMenu As IContextMenu,
                                   ByRef defaultId As String, isDefaultOnly As Boolean) As ContextMenu
        Dim contextMenu2 As IContextMenu2, contextMenu3 As IContextMenu3, hMenu As IntPtr
        contextMenu = getContextMenu(items, contextMenu2, contextMenu3, hMenu, isDefaultOnly)
        Dim defaultIdLocal As String, contextMenuLocal As IContextMenu = contextMenu
        Dim getMenu As Func(Of IntPtr, List(Of Control)) =
            Function(hMenu2 As IntPtr) As List(Of Control)
                If Not contextMenu2 Is Nothing Then
                    contextMenu2.HandleMenuMsg(WM.INITMENU, hMenu2, IntPtr.Zero)
                    contextMenu2.HandleMenuMsg(WM.INITMENUPOPUP, hMenu2, IntPtr.Zero)
                End If
                If Not contextMenu3 Is Nothing Then
                    Dim ptr3 As IntPtr
                    contextMenu3.HandleMenuMsg(WM.INITMENU, hMenu2, IntPtr.Zero)
                    contextMenu3.HandleMenuMsg2(WM.INITMENUPOPUP, hMenu2, IntPtr.Zero, ptr3)
                End If

                Dim result As List(Of Control) = New List(Of Control)()

                For i = 0 To Functions.GetMenuItemCount(hMenu2) - 1
                    Dim mii As MENUITEMINFO
                    mii.cbSize = CUInt(Marshal.SizeOf(mii))
                    mii.fMask = MIIM.MIIM_STRING
                    mii.dwTypeData = New String(" "c, 2048)
                    mii.cch = mii.dwTypeData.Length
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

                        Dim cmd As String, id As Integer
                        If mii.wID < 9999999 Then
                            id = mii.wID
                            Dim bytes(256) As Byte
                            contextMenuLocal.GetCommandString(mii.wID, GCS.VERBA, 0, bytes, 256)
                            cmd = Text.Encoding.ASCII.GetString(bytes).Trim(vbNullChar)
                        End If

                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = header.Replace("&", "_"),
                            .Icon = New Image() With {.Source = bitmapSource},
                            .Tag = id & vbTab & cmd,
                            .IsEnabled = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DISABLED), False, True),
                            .FontWeight = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
                        }

                        If CBool(mii.fState And MFS.MFS_DEFAULT) Then
                            defaultIdLocal = menuItem.Tag
                        End If

                        If mii.hSubMenu Then
                            Dim subMenu As List(Of Control) = getMenu(mii.hSubMenu)
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

        ' make our own menu
        Dim menu As ContextMenu = New ContextMenu()
        For Each item In getMenu(hMenu)
            menu.Items.Add(item)
        Next
        AddHandler menu.Closed,
            Sub(s As Object, e As RoutedEventArgs)
                Functions.DestroyMenu(hMenu)
            End Sub

        ' folder background menu
        If items Is Nothing OrElse items.Count = 0 Then
            ' remove some items that don't work anyway
            For i = 1 To 7
                menu.Items.RemoveAt(0)
            Next

            Dim menuItemWindowsShare As MenuItem = menu.Items.Cast(Of Control) _
                .FirstOrDefault(Function(c) Not c.Tag Is Nothing AndAlso c.Tag.ToString().Split(vbTab)(1) = "Windows.Share")
            Dim menuItemWindowsShareIndex As Integer = menu.Items.IndexOf(menuItemWindowsShare)
            menu.Items.RemoveAt(menuItemWindowsShareIndex - 1)
            menu.Items.RemoveAt(menuItemWindowsShareIndex - 1)

            ' update paste item
            menu.Items(0).Tag = vbTab & "paste"
            menu.Items(0).IsEnabled = Clipboard.ContainsFileDropList()
        End If

        defaultId = defaultIdLocal

        Return menu
    End Function

    Public Sub InvokeCommand(contextMenu As IContextMenu, items As IEnumerable(Of Item), id As String)
        Dim cmi As New CMInvokeCommandInfoEx
        Debug.WriteLine("InvokeCommand " & id)
        If id.Split(vbTab)(1).Length <> 0 Then
            cmi.lpVerb = Marshal.StringToHGlobalAnsi(id.Split(vbTab)(1))
            cmi.lpVerbW = Marshal.StringToHGlobalUni(id.Split(vbTab)(1))
        Else
            cmi.lpVerb = New IntPtr(Convert.ToUInt32(id.Split(vbTab)(0)))
            cmi.lpVerbW = New IntPtr(Convert.ToUInt32(id.Split(vbTab)(0)))
        End If
        cmi.lpDirectory = Me.FullPath
        cmi.lpDirectoryW = Me.FullPath
        cmi.fMask = CMIC.UNICODE Or CMIC.PTINVOKE
        cmi.nShow = SW.SHOWNORMAL
        cmi.hwnd = Shell._hwnd
        cmi.cbSize = CUInt(Marshal.SizeOf(cmi))
        cmi.ptInvoke = New Drawing.Point(0, 0)

        Dim h As HResult = contextMenu.InvokeCommand(cmi)
        Debug.WriteLine("InvokeCommand returned " & h.ToString())
    End Sub

    Private Function getContextMenu(items As IEnumerable(Of Item),
                                    ByRef contextMenu2 As IContextMenu2, ByRef contextMenu3 As IContextMenu3,
                                    ByRef hMenu As IntPtr, isDefaultOnly As Boolean) As IContextMenu
        Dim folderpidl As IntPtr, shellItemPtr As IntPtr
        shellItemPtr = Marshal.GetIUnknownForObject(_shellItem2)
        Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)

        Dim contextMenu As IContextMenu
        Dim pidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr
        Dim lastpidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr

        If Not items Is Nothing AndAlso items.Count > 0 Then
            For i = 0 To items.Count - 1
                shellItemPtr = Marshal.GetIUnknownForObject(items(i)._shellItem2)
                Functions.SHGetIDListFromObject(shellItemPtr, pidls(i))
                lastpidls(i) = Functions.ILFindLastID(pidls(i))
            Next

            Dim ptr As IntPtr
            Try
                _shellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptr)
                contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
            End Try

            Dim shellExtInitPtr As IntPtr, shellExtInit As IShellExtInit, dataObject As IDataObject
            Try
                Marshal.QueryInterface(ptr, GetType(IShellExtInit).GUID, shellExtInitPtr)
                If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
                    shellExtInit = Marshal.GetObjectForIUnknown(shellExtInitPtr)
                    Functions.SHCreateDataObject(folderpidl, lastpidls.Count, lastpidls, IntPtr.Zero, GetType(IDataObject).GUID, dataObject)
                    shellExtInit.Initialize(folderpidl, dataObject, IntPtr.Zero)
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
        Else
            Dim shellView As IShellView, ptr As IntPtr
            Try
                _shellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
                shellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
                shellView.GetItemObject(SVGIO.SVGIO_BACKGROUND, GetType(IContextMenu).GUID, ptr)
                contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
                If Not shellView Is Nothing Then
                    Marshal.ReleaseComObject(shellView)
                End If
            End Try
        End If

        Dim ptrContextMenu As IntPtr, ptr2, ptr3
        Try
            ptrContextMenu = Marshal.GetIUnknownForObject(contextMenu)
            Marshal.QueryInterface(ptrContextMenu, GetType(IContextMenu2).GUID, ptr2)
            If Not IntPtr.Zero.Equals(ptr2) Then
                contextMenu2 = Marshal.GetObjectForIUnknown(ptr2)
            End If
            Marshal.QueryInterface(ptrContextMenu, GetType(IContextMenu3).GUID, ptr3)
            If Not IntPtr.Zero.Equals(ptr3) Then
                contextMenu3 = Marshal.GetObjectForIUnknown(ptr3)
            End If
        Finally
            If Not IntPtr.Zero.Equals(ptr2) Then
                Marshal.Release(ptr2)
            End If
            If Not IntPtr.Zero.Equals(ptr3) Then
                Marshal.Release(ptr3)
            End If
            If Not IntPtr.Zero.Equals(ptrContextMenu) Then
                Marshal.Release(ptrContextMenu)
            End If
        End Try

        hMenu = Functions.CreatePopupMenu()
        Dim flags As Integer = CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS
        If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
        contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, flags)

        If _firstContextMenuCall Then
            ' somehow very first call doesn't return all items
            Functions.DestroyMenu(hMenu)
            hMenu = Functions.CreatePopupMenu()
            contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, flags)
            _firstContextMenuCall = False
        End If

        Return contextMenu
    End Function

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        Select Case e.Event
            Case SHCNE.CREATE
                If Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    Dim parentShellItem2 As IShellItem2
                    Try
                        Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                        If Not item1 Is Nothing Then
                            item1.GetParent(parentShellItem2)
                            Dim parentFullPath As String
                            parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                            If Me.FullPath.Equals(parentFullPath) Then
                                SyncLock _itemsLock
                                    If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                        _items.Add(New Item(item1, Me, _setIsLoadingAction))
                                        Dim view As ICollectionView = CollectionViewSource.GetDefaultView(_items)
                                        view.Refresh()
                                    End If
                                End SyncLock
                            End If
                        End If
                    Finally
                        If Not parentShellItem2 Is Nothing Then
                            Marshal.ReleaseComObject(parentShellItem2)
                        End If
                    End Try
                End If
            Case SHCNE.MKDIR
                If Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    Dim parentShellItem2 As IShellItem2
                    Try
                        Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                        If Not item1 Is Nothing Then
                            item1.GetParent(parentShellItem2)
                            Dim parentFullPath As String
                            parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                            If Me.FullPath.Equals(parentFullPath) Then
                                SyncLock _itemsLock
                                    If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                        _items.Add(New Folder(item1, Me, _setIsLoadingAction))
                                    End If
                                End SyncLock
                            End If
                        End If
                    Finally
                        If Not parentShellItem2 Is Nothing Then
                            Marshal.ReleaseComObject(parentShellItem2)
                        End If
                    End Try
                End If
            Case SHCNE.RMDIR, SHCNE.DELETE
                SyncLock _itemsLock
                    If Not _items Is Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                        Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                        If Not item Is Nothing Then
                            _items.Remove(item)
                        End If
                    End If
                End SyncLock
            Case SHCNE.DRIVEADD
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    SyncLock _itemsLock
                        If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                            If Not item1 Is Nothing Then
                                _items.Add(New Folder(item1, Me, _setIsLoadingAction))
                            End If
                        End If
                    End SyncLock
                End If
            Case SHCNE.DRIVEREMOVED
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    SyncLock _itemsLock
                        Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                        If Not item Is Nothing Then
                            _items.Remove(item)
                        End If
                    End SyncLock
                End If
            Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _items Is Nothing Then
                    If Not _setIsLoadingAction Is Nothing Then
                        _setIsLoadingAction(True)
                    End If

                    Dim thread As Thread = New Thread(New ThreadStart(
                        Sub()
                            updateItems(_items)

                            Application.Current.Dispatcher.Invoke(
                                Sub()
                                    Dim view As ICollectionView = CollectionViewSource.GetDefaultView(_items)
                                    view.Refresh()
                                End Sub)

                            If Not _setIsLoadingAction Is Nothing Then
                                _setIsLoadingAction(False)
                            End If
                        End Sub))

                    thread.Start()
                End If
        End Select
    End Sub

    Protected Function isWindows7OrLower() As Boolean
        Dim osVersion As Version = Environment.OSVersion.Version
        ' Windows 7 has version number 6.1
        Return osVersion.Major < 6 OrElse (osVersion.Major = 6 AndAlso osVersion.Minor <= 1)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If Not _items Is Nothing Then
                For Each item In _items
                    item.Dispose()
                Next
            End If
        End If

        If Not _shellFolder Is Nothing Then
            Marshal.ReleaseComObject(_shellFolder)
        End If
        If Not _columnManager Is Nothing Then
            Marshal.ReleaseComObject(_columnManager)
        End If

        MyBase.Dispose(disposing)
    End Sub
End Class
