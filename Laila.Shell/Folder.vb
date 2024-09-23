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

Public Class Folder
    Inherits Item

    Private _interface As IShellFolder
    Private _columns As List(Of Column)
    Private _items As ObservableCollection(Of Item)
    Private _firstContextMenuCall As Boolean = True
    Private _itemsLock As Object = New Object()

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
        bindingParent.ShellFolder.BindToObject(pidl, IntPtr.Zero, Guids.IID_IShellFolder, ptr)
        Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellFolder))
    End Function

    Public Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(shellItem2, logicalParent, setIsLoadingAction)

        _interface = shellFolder
    End Sub

    Public Sub New(bindingParent As Folder, pidl As IntPtr, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(bindingParent, pidl, logicalParent, setIsLoadingAction)
    End Sub

    Friend Overloads ReadOnly Property ShellFolder As IShellFolder
        Get
            If _interface Is Nothing Then
                If _pidl.Equals(IntPtr.Zero) Then
                    Functions.SHGetDesktopFolder(_interface)
                Else
                    _interface = GetIShellFolderFromPidl(_pidl, _bindingParent)
                    Marshal.FreeCoTaskMem(_pidl)
                End If
            End If

            Return _interface
        End Get
    End Property

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
                Me.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
                Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
                Dim columnManager As IColumnManager = shellView
                Dim count As Integer
                columnManager.GetColumnCount(CM_ENUM_FLAGS.CM_ENUM_ALL, count)
                Dim propertyKeys(count - 1) As PROPERTYKEY
                columnManager.GetColumns(CM_ENUM_FLAGS.CM_ENUM_ALL, propertyKeys, count)
                Dim index As Integer = 0
                For Each propertyKey In propertyKeys
                    _columns.Add(New Column(propertyKey, columnManager,
                                            index))
                    index += 1
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

    Private Sub updateItems(items As ObservableCollection(Of Item))
        Dim paths As List(Of String) = New List(Of String)

        Dim flags As UInt32 = CType(SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, UInt32)

        If Not isWindows7OrLower() Then
            Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
            Functions.CreateBindCtx(0, bindCtxPtr)
            bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

            Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
            propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))

            Dim var As New PROPVARIANT()
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var)

            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag)
            Dim ptr2 As IntPtr
            bindCtxPtr = Marshal.GetIUnknownForObject(bindCtx)

            ShellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
            Dim enumShellItems As IEnumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))

            Try
                Dim shellItemArray(0) As IShellItem, feteched As UInt32 = 1
                Application.Current.Dispatcher.Invoke(
                Sub()
                    enumShellItems.Next(1, shellItemArray, feteched)
                End Sub)
                While feteched = 1
                    Dim attr As Integer = SFGAO.FOLDER
                    shellItemArray(0).GetAttributes(attr, attr)
                    Dim fullPath As String
                    CType(shellItemArray(0), IShellItem2).GetDisplayName(SHGDN.FORPARSING, fullPath)
                    paths.Add(fullPath)
                    Application.Current.Dispatcher.Invoke(
                    Sub()
                        If items.FirstOrDefault(Function(i) i.FullPath = fullPath) Is Nothing Then
                            If CBool(attr And SFGAO.FOLDER) Then
                                items.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(shellItemArray(0)), shellItemArray(0), Me, _setIsLoadingAction))
                            Else
                                items.Add(New Item(shellItemArray(0), Me, _setIsLoadingAction))
                            End If
                        End If
                    End Sub)
                    Application.Current.Dispatcher.Invoke(
                    Sub()
                        enumShellItems.Next(1, shellItemArray, feteched)
                    End Sub)
                End While
            Catch ex As Exception
                ' directories getting deleted while being enumerated
            End Try
        Else
            Dim list As IEnumIDList
            Me.ShellFolder.EnumObjects(Nothing, flags, list)
            If Not list Is Nothing Then
                Dim pidl(0) As IntPtr, fetched As Integer
                While list.Next(1, pidl, fetched) = 0
                    Dim attr As Integer = SFGAO.FOLDER
                    Me.ShellFolder.GetAttributesOf(1, pidl, attr)
                    Dim newItem As Item
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                            If CBool(attr And SFGAO.FOLDER) Then
                                newItem = New Folder(Me, pidl(0), Me, _setIsLoadingAction)
                            Else
                                newItem = New Item(Me, pidl(0), Me, _setIsLoadingAction)
                            End If
                            paths.Add(newItem.FullPath)
                            If items.FirstOrDefault(Function(i) i.FullPath = newItem.FullPath) Is Nothing Then
                                items.Add(newItem)
                            Else
                                newItem.Dispose()
                            End If
                        End Sub)
                End While
            End If
        End If

        Application.Current.Dispatcher.Invoke(
          Sub()
              Dim toBeRemoved As List(Of Item) = items.Where(Function(i) Not paths.Contains(i.FullPath)).ToList()
              For Each item In toBeRemoved
                  items.Remove(item)
              Next
          End Sub)
    End Sub

    Public Function GetContextMenu(items As IEnumerable(Of Item), ByRef contextMenu As IContextMenu, ByRef defaultId As String, isDefaultOnly As Boolean) As ContextMenu
        Dim hMenu As IntPtr, contextMenu2 As IContextMenu2, contextMenu3 As IContextMenu3
        contextMenu = getIContextMenu(items, contextMenu2, contextMenu3, hMenu, isDefaultOnly)
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
                        mii.fMask = MIIM.MIIM_ID
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim cmd As String
                        If mii.wID < 9999999 Then
                            Dim bytes(256) As Byte
                            contextMenuLocal.GetCommandString(mii.wID, GCS.VERBA, 0, bytes, 256)
                            cmd = Text.Encoding.ASCII.GetString(bytes).Trim(vbNullChar)
                        End If

                        mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = header.Replace("&", "_"),
                            .Icon = New Image() With {.Source = bitmapSource},
                            .Tag = mii.wID & vbTab & cmd,
                            .IsEnabled = If(mii.fState.HasFlag(MFS.MFS_DISABLED), False, True),
                            .FontWeight = If(mii.fState.HasFlag(MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
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

                Functions.DestroyMenu(hMenu2)

                Return result
            End Function

        Dim menu As ContextMenu = New ContextMenu()
        For Each item In getMenu(hMenu)
            menu.Items.Add(item)
        Next

        ' folder background menu
        If items Is Nothing OrElse items.Count = 0 Then
            ' remove some items that don't work anyway
            For i = 1 To 9
                menu.Items.RemoveAt(0)
            Next
            menu.Items.Remove(menu.Items.Cast(Of Control) _
                .FirstOrDefault(Function(c) Not c.Tag Is Nothing AndAlso c.Tag.ToString().Split(vbTab)(1) = "Windows.Share"))

            ' add some of our own
            menu.Items.Insert(0, New MenuItem() With {.Header = "Paste", .Tag = vbTab & "paste", .IsEnabled = Clipboard.ContainsFileDropList()})
        End If

        defaultId = defaultIdLocal

        Return menu
    End Function

    Public Sub InvokeCommand(contextMenu As IContextMenu, items As IEnumerable(Of Item), id As String)
        Dim cmi As New CMINVOKECOMMANDINFOEX
        Debug.WriteLine("InvokeCommand " & id)
        If id.Split(vbTab)(1).Length <> 0 Then
            cmi.lpVerb = Marshal.StringToHGlobalAnsi(id.Split(vbTab)(1))
            cmi.lpVerbW = Marshal.StringToHGlobalAnsi(id.Split(vbTab)(1))
        Else
            cmi.lpVerb = Convert.ToUInt32(id.Split(vbTab)(0))
            cmi.lpVerbW = Convert.ToUInt32(id.Split(vbTab)(0))
        End If
        cmi.lpDirectory = Me.FullPath
        cmi.lpDirectoryW = Me.FullPath
        cmi.fMask = CMIC.NOZONECHECKS Or CMIC.ASYNCOK
        cmi.nShow = SW.SHOWNORMAL
        cmi.hwnd = IntPtr.Zero
        cmi.cbSize = CUInt(Marshal.SizeOf(cmi))

        contextMenu.InvokeCommand(cmi)
    End Sub

    Public Sub PasteFiles()
        ' Check if the clipboard contains file drop data
        If Clipboard.ContainsFileDropList() Then
            ' Get the list of files from the clipboard
            Dim files As StringCollection = Clipboard.GetFileDropList()

            ' Iterate through the files and move or copy them to the destination path
            For Each fileFullPath As String In files
                Dim fileName As String = Path.GetFileName(fileFullPath)
                Dim destFile As String = Path.Combine(Me.FullPath, fileName)

                ' Check if the file was cut or copied
                If Clipboard.GetData("Preferred DropEffect") = DragDropEffects.Move Then
                    ' Move the file
                    File.Move(fileFullPath, destFile)
                Else
                    ' Copy the file
                    File.Copy(fileFullPath, destFile)
                End If
            Next
        End If
    End Sub
    Private Function getIContextMenu(items As IEnumerable(Of Item),
                                     ByRef contextMenu2 As IContextMenu2, ByRef contextMenu3 As IContextMenu3,
                                     ByRef hMenu As IntPtr, isDefaultOnly As Boolean) As IContextMenu
        Dim contextMenu As IContextMenu, folderpidl As IntPtr, shellItemPtr As IntPtr = Marshal.GetIUnknownForObject(Me.ShellItem2)
        Dim pidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr
        Dim lastpidls(If(items Is Nothing OrElse items.Count = 0, -1, items.Count - 1)) As IntPtr
        Dim ptr As IntPtr

        Functions.SHGetIDListFromObject(shellItemPtr, folderpidl)

        If Not items Is Nothing AndAlso items.Count > 0 Then
            For i = 0 To items.Count - 1
                Functions.SHGetIDListFromObject(Marshal.GetIUnknownForObject(items(i).ShellItem2), pidls(i))
                lastpidls(i) = Functions.ILFindLastID(pidls(i))
            Next

            Me.ShellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptr)
            contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))

            Dim shellExtInitPtr As IntPtr
            Marshal.QueryInterface(ptr, GetType(IShellExtInit).GUID, shellExtInitPtr)
            If Not IntPtr.Zero.Equals(shellExtInitPtr) Then
                Dim shellExtInit As IShellExtInit = Marshal.GetObjectForIUnknown(shellExtInitPtr)
                Dim dataObject As IDataObject
                Functions.SHCreateDataObject(folderpidl, lastpidls.Count, lastpidls, IntPtr.Zero, GetType(IDataObject).GUID, dataObject)
                shellExtInit.Initialize(folderpidl, dataObject, IntPtr.Zero)
            End If
        Else
            Me.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
            Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))

            shellView.GetItemObject(SVGIO.SVGIO_BACKGROUND, GetType(IContextMenu).GUID, ptr)
            contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))
        End If

        Dim ptr3 As IntPtr = Marshal.GetIUnknownForObject(contextMenu), ptr2, ptr1
        Marshal.QueryInterface(ptr3, GetType(IContextMenu2).GUID, ptr2)
        If Not ptr2 = IntPtr.Zero Then
            contextMenu2 = Marshal.GetObjectForIUnknown(ptr2)
        End If
        Marshal.QueryInterface(ptr3, GetType(IContextMenu3).GUID, ptr1)
        If Not ptr1 = IntPtr.Zero Then
            contextMenu3 = Marshal.GetObjectForIUnknown(ptr1)
        End If

        hMenu = Functions.CreatePopupMenu()
        Dim flags As Integer = CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS
        If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
        contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, flags)
        If _firstContextMenuCall Then
            ' somehow very first call doesn't return all items
            Functions.DestroyMenu(hMenu)
            hMenu = Functions.CreatePopupMenu()
            Dim count = contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, flags)
            Debug.WriteLine("count=" & count)
            _firstContextMenuCall = False
        End If

        Return contextMenu
    End Function

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        Select Case e.Event
            Case SHCNE.CREATE
                If Not e.Item1 Is Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    Dim parentShellItem2 As IShellItem2
                    e.Item1.GetParent(parentShellItem2)
                    Dim parentFullPath As String
                    parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                    If Me.FullPath.Equals(parentFullPath) Then
                        SyncLock _itemsLock
                            If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                _items.Add(New Item(e.Item1, Me, _setIsLoadingAction))
                                Dim view As ICollectionView = CollectionViewSource.GetDefaultView(_items)
                                view.Refresh()
                            End If
                        End SyncLock
                    End If
                End If
            Case SHCNE.MKDIR
                If Not e.Item1 Is Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) Then
                    Dim parentShellItem2 As IShellItem2
                    e.Item1.GetParent(parentShellItem2)
                    Dim parentFullPath As String
                    parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                    If Me.FullPath.Equals(parentFullPath) Then
                        SyncLock _itemsLock
                            If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                                _items.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
                            End If
                        End SyncLock
                    End If
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
                            _items.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
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
                SyncLock _itemsLock
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
                End SyncLock
        End Select
    End Sub

    Protected Function isWindows7OrLower() As Boolean
        Dim osVersion As Version = Environment.OSVersion.Version
        ' Windows 7 has version number 6.1
        Return osVersion.Major < 6 OrElse (osVersion.Major = 6 AndAlso osVersion.Minor <= 1)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            Marshal.ReleaseComObject(Me.ShellFolder)

            If Not _items Is Nothing Then
                For Each item In _items
                    item.Dispose()
                Next
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub
End Class
