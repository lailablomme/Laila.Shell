Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Media.Imaging

Public Class Folder
    Inherits Item

    Private _interface As IShellFolder
    Private _columns As List(Of Column)
    Private _items As ObservableCollection(Of Item)
    Private _firstContextMenuCall As Boolean = True

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

    Public ReadOnly Property ItemsThreaded As ObservableCollection(Of Item)
        Get
            If _items Is Nothing Then
                If Not _setIsLoadingAction Is Nothing Then
                    _setIsLoadingAction(True)
                End If

                Dim t As Thread = New Thread(New ThreadStart(
                    Sub()
                        Dim result As ObservableCollection(Of Item) = Me.Items

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

    Public Property Items As ObservableCollection(Of Item)
        Get
            If _items Is Nothing Then
                Dim result As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()

                If Not isWindows7OrLower() Then
                    Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
                    Functions.CreateBindCtx(0, bindCtxPtr)
                    bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

                    Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
                    propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))

                    Dim var As New PROPVARIANT()
                    var.vt = VarEnum.VT_UI4
                    var.union.uintVal = CType(SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, UInt32) 'Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN Or SHCONTF.FOLDERS
                    propertyBag.Write("SHCONTF", var)

                    bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag)
                    Dim ptr2 As IntPtr
                    bindCtxPtr = Marshal.GetIUnknownForObject(bindCtx)

                    ShellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                    Dim enumShellItems As IEnumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))

                    Dim shellItemArray(0) As IShellItem, feteched As UInt32 = 1
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                            enumShellItems.Next(1, shellItemArray, feteched)
                        End Sub)
                    While feteched = 1
                        Dim attr As Integer = SFGAO.FOLDER
                        shellItemArray(0).GetAttributes(attr, attr)
                        If CBool(attr And SFGAO.FOLDER) Then
                            result.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(shellItemArray(0)), shellItemArray(0), Me, _setIsLoadingAction))
                        Else
                            result.Add(New Item(shellItemArray(0), Me, _setIsLoadingAction))
                        End If
                        'End If
                        Application.Current.Dispatcher.Invoke(
                            Sub()
                                enumShellItems.Next(1, shellItemArray, feteched)
                            End Sub)
                    End While
                Else
                    Dim list As IEnumIDList
                    Me.ShellFolder.EnumObjects(Nothing, SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, list)
                    If Not list Is Nothing Then
                        Dim pidl(0) As IntPtr, fetched As Integer
                        While list.Next(1, pidl, fetched) = 0
                            Dim attr As Integer = SFGAO.FOLDER
                            Me.ShellFolder.GetAttributesOf(1, pidl, attr)
                            Application.Current.Dispatcher.Invoke(
                                Sub()
                                    If CBool(attr And SFGAO.FOLDER) Then
                                        result.Add(New Folder(Me, pidl(0), Me, _setIsLoadingAction))
                                    Else
                                        result.Add(New Item(Me, pidl(0), Me, _setIsLoadingAction))
                                    End If
                                End Sub)
                        End While
                    End If
                End If

                Me.Items = result
            End If

            Return _items
        End Get
        Set(value As ObservableCollection(Of Item))
            SetValue(_items, value)
            Me.NotifyOfPropertyChange("ItemsThreaded")
        End Set
    End Property

    Public Function GetContextMenu(items As IEnumerable(Of Item), ByRef contextMenu As IContextMenu, ByRef defaultId As String) As ContextMenu
        Dim hMenu As IntPtr, contextMenu2 As IContextMenu2, contextMenu3 As IContextMenu3
        contextMenu = getIContextMenu(items, contextMenu2, contextMenu3, hMenu)
        Dim defaultIdLocal As String
        Dim getMenu As Func(Of IntPtr, List(Of Control)) =
            Function(hMenu2 As IntPtr) As List(Of Control)
                If Not contextMenu2 Is Nothing Then
                    contextMenu2.HandleMenuMsg(WM.INITMENUPOPUP, hMenu2, IntPtr.Zero)
                End If
                If Not contextMenu3 Is Nothing Then
                    Dim ptr3 As IntPtr
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
                        result.Add(New Separator())
                    Else
                        mii.fMask = MIIM.MIIM_ID
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim cmd As String
                        If mii.wID < 1000000 Then
                            Dim bytes(256) As Byte
                            contextMenu2.GetCommandString(mii.wID, GCS.VERBA, 0, bytes, 256)
                            cmd = Text.Encoding.ASCII.GetString(bytes).Trim(vbNullChar)
                        End If

                        mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = header.Replace("&", "_"),
                            .Icon = New Image() With {.Source = bitmapSource},
                            .Tag = mii.wID & vbTab & cmd,
                            .IsEnabled = Not CBool(mii.fState And MFS.MFS_DISABLED),
                            .FontWeight = If(CBool(mii.fState And MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
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

                Return result
            End Function

        Dim menu As ContextMenu = New ContextMenu()
        For Each item In getMenu(hMenu)
            menu.Items.Add(item)
        Next

        defaultId = defaultIdLocal

        Return menu
    End Function

    Public Sub InvokeCommand(contextMenu As IContextMenu, items As IEnumerable(Of Item), id As String)
        Dim cmi As New CMINVOKECOMMANDINFO
        cmi.lpVerb = Convert.ToUInt32(id.Split(vbTab)(0))
        cmi.lpDirectory = Me.FullPath
        cmi.fMask = CMIC.NOZONECHECKS Or CMIC.ASYNCOK
        cmi.nShow = SW.SHOWNORMAL
        cmi.hwnd = IntPtr.Zero
        cmi.cbSize = CUInt(Marshal.SizeOf(cmi))

        contextMenu.InvokeCommand(cmi)
    End Sub

    Private Function getIContextMenu(items As IEnumerable(Of Item),
                                     ByRef contextMenu2 As IContextMenu2, ByRef contextMenu3 As IContextMenu3,
                                     ByRef hMenu As IntPtr) As IContextMenu
        Dim contextMenu As IContextMenu

        If Not items Is Nothing AndAlso items.Count > 0 Then
            Dim pidls(items.Count - 1) As IntPtr
            Dim lastpidls(items.Count - 1) As IntPtr

            For i = 0 To items.Count - 1
                Functions.SHGetIDListFromObject(Marshal.GetIUnknownForObject(items(i).ShellItem2), pidls(i))
                lastpidls(i) = Functions.ILFindLastID(pidls(i))
            Next

            Dim ptr As IntPtr
            Me.ShellFolder.GetUIObjectOf(IntPtr.Zero, lastpidls.Length, lastpidls, GetType(IContextMenu).GUID, 0, ptr)
            contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))
        Else
            Dim ptr As IntPtr
            Me.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
            Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))

            shellView.GetItemObject(SVGIO.SVGIO_BACKGROUND, GetType(IContextMenu).GUID, ptr)
            contextMenu = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IContextMenu))
        End If

        Dim ptr3 As IntPtr = Marshal.GetIUnknownForObject(contextMenu), ptr2
        Marshal.QueryInterface(ptr3, GetType(IContextMenu2).GUID, ptr2)
        If Not ptr2 = IntPtr.Zero Then
            contextMenu2 = CType(Marshal.GetObjectForIUnknown(ptr2), IContextMenu2)
        End If
        Marshal.QueryInterface(ptr3, GetType(IContextMenu3).GUID, ptr2)
        If Not ptr2 = IntPtr.Zero Then
            contextMenu3 = CType(Marshal.GetObjectForIUnknown(ptr2), IContextMenu3)
        End If

        hMenu = Functions.CreatePopupMenu()
        contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS)
        If _firstContextMenuCall Then
            ' somehow very first call doesn't return all items
            Functions.DestroyMenu(hMenu)
            hMenu = Functions.CreatePopupMenu()
            Dim count = contextMenu.QueryContextMenu(hMenu, 0, 0, Integer.MaxValue, CMF.CMF_NORMAL Or CMF.CMF_EXTENDEDVERBS)
            Debug.WriteLine("count=" & count)
            _firstContextMenuCall = False
        End If

        Return contextMenu
    End Function

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        Select Case e.Event
            Case SHCNE.CREATE
                If Not e.Item1 Is Nothing Then
                    Dim parentShellItem2 As IShellItem2
                    e.Item1.GetParent(parentShellItem2)
                    Dim parentFullPath As String
                    parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                    If Me.FullPath.Equals(parentFullPath) Then
                        If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                            _items.Add(New Item(e.Item1, Me, _setIsLoadingAction))
                            Dim view As ICollectionView = CollectionViewSource.GetDefaultView(_items)
                            view.Refresh()
                        End If
                    End If
                End If
            Case SHCNE.MKDIR
                If Not e.Item1 Is Nothing Then
                    Dim parentShellItem2 As IShellItem2
                    e.Item1.GetParent(parentShellItem2)
                    Dim parentFullPath As String
                    parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                    If Me.FullPath.Equals(parentFullPath) Then
                        If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                            _items.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
                        End If
                    End If
                End If
            Case SHCNE.RMDIR, SHCNE.DELETE
                If Not _items Is Nothing Then
                    Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                    If Not item Is Nothing Then
                        _items.Remove(item)
                    End If
                End If
            Case SHCNE.DRIVEADD
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                    If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path) Is Nothing Then
                        _items.Add(New Folder(Folder.GetIShellFolderFromIShellItem2(e.Item1), e.Item1, Me, _setIsLoadingAction))
                    End If
                End If
            Case SHCNE.DRIVEREMOVED
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                    Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path)
                    If Not item Is Nothing Then
                        _items.Remove(item)
                    End If
                End If
            Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                If Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path) Then
                    _items = Nothing
                    Me.NotifyOfPropertyChange("ItemsThreaded")
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
