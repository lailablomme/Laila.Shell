Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Laila.Shell.Helpers

Public Class Folder
    Inherits Item

    Public Event LoadingStateChanged(isLoading As Boolean)

    Private _columns As List(Of Column)
    Friend _items As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()
    Friend _shellFolder As IShellFolder
    Private _columnManager As IColumnManager
    Private _isExpanded As Boolean
    Private _isLoading As Boolean
    Private _isRefreshingItems As Boolean
    Private _lock As Object = New Object()
    Private _isLoaded As Boolean
    Private _enumerationException As Exception
    Private _isEnumerated As Boolean

    Public Shared Function FromKnownFolderGuid(knownFolderGuid As Guid) As Folder
        Return FromParsingName("shell:::" & knownFolderGuid.ToString("B"), Nothing)
    End Function

    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim result As IShellFolder
        UIHelper.OnUIThread(
            Sub()
                Dim ptr2 As IntPtr
                shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, ptr2)
                result = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IShellFolder))
            End Sub)
        Return result
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

    Public Sub New(shellItem2 As IShellItem2, parent As Folder)
        MyBase.New(shellItem2, parent)

        If Not shellItem2 Is Nothing Then
            _shellFolder = Folder.GetIShellFolderFromIShellItem2(shellItem2)
        End If
    End Sub

    Public Overrides Property IsExpanded As Boolean
        Get
            Return _isExpanded AndAlso (Me.TreeRootIndex <> -1 OrElse Me.Parent.IsExpanded)
        End Get
        Set(value As Boolean)
            SetValue(_isExpanded, value)

            If value Then
                Dim t As Func(Of Task) =
                    Async Function() As Task
                        SyncLock _lock
                            If _items.Count = 0 Then
                                updateItems(_items, True)
                            End If
                        End SyncLock
                    End Function

                Task.Run(t)
            End If
        End Set
    End Property

    Public Overrides Property IsLoading As Boolean
        Get
            Return _isLoading
        End Get
        Set(value As Boolean)
            SetValue(_isLoading, value)
        End Set
    End Property

    Public Property IsRefreshingItems As Boolean
        Get
            Return _isRefreshingItems
        End Get
        Set(value As Boolean)
            SetValue(_isRefreshingItems, value)
        End Set
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

    Public Overrides ReadOnly Property HasSubFolders As Boolean
        Get
            If _isEnumerated Then
                Return Not _items.FirstOrDefault(Function(i) TypeOf i Is Folder) Is Nothing
            Else
                Return Me.Attributes.HasFlag(SFGAO.HASSUBFOLDER)
            End If
        End Get
    End Property

    Public Property EnumerationException As Exception
        Get
            Return _enumerationException
        End Get
        Set(value As Exception)
            SetValue(_enumerationException, value)
        End Set
    End Property

    Public Overridable ReadOnly Property Items As ObservableCollection(Of Item)
        Get
            Return _items
        End Get
    End Property

    Public Sub RefreshItems()
        Me.IsRefreshingItems = True
        For Each item In _items.ToList()
            _items.Remove(item)
        Next
        Me.GetItems()
        Me.IsRefreshingItems = False
    End Sub

    Public Async Function RefreshItemsAsync() As Task
        Me.IsRefreshingItems = True
        For Each item In _items.ToList()
            _items.Remove(item)
        Next
        Await Me.GetItemsAsync()
        Me.IsRefreshingItems = False
    End Function

    Public Overridable Function GetItems() As List(Of Item)
        If _items.Count = 0 Then
            updateItems(_items, False)
        End If

        Return _items.ToList()
    End Function

    Public Overridable Async Function GetItemsAsync() As Task(Of List(Of Item))
        Dim func As Func(Of Task(Of List(Of Item))) =
            Async Function() As Task(Of List(Of Item))
                SyncLock _lock
                    If _items.Count = 0 Then
                        updateItems(_items, False)
                    End If
                End SyncLock

                Return _items.ToList()
            End Function

        Return Await Task.Run(func)
    End Function

    Protected Sub updateItems(items As ObservableCollection(Of Item), isFromThread As Boolean)
        Me.IsLoading = True

        updateItems(SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN Or SHCONTF.STORAGE,
                    Function(fullPath As String) As Boolean
                        Return Not items.FirstOrDefault(Function(i) i.FullPath = fullPath AndAlso Not i.disposedValue) Is Nothing
                    End Function,
                    Sub(item As Item)
                        items.Add(item)
                    End Sub,
                    Sub(item As Item)
                        items.Remove(item)
                    End Sub,
                    Function() As List(Of String)
                        Return items.Select(Function(f) f.FullPath).ToList()
                    End Function,
                    Function(pathsBefore As List(Of String), pathsAfter As List(Of String)) As List(Of Item)
                        Return items.Where(Function(i) pathsBefore.Contains(i.FullPath) AndAlso Not pathsAfter.Contains(i.FullPath)).ToList()
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Folder(shellItem2, Me)
                    End Function,
                    Function(shellItem2 As IShellItem2)
                        Return New Item(shellItem2, Me)
                    End Function,
                    Sub(path As String)
                        Dim item As Item = items.FirstOrDefault(Function(i) i.FullPath = path AndAlso Not i.disposedValue)
                        If Not item Is Nothing Then
                            item.Refresh()
                        End If
                    End Sub)
    End Sub

    Protected Sub updateItems(flags As UInt32,
                              exists As Func(Of String, Boolean), add As Action(Of Item), remove As Action(Of Item),
                              getPathsBefore As Func(Of List(Of String)), getToBeRemoved As Func(Of List(Of String), List(Of String), List(Of Item)),
                              makeNewFolder As Func(Of IShellItem2, Item), makeNewItem As Func(Of IShellItem2, Item),
                              updateProperties As Action(Of String))
        If _shellItem2 Is Nothing Then Return

        Dim pathsAfter As List(Of String) = New List(Of String)
        Dim pathsBefore As List(Of String) = New List(Of String)
        Dim toAdd As List(Of Item) = New List(Of Item)()
        Dim toUpdate As List(Of String) = New List(Of String)()

        UIHelper.OnUIThread(
            Sub()
                pathsBefore = getPathsBefore()
            End Sub)

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
                        propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 

                        bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 

                        Dim ptr2 As IntPtr
                        Try
                            _shellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                            If Not IntPtr.Zero.Equals(ptr2) Then
                                enumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))
                            End If
                        Finally
                            If Not IntPtr.Zero.Equals(ptr2) Then Marshal.Release(ptr2)
                        End Try
                        If Not enumShellItems Is Nothing Then
                            Dim shellItems(0) As IShellItem, fetched As UInt32 = 1
                            enumShellItems.Next(1, shellItems, fetched)
                            While fetched = 1
                                Dim attr2 As Integer = SFGAO.FOLDER
                                shellItems(0).GetAttributes(attr2, attr2)
                                Dim fullPath As String = Item.GetFullPathFromShellItem2(shellItems(0))
                                pathsAfter.Add(fullPath)
                                If Not exists(fullPath) Then
                                    Dim newItem As Item
                                    If CBool(attr2 And SFGAO.FOLDER) Then
                                        newItem = makeNewFolder(shellItems(0))
                                    Else
                                        newItem = makeNewItem(shellItems(0))
                                    End If
                                    toAdd.Add(newItem)
                                Else
                                    toUpdate.Add(fullPath)
                                    Marshal.ReleaseComObject(shellItems(0))
                                End If
                                enumShellItems.Next(1, shellItems, fetched)
                            End While
                        End If
                        Me.EnumerationException = Nothing
                    Catch ex As Exception
                        Me.EnumerationException = ex
                    Finally
                        If Not enumShellItems Is Nothing Then
                            Marshal.ReleaseComObject(enumShellItems)
                        End If
                        Marshal.ReleaseComObject(bindCtx)
                        Marshal.ReleaseComObject(propertyBag)
                        var.Dispose()
                    End Try
                End If
            End If
        Else
            Try
                Dim list As IEnumIDList
                _shellFolder.EnumObjects(Nothing, flags, list)
                If Not list Is Nothing Then
                    Dim pidl(0) As IntPtr, fetched As Integer, count As Integer = 0
                    While list.Next(1, pidl, fetched) = 0
                        Dim shellItem2 As IShellItem2 = Item.GetIShellItem2FromPidl(pidl(0), _shellFolder)
                        Dim fullPath As String = Item.GetFullPathFromShellItem2(shellItem2)
                        pathsAfter.Add(fullPath)
                        If Not exists(fullPath) Then
                            Dim attr2 As Integer = SFGAO.FOLDER
                            _shellFolder.GetAttributesOf(1, pidl, attr2)
                            Dim newItem As Item
                            If CBool(attr2 And SFGAO.FOLDER) Then
                                newItem = makeNewFolder(shellItem2)
                            Else
                                newItem = makeNewItem(shellItem2)
                            End If
                            toAdd.Add(newItem)
                        Else
                            toUpdate.Add(fullPath)
                            Marshal.ReleaseComObject(shellItem2)
                        End If
                        If count Mod 100 = 0 Then Thread.Sleep(1)
                        count += 1
                    End While
                End If
                Me.EnumerationException = Nothing
            Catch ex As Exception
                Me.EnumerationException = ex
            End Try
        End If

        _isEnumerated = True
        _isLoaded = True

        For Each item In toUpdate
            updateProperties(item)
        Next

        UIHelper.OnUIThread(
            Sub()
                For Each item In toAdd
                    add(item)
                Next
                For Each item In getToBeRemoved(pathsBefore, pathsAfter)
                    If TypeOf item Is Folder Then
                        Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                           .Folder = item,
                           .[Event] = SHCNE.RMDIR
                        })
                    End If
                    remove(item)
                    item.Dispose()
                Next

                Me.NotifyOfPropertyChange("HasSubFolders")

                Me.IsLoading = False
            End Sub)
    End Sub


    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        If Not _shellItem2 Is Nothing AndAlso Not disposedValue Then
            Select Case e.Event
                Case SHCNE.CREATE
                    If Not String.IsNullOrWhiteSpace(e.Item1Path) AndAlso _isLoaded Then
                        Dim parentShellItem2 As IShellItem2
                        Try
                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                            If Not item1 Is Nothing Then
                                item1.GetParent(parentShellItem2)
                                Dim parentFullPath As String
                                parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                                If Me.FullPath.Equals(parentFullPath) Then
                                    If Not Me.IsLoading AndAlso Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                                        Dim attr As SFGAO = SFGAO.FOLDER
                                        item1.GetAttributes(attr, attr)
                                        If attr.HasFlag(SFGAO.FOLDER) Then
                                            _items.Add(New Folder(item1, Me))
                                        Else
                                            _items.Add(New Item(item1, Me))
                                        End If
                                    End If
                                End If
                            End If
                        Finally
                            If Not parentShellItem2 Is Nothing Then
                                Marshal.ReleaseComObject(parentShellItem2)
                            End If
                        End Try
                    End If
                Case SHCNE.MKDIR
                    If Not String.IsNullOrWhiteSpace(e.Item1Path) AndAlso _isLoaded Then
                        Dim parentShellItem2 As IShellItem2
                        Try
                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                            If Not item1 Is Nothing Then
                                item1.GetParent(parentShellItem2)
                                Dim parentFullPath As String
                                parentShellItem2.GetDisplayName(SHGDN.FORPARSING, parentFullPath)
                                If Me.FullPath.Equals(parentFullPath) Then
                                    If Not Me.IsLoading AndAlso Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                                        _items.Add(New Folder(item1, Me))
                                    End If
                                End If
                            End If
                        Finally
                            If Not parentShellItem2 Is Nothing Then
                                Marshal.ReleaseComObject(parentShellItem2)
                            End If
                        End Try
                    End If
                Case SHCNE.RMDIR, SHCNE.DELETE
                    If Not Me.IsLoading AndAlso Not _items Is Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) AndAlso _isLoaded Then
                        Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
                        If Not item Is Nothing Then
                            If TypeOf item Is Folder Then
                                Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                    .Folder = item,
                                    .[Event] = e.Event
                                })
                            End If
                            _items.Remove(item)
                        End If
                    End If
                Case SHCNE.DRIVEADD
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) AndAlso _isLoaded Then
                        If Not Me.IsLoading AndAlso Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue) Is Nothing Then
                            Dim item1 As IShellItem2 = Item.GetIShellItem2FromParsingName(e.Item1Path)
                            If Not item1 Is Nothing Then
                                _items.Add(New Folder(item1, Me))
                            End If
                        End If
                    End If
                Case SHCNE.DRIVEREMOVED
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso Not String.IsNullOrWhiteSpace(e.Item1Path) AndAlso _isLoaded Then
                        Dim item As Item = _items.FirstOrDefault(Function(i) i.FullPath = e.Item1Path AndAlso Not i.disposedValue)
                        If Not Me.IsLoading AndAlso Not item Is Nothing AndAlso TypeOf item Is Folder Then
                            Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                .Folder = item,
                                .[Event] = e.Event
                            })
                            _items.Remove(item)
                        End If
                    End If
                Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                    If (Me.FullPath.Equals(e.Item1Path) OrElse Shell.Desktop.FullPath.Equals(e.Item1Path)) AndAlso Not _items Is Nothing AndAlso _isLoaded Then
                        Dim func As Func(Of Task) =
                            Async Function() As Task
                                SyncLock _lock
                                    updateItems(_items, False)
                                End SyncLock
                            End Function

                        Task.Run(func)
                    End If
            End Select
        End If
    End Sub

    Protected Function isWindows7OrLower() As Boolean
        Dim osVersion As Version = Environment.OSVersion.Version
        ' Windows 7 has version number 6.1
        Return osVersion.Major < 6 OrElse (osVersion.Major = 6 AndAlso osVersion.Minor <= 1)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
        End If

        If Not _shellFolder Is Nothing Then
            Marshal.ReleaseComObject(_shellFolder)
            _shellFolder = Nothing
        End If
        If Not _columnManager Is Nothing Then
            Marshal.ReleaseComObject(_columnManager)
            _columnManager = Nothing
        End If

        MyBase.Dispose(disposing)
    End Sub
End Class
