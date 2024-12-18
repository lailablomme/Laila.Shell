Imports System.Collections.Concurrent
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Forms
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.Helpers

Public Class Folder
    Inherits Item

    Public Event LoadingStateChanged(isLoading As Boolean)
    Public Event ExpandAllGroups As EventHandler
    Public Event CollapseAllGroups As EventHandler

    Public Property LastScrollOffset As Point
    Public Property LastScrollSize As Size

    Private _columns As List(Of Column)
    Friend _items As CustomObservableCollection(Of Item) = New CustomObservableCollection(Of Item)()
    Private _isExpanded As Boolean
    Private _isLoading As Boolean
    Private _isRefreshingItems As Boolean
    Friend _lock As Object = New Object()
    Private _isLoaded As Boolean
    Private _enumerationException As Exception
    Friend _isEnumerated As Boolean
    Private _isActiveInFolderView As Boolean
    Private _isVisibleInAddressBar As Boolean
    Private _isInHistory As Boolean
    Private _view As String
    Private _hasSubFolders As Boolean?
    Private _pendingUpdateCounter As Integer
    Private _itemsSortPropertyName As String
    Private _itemsSortDirection As ListSortDirection
    Private _itemsGroupByPropertyName As String
    Protected _enumerationCancellationTokenSource As CancellationTokenSource
    Private _threadsCancellationTokenSource As CancellationTokenSource
    Private _doSkipUPDATEDIR As DateTime?
    Private _shellFolder As IShellFolder
    Private _isEmpty As Boolean
    Private _threads As List(Of Thread) = New List(Of Thread)()
    Private _threadsLock As Object = New Object()
    Private _isListening As Boolean
    Private _listeningLock As Object = New Object()

    Public Shared Function FromDesktop() As Folder
        Dim ptr As IntPtr, pidl As IntPtr, shellFolder As IShellFolder, shellItem2 As IShellItem2
        Try
            Functions.SHGetDesktopFolder(shellFolder)
            ptr = Marshal.GetIUnknownForObject(shellFolder)
            Functions.SHGetIDListFromObject(ptr, pidl)
            Marshal.Release(ptr)
            Functions.SHCreateItemFromIDList(pidl, GetType(IShellItem2).GUID, ptr)
            shellItem2 = Marshal.GetObjectForIUnknown(ptr)
        Finally
            If Not IntPtr.Zero.Equals(ptr) Then
                Marshal.Release(ptr)
            End If
            If Not IntPtr.Zero.Equals(pidl) Then
                Marshal.FreeCoTaskMem(pidl)
            End If
        End Try

        Return New Folder(shellFolder, shellItem2, Nothing)
    End Function

    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim result As IShellFolder
        Dim ptr2 As IntPtr

        Try
            shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, ptr2)
            If Not IntPtr.Zero.Equals(ptr2) Then
                result = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IShellFolder))
            End If
        Finally
            If Not IntPtr.Zero.Equals(ptr2) Then
                Marshal.Release(ptr2)
            End If
        End Try

        Return result
    End Function

    'Friend Shared Function GetIShellFolderFromPidl(pidl As IntPtr, bindingParent As Folder) As IShellFolder
    '    Dim ptr As IntPtr
    '    Try
    '        bindingParent._shellFolder.BindToObject(pidl, IntPtr.Zero, Guids.IID_IShellFolder, ptr)
    '        Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellFolder))
    '    Finally
    '        Marshal.Release(ptr)
    '    End Try
    'End Function

    ''' <summary>
    ''' This one is used for creating the root Desktop folder only.
    ''' </summary>
    Friend Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, parent As Folder)
        MyBase.New(shellItem2, parent, True, True)

        _shellFolder = shellFolder

        Shell.StartListening(Me)
        _isListening = True
    End Sub

    Public Sub New(shellItem2 As IShellItem2, parent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean)
        MyBase.New(shellItem2, parent, doKeepAlive, doHookUpdates)
    End Sub

    Public ReadOnly Property ShellFolder As IShellFolder
        Get
            If Not disposedValue AndAlso _shellFolder Is Nothing Then
                _shellFolder = Folder.GetIShellFolderFromIShellItem2(Me.ShellItem2)
            End If
            Return _shellFolder
        End Get
    End Property

    Public Property IsEmpty As Boolean
        Get
            Return _isEmpty
        End Get
        Protected Set(value As Boolean)
            SetValue(_isEmpty, value)
        End Set
    End Property

    Public ReadOnly Property IsRootFolder As Boolean
        Get
            Return Shell.GetSpecialFolders().Values.Contains(Me) OrElse Me.TreeRootIndex <> -1
        End Get
    End Property

    Public Overrides Property IsExpanded As Boolean
        Get
            Return _isExpanded AndAlso (Me.TreeRootIndex <> -1 OrElse Me._parent Is Nothing OrElse Me._parent.IsExpanded)
        End Get
        Set(value As Boolean)
            SetValue(_isExpanded, value)

            If value Then
                Me.GetItemsAsync()
            Else
                For Each item In _items.Where(Function(i) TypeOf i Is Folder).ToList()
                    item.IsExpanded = False
                Next
            End If
        End Set
    End Property

    Public Property IsActiveInFolderView As Boolean
        Get
            Return _isActiveInFolderView
        End Get
        Set(value As Boolean)
            SetValue(_isActiveInFolderView, value)
        End Set
    End Property

    Public Property IsVisibleInAddressBar As Boolean
        Get
            Return _isVisibleInAddressBar
        End Get
        Set(value As Boolean)
            SetValue(_isVisibleInAddressBar, value)
        End Set
    End Property

    Public Property IsInHistory As Boolean
        Get
            Return _isInHistory
        End Get
        Set(value As Boolean)
            SetValue(_isInHistory, value)
        End Set
    End Property

    Public Overrides ReadOnly Property IsReadyForDispose As Boolean
        Get
            Return Not _doKeepAlive AndAlso Not Me.IsActiveInFolderView AndAlso Not Me.IsExpanded _
                AndAlso Not Me.IsRootFolder AndAlso Not Me.IsVisibleInTree AndAlso Not Me.IsVisibleInAddressBar _
                AndAlso (Me._parent Is Nothing _
                         OrElse (Not Me._parent.IsActiveInFolderView _
                                 AndAlso Not _parent.IsVisibleInAddressBar)) _
                AndAlso Not Me.IsInHistory AndAlso _items.Count = 0
        End Get
    End Property

    Public Overrides Sub MaybeDispose()
        If Me.IsReadyForDispose Then
            Me.Dispose()
        End If
    End Sub

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

    Public ReadOnly Property ColumnManager As IColumnManager
        Get
            Dim ptr As IntPtr, shellFolder As IShellFolder
            Try
                shellFolder = Me.ShellFolder
                shellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
                Dim shellView As IShellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
                Return shellView
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
            End Try
        End Get
    End Property

    Public ReadOnly Property Columns As IEnumerable(Of Column)
        Get
            If _columns Is Nothing Then
                _columns = New List(Of Column)()

                ' get columns from shell
                Dim columnManager As IColumnManager = Me.ColumnManager
                Dim count As Integer
                columnManager.GetColumnCount(CM_ENUM_FLAGS.CM_ENUM_ALL, count)
                Dim propertyKeys(count - 1) As PROPERTYKEY
                columnManager.GetColumns(CM_ENUM_FLAGS.CM_ENUM_ALL, propertyKeys, count)
                Dim index As Integer = 0
                For Each propertyKey In propertyKeys
                    Dim info As CM_COLUMNINFO
                    info.dwMask = CM_MASK.CM_MASK_NAME Or CM_MASK.CM_MASK_DEFAULTWIDTH Or CM_MASK.CM_MASK_IDEALWIDTH _
                                      Or CM_MASK.CM_MASK_STATE Or CM_MASK.CM_MASK_WIDTH
                    info.cbSize = Marshal.SizeOf(Of CM_COLUMNINFO)
                    columnManager.GetColumnInfo(propertyKey, info)
                    Dim col As Column = New Column(propertyKey, info, index)
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
            If Not disposedValue Then
                Dim tcs As New TaskCompletionSource(Of Boolean)

                Shell.SlowTaskQueue.Add(
                    Sub()
                        Try
                            If _hasSubFolders.HasValue Then
                                tcs.SetResult(_hasSubFolders.Value)
                            ElseIf Not disposedValue Then
                                Dim attr As SFGAO = SFGAO.HASSUBFOLDER
                                Me.ShellItem2.GetAttributes(attr, attr)
                                tcs.SetResult(attr.HasFlag(SFGAO.HASSUBFOLDER))
                            Else
                                tcs.SetResult(False)
                            End If
                        Catch ex As Exception
                            tcs.SetException(ex)
                        End Try
                    End Sub)

                tcs.Task.Wait(Shell.ShuttingDownToken)
                If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                    Return tcs.Task.Result
                Else
                    Return False
                End If
            Else
                Return False
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
        _isEnumerated = False
        Me.GetItems()
        Me.IsRefreshingItems = False
    End Sub

    Public Async Function RefreshItemsAsync() As Task
        Me.IsRefreshingItems = True
        _isEnumerated = False
        Await GetItemsAsync()
        Me.IsRefreshingItems = False
    End Function

    Public Overridable Function GetItems() As List(Of Item)
        SyncLock _lock
            If Not _isEnumerated Then
                enumerateItems(_items, False)
            End If
            Return _items.ToList()
        End SyncLock
    End Function

    Public Overridable Async Function GetItemsAsync() As Task(Of List(Of Item))
        Dim tcs As New TaskCompletionSource(Of List(Of Item))

        Dim t As Thread = New Thread(
            Sub()
                Try
                    SyncLock _lock
                        If Not _isEnumerated Then
                            enumerateItems(_items, True)
                        End If
                        tcs.SetResult(_items.ToList())
                    End SyncLock
                Catch ex As Exception
                    tcs.SetException(ex)
                End Try
            End Sub)
        t.SetApartmentState(ApartmentState.MTA)
        t.Start()

        Await tcs.Task.WaitAsync(Shell.ShuttingDownToken)
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            Return tcs.Task.Result
        End If
        Return Nothing
    End Function

    Protected Sub enumerateItems(items As ObservableCollection(Of Item), Optional isAsync As Boolean = False)
        Debug.WriteLine("Start loading " & Me.DisplayName & " (" & Me.FullPath & ")")
        Me.IsLoading = True
        Me.IsEmpty = False

        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
        End If

        Dim cts As CancellationTokenSource = New CancellationTokenSource()
        _enumerationCancellationTokenSource = cts

        Dim flags As UInt32 = SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN Or SHCONTF.STORAGE
        If isAsync Then flags = flags Or SHCONTF.ENABLE_ASYNC

        enumerateItems(flags,
            Function(shellItem2 As IShellItem2)
                Return New Folder(shellItem2, Me, False, True)
            End Function,
            Function(shellItem2 As IShellItem2)
                Return New Item(shellItem2, Me, False, True)
            End Function, cts.Token)

        If _enumerationCancellationTokenSource.Equals(cts) Then
            Me.IsLoading = False
            Me.IsEmpty = _items.Count = 0
            Debug.WriteLine("End loading " & Me.DisplayName & If(cts.Token.IsCancellationRequested, " cancelled", ""))
        Else
            Debug.WriteLine("End loading " & Me.DisplayName & " (didn't mark completion)" & If(cts.Token.IsCancellationRequested, " cancelled", ""))
        End If
    End Sub

    Protected Sub enumerateItems(flags As UInt32,
                                 makeNewFolder As Func(Of IShellItem2, Item), makeNewItem As Func(Of IShellItem2, Item),
                                 cancellationToken As CancellationToken)
        If disposedValue Then Return

        Dim result As Dictionary(Of String, Item) = New Dictionary(Of String, Item)
        Dim newFullPaths As HashSet(Of String) = New HashSet(Of String)()
        Dim doRefreshAfter As Boolean

        If Me.FullPath = "::{645FF040-5081-101B-9F08-00AA002F954E}" Then _doSkipUPDATEDIR = DateTime.Now

        If TypeOf Me Is SearchFolder Then
            UIHelper.OnUIThread(
                Sub()
                    For Each item In Me.Items
                        item._parent = Nothing
                    Next
                    Me.Items.Clear()
                End Sub)
        End If

        Dim addItems As System.Action =
            Sub()
                If result.Count > 0 Then
                    Dim existingItems() As Tuple(Of Item, Item)

                    UIHelper.OnUIThread(
                        Sub()
                            If Not TypeOf Me Is SearchFolder Then
                                Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)

                                ' save sorting/grouping
                                Dim sortPropertyName As String = Me.ItemsSortPropertyName
                                Dim sortDirection As ListSortDirection = Me.ItemsSortDirection
                                Dim groupByPropertyName As String = Me.ItemsGroupByPropertyName

                                ' disable sorting grouping
                                Me.IsNotifying = False
                                Using view.DeferRefresh()
                                    Me.ItemsSortPropertyName = Nothing
                                    Me.ItemsGroupByPropertyName = Nothing
                                End Using

                                If _items.Count = 0 Then
                                    ' add items
                                    _items.UpdateRange(result.Values, Nothing)
                                Else
                                    Dim previousFullPaths As HashSet(Of String) = New HashSet(Of String)()
                                    For Each item In _items
                                        previousFullPaths.Add(If(item.FullPath, item.DisplayName))
                                    Next
                                    Dim newItems As Item() = result.Values.Where(Function(i) Not previousFullPaths.Contains(If(i.FullPath, i.DisplayName))).ToArray()
                                    Dim removedItems As Item() = _items.Where(Function(i) Not newFullPaths.Contains(If(i.FullPath, i.DisplayName))).ToArray()
                                    existingItems = _items.Where(Function(i) newFullPaths.Contains(If(i.FullPath, i.DisplayName))) _
                                        .Select(Function(i) New Tuple(Of Item, Item)(i, result(If(i.FullPath, i.DisplayName)))).ToArray()

                                    ' add/remove items
                                    _items.UpdateRange(newItems, removedItems)
                                End If

                                ' restore sorting/grouping
                                Using view.DeferRefresh()
                                    Me.ItemsSortPropertyName = sortPropertyName
                                    Me.ItemsSortDirection = sortDirection
                                    Me.ItemsGroupByPropertyName = groupByPropertyName
                                End Using
                                Me.IsNotifying = True
                            Else
                                For Each item In result.Values
                                    Me.Items.Add(item)
                                Next

                                doRefreshAfter = True
                            End If
                        End Sub)

                    ' refresh existing items
                    If Not existingItems Is Nothing AndAlso existingItems.Count > 0 _
                        AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested _
                        AndAlso Not _enumerationCancellationTokenSource.IsCancellationRequested Then

                        Dim size As Integer = Math.Max(1, Math.Min(existingItems.Count / 10, 100))
                        Dim chuncks()() As Tuple(Of Item, Item) = existingItems.Chunk(existingItems.Count / size).ToArray()

                        SyncLock _threadsLock
                            If _threadsCancellationTokenSource Is Nothing Then
                                _threadsCancellationTokenSource = New CancellationTokenSource()
                                Shell.AddToCancellationTokenSources(_threadsCancellationTokenSource)
                            End If

                            ' threads for refreshing
                            For i = 0 To chuncks.Count - 1
                                Dim j As Integer = i
                                Dim mtaThread As Thread = New Thread(
                                    Sub()
                                        Try
                                            'Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") started for " & Me.FullPath)

                                            ' Process tasks from the queue
                                            For Each item In chuncks(j)
                                                If _threadsCancellationTokenSource.Token.IsCancellationRequested Then
                                                    Exit For
                                                End If

                                                Try
                                                    SyncLock item.Item2._refreshLock
                                                        item.Item1.Refresh(item.Item2.ShellItem2)
                                                        item.Item2._shellItem2 = Nothing
                                                    End SyncLock
                                                    item.Item2._parent = Nothing

                                                    ' preload sort property
                                                    If Not String.IsNullOrWhiteSpace(Me.ItemsSortPropertyName) Then
                                                        If Me.ItemsSortPropertyName.Contains("PropertiesByKeyAsText") Then
                                                            Dim pkey As String = Me.ItemsSortPropertyName.Substring(Me.ItemsSortPropertyName.IndexOf("[") + 1)
                                                            pkey = pkey.Substring(0, pkey.IndexOf("]"))
                                                            Dim sortValue As Object = item.Item1.PropertiesByKeyAsText(pkey).Value
                                                        ElseIf Me.ItemsSortPropertyName = "ItemNameDisplaySortValue" Then
                                                            Dim sortValue As Object = item.Item1.ItemNameDisplaySortValue
                                                        End If
                                                    End If
                                                Catch ex As Exception
                                                    Debug.WriteLine("Exception refreshing " & item.Item1.FullPath & ": " & ex.Message)
                                                End Try
                                            Next

                                            SyncLock _threadsLock
                                                _threads.Remove(mtaThread)
                                            End SyncLock
                                            'Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") finished for " & Me.FullPath)
                                        Catch ex As OperationCanceledException
                                            Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") was canceled for " & Me.FullPath)
                                        End Try
                                    End Sub)
                                mtaThread.SetApartmentState(ApartmentState.MTA)
                                mtaThread.Start()
                                _threads.Add(mtaThread)
                            Next
                        End SyncLock
                    End If
                End If
            End Sub

        SyncLock _listeningLock
            If Not _isListening AndAlso Not _shellItem2 Is Nothing Then
                Shell.StartListening(Me)
                _isListening = True
            End If
        End SyncLock

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
                        Dim shellItem2 As IShellItem2 = Me.ShellItem2
                        If Not shellItem2 Is Nothing Then
                            shellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                            If Not IntPtr.Zero.Equals(ptr2) Then
                                enumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))
                            End If
                        End If
                    Finally
                        If Not IntPtr.Zero.Equals(ptr2) Then Marshal.Release(ptr2)
                    End Try
                    If Not enumShellItems Is Nothing Then
                        Dim celt As Integer = If(TypeOf Me Is SearchFolder, 1, 25000)
                        Dim shellItems(celt - 1) As IShellItem, fetched As UInt32 = 1
                        Dim lastRefresh As DateTime = DateTime.Now, lastUpdate As DateTime = DateTime.Now
                        'Debug.WriteLine("{0:HH:mm:ss.ffff} Fetching first", DateTime.Now)
                        If Not cancellationToken.IsCancellationRequested Then
                            Dim h As HRESULT = enumShellItems.Next(celt, shellItems, fetched)
                            While fetched > 0
                                'Debug.WriteLine("{0:HH:mm:ss.ffff} Fetched " & fetched & " items", DateTime.Now)
                                For x = 0 To fetched - 1
                                    'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting attributes", DateTime.Now)
                                    Dim attr2 As Integer = SFGAO.FOLDER
                                    shellItems(x).GetAttributes(attr2, attr2)
                                    Dim newItem As Item
                                    If CBool(attr2 And SFGAO.FOLDER) Then
                                        newItem = makeNewFolder(shellItems(x))
                                    Else
                                        newItem = makeNewItem(shellItems(x))
                                    End If

                                    result.Add(If(newItem.FullPath, newItem.DisplayName), newItem)

                                    ' preload sort property
                                    If Not String.IsNullOrWhiteSpace(Me.ItemsSortPropertyName) Then
                                        If Me.ItemsSortPropertyName.Contains("PropertiesByKeyAsText") Then
                                            Dim pkey As String = Me.ItemsSortPropertyName.Substring(Me.ItemsSortPropertyName.IndexOf("[") + 1)
                                            pkey = pkey.Substring(0, pkey.IndexOf("]"))
                                            Dim sortValue As Object = newItem.PropertiesByKeyAsText(pkey).Value
                                        ElseIf Me.ItemsSortPropertyName = "ItemNameDisplaySortValue" Then
                                            Dim sortValue As Object = newItem.ItemNameDisplaySortValue
                                        End If
                                    End If

                                    ' preload System_StorageProviderUIStatus images
                                    'Dim System_StorageProviderUIStatus As System_StorageProviderUIStatusProperty _
                                    '    = newItem.PropertiesByKey(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey)
                                    'If Not System_StorageProviderUIStatus Is Nothing _
                                    '    AndAlso System_StorageProviderUIStatus.RawValue.vt <> 0 Then
                                    '    Dim imgrefs As String() = System_StorageProviderUIStatus.ImageReferences16
                                    'End If

                                    newFullPaths.Add(If(newItem.FullPath, newItem.DisplayName))

                                    If cancellationToken.IsCancellationRequested Then Exit While

                                    If DateTime.Now.Subtract(lastUpdate).TotalMilliseconds >= 2000 _
                                        AndAlso result.Count > 0 AndAlso TypeOf Me Is SearchFolder Then
                                        addItems()
                                        result.Clear()
                                        lastUpdate = DateTime.Now

                                        If DateTime.Now.Subtract(lastRefresh).TotalMilliseconds >= 25000 Then
                                            UIHelper.OnUIThread(
                                                Sub()
                                                    Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                                                    view.Refresh()
                                                End Sub)
                                            lastRefresh = DateTime.Now
                                        End If
                                    End If
                                    If cancellationToken.IsCancellationRequested Then Exit While
                                Next
                                'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting next", DateTime.Now)
                                h = enumShellItems.Next(celt, shellItems, fetched)
                            End While
                            If fetched = 0 AndAlso Not (h = HRESULT.S_OK OrElse h = HRESULT.S_FALSE OrElse h = HRESULT.ERROR_INVALID_PARAMETER) Then
                                Throw Marshal.GetExceptionForHR(h)
                            End If
                        End If
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

        If Not cancellationToken.IsCancellationRequested Then
            ' add new items
            addItems()

            If doRefreshAfter Then
                UIHelper.OnUIThread(
                    Sub()
                        Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                        view.Refresh()
                    End Sub)
            End If

            _isEnumerated = True
            _isLoaded = True

            UIHelper.OnUIThread(
                Sub()
                    ' set and update HasSubFolders property
                    _hasSubFolders = Not Me.Items.FirstOrDefault(Function(i) TypeOf i Is Folder) Is Nothing
                    Me.NotifyOfPropertyChange("HasSubFolders")
                End Sub)
        End If
    End Sub

    Public Property ItemsSortPropertyName As String
        Get
            Return _itemsSortPropertyName
        End Get
        Set(value As String)
            Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            If Not view Is Nothing Then
                Dim desc As SortDescription
                If Not String.IsNullOrWhiteSpace(value) Then
                    desc = New SortDescription() With {
                        .PropertyName = value,
                        .Direction = Me.ItemsSortDirection
                    }
                End If
                If view.SortDescriptions.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    view.SortDescriptions.Add(desc)
                ElseIf Not String.IsNullOrWhiteSpace(value) Then
                    view.SortDescriptions(view.SortDescriptions.Count - 1) = desc
                ElseIf Not String.IsNullOrWhiteSpace(Me.ItemsGroupByPropertyName) _
                    AndAlso String.IsNullOrWhiteSpace(value) _
                    AndAlso view.SortDescriptions.Count > 1 Then
                    For x = 2 To view.SortDescriptions.Count
                        view.SortDescriptions.RemoveAt(view.SortDescriptions.Count - 1)
                    Next
                ElseIf String.IsNullOrWhiteSpace(Me.ItemsGroupByPropertyName) _
                    AndAlso String.IsNullOrWhiteSpace(value) Then
                    view.SortDescriptions.Clear()
                End If

                Me.SetValue(_itemsSortPropertyName, value)
            End If
        End Set
    End Property

    Public Property ItemsSortDirection As ListSortDirection
        Get
            Return _itemsSortDirection
        End Get
        Set(value As ListSortDirection)
            Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            If Not view Is Nothing Then
                For x = 0 To view.SortDescriptions.Count - 1
                    Dim desc As SortDescription = New SortDescription() With {
                        .PropertyName = view.SortDescriptions(x).PropertyName,
                        .Direction = value
                    }
                    view.SortDescriptions(x) = desc
                Next

                Me.SetValue(_itemsSortDirection, value)
            End If
        End Set
    End Property

    Public Property ItemsGroupByPropertyName As String
        Get
            Return _itemsGroupByPropertyName
        End Get
        Set(value As String)
            Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            If Not view Is Nothing Then
                If Not String.IsNullOrWhiteSpace(value) Then
                    Dim groupDescription As PropertyGroupDescription = New PropertyGroupDescription(value)
                    Dim groupSortDesc As SortDescription = New SortDescription() With {
                        .PropertyName = value,
                        .Direction = Me.ItemsSortDirection
                    }
                    If view.GroupDescriptions.Count > 0 Then
                        view.GroupDescriptions(0) = groupDescription
                    Else
                        view.GroupDescriptions.Add(groupDescription)
                    End If
                    If view.SortDescriptions.Count = 1 Then
                        view.SortDescriptions.Insert(0, groupSortDesc)
                    ElseIf view.SortDescriptions.Count = 2 Then
                        view.SortDescriptions(0) = groupSortDesc
                    End If
                ElseIf view.GroupDescriptions.Count > 0 Then
                    view.GroupDescriptions.Clear()
                    If view.SortDescriptions.Count > 0 Then
                        view.SortDescriptions.RemoveAt(0)
                    End If
                End If

                Me.SetValue(_itemsGroupByPropertyName, value)
            End If
        End Set
    End Property

    Public Property View As String
        Get
            Return _view
        End Get
        Set(value As String)
            SetValue(_view, value)
        End Set
    End Property

    Public Sub TriggerExpandAllGroups()
        RaiseEvent ExpandAllGroups(Me, New EventArgs())
    End Sub

    Public Sub TriggerCollapseAllGroups()
        RaiseEvent CollapseAllGroups(Me, New EventArgs())
    End Sub

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        MyBase.shell_Notification(sender, e)

        If Not disposedValue Then
            Select Case e.Event
                Case SHCNE.CREATE
                    If _isLoaded Then
                        If Not e.Item1.Parent Is Nothing AndAlso e.Item1.Parent.FullPath?.Equals(Me.FullPath) Then
                            If Not _items Is Nothing Then
                                UIHelper.OnUIThread(
                                    Sub()
                                        Dim existing As Item = _items.FirstOrDefault(Function(i) Not i.disposedValue AndAlso i.FullPath?.Equals(e.Item1.FullPath))
                                        If existing Is Nothing Then
                                            _items.Add(Item.FromPidl(e.Item1.Pidl.AbsolutePIDL, Me, False))
                                        Else
                                            existing.Refresh()
                                        End If
                                    End Sub)
                            End If
                        End If
                    End If
                Case SHCNE.MKDIR
                    If _isLoaded Then
                        If Not e.Item1.Parent Is Nothing AndAlso e.Item1.Parent.FullPath?.Equals(Me.FullPath) Then
                            If Not _items Is Nothing Then
                                UIHelper.OnUIThread(
                                    Sub()
                                        Dim existing As Item = _items.FirstOrDefault(Function(i) Not i.disposedValue AndAlso i.FullPath?.Equals(e.Item1.FullPath))
                                        If existing Is Nothing Then
                                            _items.Add(Item.FromPidl(e.Item1.Pidl.AbsolutePIDL, Me, False))
                                        Else
                                            existing.Refresh()
                                        End If
                                    End Sub)
                            End If
                        End If
                    End If
                Case SHCNE.RMDIR, SHCNE.DELETE
                    If Not _items Is Nothing AndAlso _isLoaded Then
                        UIHelper.OnUIThread(
                            Sub()
                                Dim item2 As Item
                                item2 = _items.FirstOrDefault(Function(i) Not i.disposedValue AndAlso i.FullPath?.Equals(e.Item1.FullPath))
                                If Not item2 Is Nothing Then
                                    If TypeOf item2 Is Folder Then
                                        Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                            .Folder = item2,
                                            .[Event] = e.Event
                                        })
                                    End If
                                    item2.Dispose()
                                End If
                            End Sub)
                    End If
                Case SHCNE.DRIVEADD
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        UIHelper.OnUIThread(
                            Sub()
                                If Not _items Is Nothing AndAlso _items.FirstOrDefault(Function(i) Not i.disposedValue AndAlso i.FullPath?.Equals(e.Item1.FullPath)) Is Nothing Then
                                    Dim item1 As IShellItem2 = Item.GetIShellItem2FromPidl(e.Item1.Pidl.AbsolutePIDL, Nothing)
                                    If Not item1 Is Nothing Then
                                        _items.Add(New Folder(item1, Me, False, True))
                                    End If
                                End If
                            End Sub)
                    End If
                Case SHCNE.DRIVEREMOVED
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        UIHelper.OnUIThread(
                            Sub()
                                Dim item As Item
                                item = _items.FirstOrDefault(Function(i) Not i.disposedValue AndAlso i.FullPath?.Equals(e.Item1.FullPath))
                                If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                                    Shell.RaiseFolderNotificationEvent(Me, New Events.FolderNotificationEventArgs() With {
                                        .Folder = item,
                                        .[Event] = e.Event
                                    })
                                    item.Dispose()
                                End If
                            End Sub)
                    End If
                Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                    If (Me.FullPath?.Equals(e.Item1.FullPath) OrElse Shell.Desktop.FullPath.Equals(e.Item1.FullPath)) Then
                        If Not Me.Items Is Nothing AndAlso _isLoaded AndAlso _pendingUpdateCounter <= 2 _
                            AndAlso (Me.IsExpanded OrElse Me.IsActiveInFolderView OrElse Me.IsVisibleInAddressBar) _
                            AndAlso (Not _doSkipUPDATEDIR.HasValue _
                                     OrElse DateTime.Now.Subtract(_doSkipUPDATEDIR.Value).TotalMilliseconds > 1000) _
                            AndAlso Not TypeOf Me Is SearchFolder Then
                            _isEnumerated = False
                            Me.GetItemsAsync()
                        End If
                    End If
                    _doSkipUPDATEDIR = Nothing
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
            MyBase.Dispose(disposing)

            SyncLock _listeningLock
                If _isListening Then
                    Shell.StopListening(Me)
                End If
            End SyncLock

            If Not _enumerationCancellationTokenSource Is Nothing Then
                _enumerationCancellationTokenSource.Cancel()
            End If

            If Not _threadsCancellationTokenSource Is Nothing Then
                _threadsCancellationTokenSource.Cancel()
                Shell.RemoveFromCancellationTokenSources(_threadsCancellationTokenSource)
            End If

            For Each item In _items
                item.Dispose()
            Next

            If Not _shellFolder Is Nothing Then
                Marshal.ReleaseComObject(_shellFolder)
                _shellFolder = Nothing
            End If
        End If
    End Sub
End Class
