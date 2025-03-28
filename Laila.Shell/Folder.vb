﻿Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Functions
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Public Class Folder
    Inherits Item
    Implements INotify

    Public Event ExpandAllGroups As EventHandler
    Public Event CollapseAllGroups As EventHandler

    Public Property LastScrollOffset As Point
    Public Property LastScrollSize As Size

    Protected _columns As List(Of Column)
    Friend _items As ItemsCollection(Of Item) = New ItemsCollection(Of Item)()
    Private _isExpanded As Boolean
    Private _isLoading As Boolean
    Private _isRefreshingItems As Boolean
    Friend _enumerationLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Protected _isLoaded As Boolean
    Private _enumerationException As Exception
    Friend _isEnumerated As Boolean
    Friend _isEnumeratedForTree As Boolean
    Private _isActiveInFolderView As Boolean
    Private _isInHistory As Boolean
    Private _view As String
    Protected _hasSubFolders As Boolean?
    Private _itemsSortPropertyName As String
    Private _itemsSortDirection As ListSortDirection
    Private _itemsGroupByPropertyName As String
    Protected _enumerationCancellationTokenSource As CancellationTokenSource
    Private _shellFolder As IShellFolder
    Private _isEmpty As Boolean
    Private _isListening As Boolean
    Private _listeningLock As Object = New Object()
    Private _wasActivity As Boolean
    Protected _initializeItemsGroupByPropertyName As String
    Protected _initializeItemsSortDirection As ListSortDirection = -1
    Protected _initializeItemsSortPropertyName As String
    Private _isInitializing As Boolean
    Protected _hookFolderFullPath As String
    Private _notificationSubscribersLock As Object = New Object()
    Private _notificationSubscribers As List(Of IProcessNotifications) = New List(Of IProcessNotifications)()
    Private _notificationThreadPool As Helpers.ThreadPool

    Public Shared Function FromDesktop() As Folder
        Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId()
        Return Shell.GlobalThreadPool.Run(
            Function() As Folder
                Dim pidl As IntPtr, shellFolder As IShellFolder = Nothing, shellItem2 As IShellItem2 = Nothing
                Try
                    Functions.SHGetDesktopFolder(shellFolder)
                    Functions.SHGetIDListFromObject(shellFolder, pidl)
                    Functions.SHCreateItemFromIDList(pidl, GetType(IShellItem2).GUID, shellItem2)
                Finally
                    If Not IntPtr.Zero.Equals(pidl) Then
                        Marshal.FreeCoTaskMem(pidl)
                        pidl = IntPtr.Zero
                    End If
                End Try

                Return New Folder(shellFolder, shellItem2, Nothing, threadId)
            End Function,, threadId)
    End Function

    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim result As IShellFolder = Nothing
        shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, result)
        Return result
    End Function

    ''' <summary>
    ''' This one is used for creating the root Desktop folder only.
    ''' </summary>
    Friend Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, parent As Folder, threadId As Integer?)
        Me.New(shellItem2, parent, True, True, threadId)

        _shellFolder = shellFolder

        Shell.StartListening(Me)
        _isListening = True
    End Sub

    Public Sub New(shellItem2 As IShellItem2, parent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, threadId As Integer?, Optional pidl As Pidl = Nothing)
        MyBase.New(shellItem2, parent, doKeepAlive, doHookUpdates, threadId, pidl)
        _canShowInTree = True
    End Sub

    Public Overrides Property IsProcessingNotifications As Boolean
        Get
            Return _isProcessingNotifications AndAlso (Me.IsVisibleInAddressBar OrElse Me.IsVisibleInTree OrElse If(Me._logicalParent?.IsActiveInFolderView, False) OrElse Me.IsActiveInFolderView)
        End Get
        Friend Set(value As Boolean)
            MyBase.IsProcessingNotifications = value
        End Set
    End Property

    Friend Overridable Function GetShellFolderOnCurrentThread() As IShellFolderForIContextMenu
        Return Folder.GetIShellFolderFromIShellItem2(Me.ShellItem2)
    End Function

    Public ReadOnly Property ShellFolder As IShellFolder
        Get
            If _shellFolder Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                Shell.GlobalThreadPool.Run(
                    Sub()
                        SyncLock _shellItemLock
                            If _shellFolder Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                                _shellFolder = Folder.GetIShellFolderFromIShellItem2(Me.ShellItem2)
                            End If
                        End SyncLock
                    End Sub)
            End If

            Return _shellFolder
        End Get
    End Property

    Public Property IsEmpty As Boolean
        Get
            Return _isEmpty
        End Get
        Friend Set(value As Boolean)
            SetValue(_isEmpty, value)
        End Set
    End Property

    Protected Friend ReadOnly Property IsRootFolder As Boolean
        Get
            Return Shell.GetSpecialFolders().Values.Contains(Me) OrElse Me.TreeRootIndex <> -1
        End Get
    End Property

    Public Overrides Property IsExpanded As Boolean
        Get
            Return _isExpanded 'AndAlso (Me.TreeRootIndex <> -1 OrElse If(_logicalParent, _parent) Is Nothing OrElse If(_logicalParent, _parent).IsExpanded)
        End Get
        Set(value As Boolean)
            SetValue(_isExpanded, value)

            If value Then
                Dim __ = Me.GetItemsAsync(,, True)
            Else
                For Each item In _items.ToList().Where(Function(i) TypeOf i Is Folder).ToList()
                    item.IsExpanded = False
                Next
            End If
        End Set
    End Property

    Protected Friend Property IsActiveInFolderView As Boolean
        Get
            Return _isActiveInFolderView
        End Get
        Set(value As Boolean)
            SetValue(_isActiveInFolderView, value)

            If Not value AndAlso TypeOf Me Is SearchFolder AndAlso Me.IsLoading Then
                Me.CancelEnumeration()
            End If
        End Set
    End Property

    Protected Friend Property IsInHistory As Boolean
        Get
            Return _isInHistory
        End Get
        Set(value As Boolean)
            SetValue(_isInHistory, value)
        End Set
    End Property

    Protected Friend Overrides ReadOnly Property IsReadyForDispose As Boolean
        Get
            Return Not _doKeepAlive AndAlso Not Me.IsActiveInFolderView AndAlso Not Me.IsExpanded _
                AndAlso Not Me.IsRootFolder AndAlso Not Me.IsVisibleInTree AndAlso Not Me.IsVisibleInAddressBar _
                AndAlso (If(_logicalParent, _parent) Is Nothing _
                         OrElse (Not If(_logicalParent, _parent).IsActiveInFolderView _
                                 AndAlso Not If(_logicalParent, _parent).IsVisibleInAddressBar)) _
                AndAlso Not Me.IsInHistory AndAlso _items.Count = 0
        End Get
    End Property

    Protected Friend Overrides Sub MaybeDispose()
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

    Public Function GetInfoTipFolderSizeAsync(cancellationToken As CancellationToken) As String
        Dim folderList As List(Of String) = Shell.GlobalThreadPool.Run(
            Function() As List(Of String)
                Dim flags As UInt32 = SHCONTF.FOLDERS Or SHCONTF.ENABLE_ASYNC
                If Shell.Settings.DoShowHiddenFilesAndFolders Then flags = flags Or SHCONTF.INCLUDEHIDDEN
                If Shell.Settings.DoShowProtectedOperatingSystemFiles Then flags = flags Or SHCONTF.INCLUDESUPERHIDDEN
                Return quickEnum(flags, 11)
            End Function)

        Dim fileList As List(Of String) = Shell.GlobalThreadPool.Run(
            Function() As List(Of String)
                Dim flags As UInt32 = SHCONTF.NONFOLDERS Or SHCONTF.ENABLE_ASYNC
                If Shell.Settings.DoShowHiddenFilesAndFolders Then flags = flags Or SHCONTF.INCLUDEHIDDEN
                If Shell.Settings.DoShowProtectedOperatingSystemFiles Then flags = flags Or SHCONTF.INCLUDESUPERHIDDEN
                Return quickEnum(flags, 11)
            End Function)

        Dim result As List(Of String) = New List(Of String)()

        If Me.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) Then
            Dim size As UInt64? = Shell.GlobalThreadPool.Run(
                Function() As UInt64?
                    Return getSizeRecursive(2500, cancellationToken)
                End Function)

            If size.HasValue AndAlso Not size.Value = 0 Then
                Dim prop As [Property] = [Property].FromCanonicalName("System.Size")
                prop._rawValue = New PROPVARIANT()
                prop._rawValue.SetValue(size.Value)
                result.Add(prop.DisplayNameWithColon & " " & prop.Text)
                prop.Dispose()
            End If
        End If

        Dim toString As Func(Of List(Of String), String) =
            Function(list As List(Of String)) As String
                Dim isMore As Boolean = list.Count = 11
                Dim text As String = String.Join(", ", list.Take(10))
                If text.Length > 75 Then
                    isMore = True
                    text = text.Substring(0, 75)
                    If text.LastIndexOf(",") >= 0 Then
                        text = text.Substring(0, text.LastIndexOf(","))
                    End If
                End If
                If isMore Then
                    If list(0).Length > 75 Then
                        text &= "..."
                    Else
                        text &= ", ..."
                    End If
                End If
                Return text
            End Function

        If Not folderList Is Nothing Then
            result.Add("Folders: " & toString(folderList))
        End If

        If Not fileList Is Nothing Then
            result.Add("Files: " & toString(fileList))
        End If

        If folderList Is Nothing AndAlso fileList Is Nothing Then
            result.Add("Empty folder")
        End If

        Return String.Join(Environment.NewLine, result)
    End Function

    Protected Function getSizeRecursive(timeout As Integer, cancellationToken As CancellationToken,
        Optional startTime As DateTime? = Nothing, Optional shellItem2 As IShellItem2 = Nothing) As UInt64?

        If Not startTime.HasValue Then startTime = DateTime.Now

        Dim flags As UInt32 = SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS Or SHCONTF.ENABLE_ASYNC
        If Shell.Settings.DoShowHiddenFilesAndFolders Then flags = flags Or SHCONTF.INCLUDEHIDDEN
        If Shell.Settings.DoShowProtectedOperatingSystemFiles Then flags = flags Or SHCONTF.INCLUDESUPERHIDDEN

        Dim result As UInt64? = 0
        Dim subFolders As List(Of IShellItem2) = New List(Of IShellItem2)()

        Dim bindCtx As ComTypes.IBindCtx = Nothing, propertyBag As IPropertyBag = Nothing
        Dim var As PROPVARIANT, enumShellItems As IEnumShellItems = Nothing
        Try
            Functions.CreateBindCtx(0, bindCtx)
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBag)
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 
            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 
            If shellItem2 Is Nothing Then
                SyncLock _shellItemLock
                    If Not Me.ShellItem2 Is Nothing Then
                        CType(Me.ShellItem2, IShellItem2ForIEnumShellItems).BindToHandler _
                            (bindCtx, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, enumShellItems)
                    End If
                End SyncLock
            Else
                CType(shellItem2, IShellItem2ForIEnumShellItems).BindToHandler _
                    (bindCtx, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, enumShellItems)
            End If

            If Not enumShellItems Is Nothing Then
                Dim celt As Integer = 5000
                Dim shellItems(celt - 1) As IShellItem, fetched As UInt32 = 1
                Dim h As HRESULT = enumShellItems.Next(celt, shellItems, fetched)
                While fetched > 0
                    For x As UInt32 = 0 To fetched - 1
                        If Not cancellationToken.IsCancellationRequested Then
                            Dim attr As SFGAO = SFGAO.FOLDER
                            shellItems(x).GetAttributes(attr, attr)
                            If attr.HasFlag(SFGAO.FOLDER) Then
                                subFolders.Add(shellItems(x))
                            Else
                                Dim prop As [Property] = [Property].FromCanonicalName("System.Size", CType(shellItems(x), IShellItem2))
                                If Not prop?.Value Is Nothing AndAlso TypeOf prop.Value Is UInt64 Then
                                    result += prop.Value
                                End If
                                prop?.Dispose()
                                If Not shellItems(x) Is Nothing Then
                                    Marshal.ReleaseComObject(shellItems(x))
                                    shellItems(x) = Nothing
                                End If
                            End If
                        Else
                            If Not shellItems(x) Is Nothing Then
                                Marshal.ReleaseComObject(shellItems(x))
                                shellItems(x) = Nothing
                            End If
                        End If
                        Thread.Sleep(2)
                    Next
                    If DateTime.Now.Subtract(startTime).TotalMilliseconds <= timeout _
                                AndAlso Not cancellationToken.IsCancellationRequested Then
                        h = enumShellItems.Next(celt, shellItems, fetched)
                    Else
                        result = Nothing
                        Exit While
                    End If
                End While
            End If
        Finally
            If Not enumShellItems Is Nothing Then
                Marshal.ReleaseComObject(enumShellItems)
                enumShellItems = Nothing
            End If
            If Not propertyBag Is Nothing Then
                Marshal.ReleaseComObject(propertyBag)
                propertyBag = Nothing
            End If
            If Not bindCtx Is Nothing Then
                Marshal.ReleaseComObject(bindCtx)
                bindCtx = Nothing
            End If
            var.Dispose()
        End Try

        If subFolders.Count > 0 Then
            For Each subFolderShellItem2 In subFolders
                If Not cancellationToken.IsCancellationRequested Then
                    If Not result Is Nothing Then
                        Dim subFolderSize As UInt64? = getSizeRecursive(timeout, cancellationToken, startTime, subFolderShellItem2)
                        If Not subFolderSize.HasValue Then
                            result = Nothing
                        Else
                            result += subFolderSize
                        End If
                    End If
                End If
                If Not subFolderShellItem2 Is Nothing Then
                    Marshal.ReleaseComObject(subFolderShellItem2)
                    subFolderShellItem2 = Nothing
                End If
                Thread.Sleep(2)
            Next
        End If

        Return result
    End Function

    Protected Function quickEnum(flags As UInt32, celt As Integer) As List(Of String)
        Dim bindCtx As ComTypes.IBindCtx = Nothing, propertyBag As IPropertyBag = Nothing
        Dim var As PROPVARIANT, enumShellItems As IEnumShellItems = Nothing
        Try
            Functions.CreateBindCtx(0, bindCtx)
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBag)
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 
            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 
            SyncLock _shellItemLock
                If Not Me.ShellItem2 Is Nothing Then
                    CType(Me.ShellItem2, IShellItem2ForIEnumShellItems).BindToHandler _
                        (bindCtx, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, enumShellItems)
                End If
            End SyncLock

            If Not enumShellItems Is Nothing Then
                Dim shellItems(celt - 1) As IShellItem, fetched As UInt32 = 1
                Dim h As HRESULT = enumShellItems.Next(celt, shellItems, fetched)
                If fetched > 0 Then
                    Dim displayName As String = Nothing
                    Dim displayNameList As List(Of String) = New List(Of String)()
                    For x As UInt32 = 0 To fetched - 1
                        shellItems(x).GetDisplayName(SHGDN.NORMAL, displayName)
                        displayNameList.Add(displayName)
                        If Not shellItems(x) Is Nothing Then
                            Marshal.ReleaseComObject(shellItems(x))
                            shellItems(x) = Nothing
                        End If
                    Next
                    Return displayNameList
                End If
            End If
        Finally
            If Not enumShellItems Is Nothing Then
                Marshal.ReleaseComObject(enumShellItems)
                enumShellItems = Nothing
            End If
            If Not propertyBag Is Nothing Then
                Marshal.ReleaseComObject(propertyBag)
                propertyBag = Nothing
            End If
            If Not bindCtx Is Nothing Then
                Marshal.ReleaseComObject(bindCtx)
                bindCtx = Nothing
            End If
            var.Dispose()
        End Try
        Return Nothing
    End Function

    Public ReadOnly Property Columns(canonicalName As String) As Column
        Get
            Return Me.Columns.SingleOrDefault(Function(c) canonicalName.Equals(c.CanonicalName))
        End Get
    End Property

    Public ReadOnly Property ColumnManager As IColumnManager
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As IShellView
                    Dim shellView As IShellView = Nothing
                    Me.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, shellView)
                    Return shellView
                End Function)
        End Get
    End Property

    Public ReadOnly Property Columns As IEnumerable(Of Column)
        Get
            If _columns Is Nothing Then
                _columns = New List(Of Column)()

                ' get columns from shell
                Dim columnManager As IColumnManager = Nothing
                Try
                    columnManager = Me.ColumnManager
                    Dim count As Integer
                    columnManager.GetColumnCount(CM_ENUM_FLAGS.CM_ENUM_ALL, count)
                    Dim propertyKeys(count - 1) As PROPERTYKEY
                    columnManager.GetColumns(CM_ENUM_FLAGS.CM_ENUM_ALL, propertyKeys, count)
                    Dim index As Integer = 0
                    For Each propertyKey In propertyKeys
                        Dim info As CM_COLUMNINFO = New CM_COLUMNINFO()
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
                    If Not _columns.Any(Function(c) c.IsVisible) Then
                        For Each col In _columns.Take(5)
                            col.IsVisible = True
                        Next
                    End If
                Finally
                    If Not columnManager Is Nothing Then
                        Marshal.ReleaseComObject(columnManager)
                        columnManager = Nothing
                    End If
                End Try
            End If

            Return _columns
        End Get
    End Property

    Public Overrides Property HasSubFolders As Boolean
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As Boolean
                    If _hasSubFolders.HasValue Then
                        Return _hasSubFolders.Value
                    ElseIf Not disposedValue Then
                        Dim attr As SFGAO = SFGAO.HASSUBFOLDER
                        SyncLock _shellItemLock
                            If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                                Me.ShellItem2.GetAttributes(attr, attr)
                            Else
                                attr = 0
                            End If
                        End SyncLock
                        Return attr.HasFlag(SFGAO.HASSUBFOLDER)
                    Else
                        Return False
                    End If
                End Function)
        End Get
        Set(value As Boolean)
            SetValue(_hasSubFolders, value)
        End Set
    End Property

    Public Overridable Sub AddRightClickMenuItems(menu As RightClickMenu)
        Dim osver As Version = Environment.OSVersion.Version
        Dim isWindows11 As Boolean = osver.Major = 10 AndAlso osver.Minor = 0 AndAlso osver.Build >= 22000

        ' add menu items
        Dim menuItems As List(Of Control) = menu.GetMenuItems()
        Dim lastMenuItem As Control = Nothing
        For Each item In menuItems
            Dim verb As String = If(Not item.Tag Is Nothing, CType(item.Tag, Tuple(Of Integer, String)).Item2, Nothing)
            Select Case verb
                Case "copy", "cut", "paste", "delete", "pintohome", "rename"
                    ' don't add these
                Case Else
                    Dim isNotDoubleSeparator As Boolean = Not (TypeOf item Is Separator AndAlso
                                (Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator))
                    Dim isNotInitialSeparator As Boolean = Not (TypeOf item Is Separator AndAlso Me.Items.Count = 0)
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

        ' add buttons
        Dim hasPaste As Boolean =
                Not menu.IsDefaultOnly _
                AndAlso (menu.SelectedItems Is Nothing OrElse menu.SelectedItems.Count = 0) _
                AndAlso Clipboard.CanPaste(Me)

        Dim menuItem As MenuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "cut")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "copy")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "paste")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", ""))) _
                Else If hasPaste Then menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String)(-1, "paste"), "Paste"))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "rename")
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 AndAlso menu.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANRENAME)) Then _
                menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String)(-1, "rename"), "Rename"))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "delete")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 Then
            Dim isPinned As Boolean = PinnedItems.GetIsPinned(menu.SelectedItems(0))
            menu.Buttons.Add(menu.MakeToggleButton(New Tuple(Of Integer, String)(-1, "laila.shell.(un)pin"),
                                                        If(isPinned, "Unpin item", "Pin item"), isPinned))
        End If
    End Sub

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
            If Not _isInitializing AndAlso Not String.IsNullOrWhiteSpace(_initializeItemsGroupByPropertyName) Then
                _isInitializing = True
                Me.ItemsGroupByPropertyName = _initializeItemsGroupByPropertyName
                If Me.ItemsGroupByPropertyName = _initializeItemsGroupByPropertyName Then
                    _initializeItemsGroupByPropertyName = Nothing
                End If
                _isInitializing = False
            End If
            Return _items
        End Get
    End Property

    Public Async Function RefreshItemsAsync() As Task
        Using Shell.OverrideCursor(Cursors.AppStarting)
            Me.IsRefreshingItems = True
            If Me.IsLoading Then
                Me.CancelEnumeration()
            End If
            _isEnumerated = False
            _isEnumeratedForTree = False
            Await GetItemsAsync()
            Me.IsRefreshingItems = False
        End Using
    End Function

    Public Overridable Function GetItems(Optional doRefreshAllExistingItems As Boolean = True, Optional isForTree As Boolean = False) As List(Of Item)
        _enumerationLock.Wait()
        Try
            If (Not _isEnumerated AndAlso Not isForTree) OrElse (Not _isEnumeratedForTree AndAlso isForTree) Then
                Me.IsLoading = True
                Dim prevEnumerationCancellationTokenSource As CancellationTokenSource _
                    = _enumerationCancellationTokenSource

                Dim cts As CancellationTokenSource = New CancellationTokenSource()
                _enumerationCancellationTokenSource = cts

                enumerateItems(False, cts.Token, -1, doRefreshAllExistingItems)

                ' terminate previous enumeration thread
                If Not prevEnumerationCancellationTokenSource Is Nothing Then
                    prevEnumerationCancellationTokenSource.Cancel()
                End If
            End If
            Return _items.ToList()
        Finally
            If _enumerationLock.CurrentCount = 0 Then
                _enumerationLock.Release()
            End If
            Me.IsLoading = False
        End Try
    End Function

    Public Overridable Async Function GetItemsAsync(Optional doRefreshAllExistingItems As Boolean = True,
                                                    Optional doRecursive As Boolean = False,
                                                    Optional isForTree As Boolean = False) As Task(Of List(Of Item))
        Dim tcs As New TaskCompletionSource(Of List(Of Item))

        Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId()
        Shell.GlobalThreadPool.Add(
            Sub()
                _enumerationLock.Wait()
                Try
                    If (Not _isEnumerated AndAlso Not isForTree) OrElse (Not _isEnumeratedForTree AndAlso isForTree) Then
                        Me.IsLoading = True
                        _enumerationCancellationTokenSource = New CancellationTokenSource()
                        enumerateItems(True, _enumerationCancellationTokenSource.Token, threadId, doRefreshAllExistingItems, doRecursive)
                    End If
                    tcs.SetResult(_items.ToList())
                Catch ex As Exception
                    tcs.SetException(ex)
                Finally
                    If _enumerationLock.CurrentCount = 0 Then
                        _enumerationLock.Release()
                    End If
                    Me.IsLoading = False
                End Try
            End Sub, threadId)

        Await tcs.Task.WaitAsync(Shell.ShuttingDownToken)
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            Return tcs.Task.Result
        End If
        Return Nothing
    End Function

    Public Overridable Sub CancelEnumeration()
        If Not _enumerationCancellationTokenSource Is Nothing Then
            _enumerationCancellationTokenSource.Cancel()
            If _enumerationLock.CurrentCount = 0 Then
                _enumerationLock.Release()
            End If
        End If
    End Sub

    Protected Sub enumerateItems(isAsync As Boolean, cancellationToken As CancellationToken, threadId As Integer?,
                                 Optional doRefreshAllExistingItems As Boolean = True, Optional doRecursive As Boolean = False)
        Debug.WriteLine("Start loading " & Me.DisplayName & " (" & Me.FullPath & ")")
        Me.IsEmpty = False

        Dim flags As UInt32 = SHCONTF.FOLDERS Or SHCONTF.NONFOLDERS
        If Shell.Settings.DoShowHiddenFilesAndFolders Then flags = flags Or SHCONTF.INCLUDEHIDDEN
        If Shell.Settings.DoShowProtectedOperatingSystemFiles Then flags = flags Or SHCONTF.INCLUDESUPERHIDDEN
        If isAsync Then flags = flags Or SHCONTF.ENABLE_ASYNC

        enumerateItems(flags, cancellationToken, threadId, doRefreshAllExistingItems, doRecursive)

        Dim poolSize As Integer = Math.Min(100, Math.Max(1, _notificationSubscribers.Count / 1000))
        If _notificationThreadPool Is Nothing Then
            _notificationThreadPool = New Helpers.ThreadPool(poolSize)
        ElseIf poolSize > _notificationThreadPool.Size Then
            _notificationThreadPool.Redimension(poolSize)
        End If

        If Not cancellationToken.IsCancellationRequested Then
            _wasActivity = False
            Me.IsEmpty = _items.Count = 0
            Debug.WriteLine("End loading " & Me.DisplayName)
        Else
            Debug.WriteLine("Cancelled loading " & Me.DisplayName)
        End If
    End Sub

    Protected Sub enumerateItems(flags As UInt32, cancellationToken As CancellationToken, threadId As Integer?,
                                 doRefreshAllExistingItems As Boolean, doRecursive As Boolean)
        If disposedValue Then Return

        Dim result As Dictionary(Of String, Item) = New Dictionary(Of String, Item)
        Dim newFullPaths As HashSet(Of String) = New HashSet(Of String)()

        ' pre-parse sort property 
        Dim isSortPropertyByText As Boolean, sortPropertyKey As String = Nothing, isSortPropertyDisplaySortValue As Boolean
        If Not String.IsNullOrWhiteSpace(Me.ItemsSortPropertyName) Then
            isSortPropertyByText = Me.ItemsSortPropertyName.Contains("PropertiesByKeyAsText")
            If isSortPropertyByText Then
                sortPropertyKey = Me.ItemsSortPropertyName.Substring(Me.ItemsSortPropertyName.IndexOf("[") + 1)
                sortPropertyKey = sortPropertyKey.Substring(0, sortPropertyKey.IndexOf("]"))
            End If
            isSortPropertyDisplaySortValue = Me.ItemsSortPropertyName = "ItemNameDisplaySortValue"
        End If

        Dim addItems As System.Action =
            Sub()
                Dim existingItems() As Tuple(Of Item, Item) = Nothing

                UIHelper.OnUIThread(
                    Sub()
                        ' init collection
                        If Not String.IsNullOrWhiteSpace(_initializeItemsSortPropertyName) Then Me.ItemsSortPropertyName = _initializeItemsSortPropertyName
                        If Not _initializeItemsSortDirection = -1 Then Me.ItemsSortDirection = _initializeItemsSortDirection
                        If Not String.IsNullOrWhiteSpace(_initializeItemsGroupByPropertyName) Then Me.ItemsGroupByPropertyName = _initializeItemsGroupByPropertyName

                        If Not cancellationToken.IsCancellationRequested _
                                AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested Then
                            If Not TypeOf Me Is SearchFolder Then
                                If _items.Count = 0 Then
                                    ' this happens the first time a folder is loaded
                                    _items.UpdateRange(result.Values, Nothing)
                                    For Each item In result.Values
                                        item.IsProcessingNotifications = True
                                    Next
                                Else
                                    ' this happens when a folder is refreshed
                                    Dim hasDupes As HashSet(Of String) = New HashSet(Of String)
                                    Dim previousFullPaths As HashSet(Of String) = New HashSet(Of String)
                                    For Each item In _items
                                        If Not hasDupes.Contains(item.FullPath & item.DeDupeKey) Then
                                            Try
                                                previousFullPaths.Add(item.FullPath & item.DeDupeKey)
                                            Catch ex As Exception
                                                hasDupes.Add(item.FullPath & item.DeDupeKey)
                                                previousFullPaths.Remove(item.FullPath & item.DeDupeKey)
                                            End Try
                                        Else
                                            previousFullPaths.Add(item.Pidl.ToString() & item.DeDupeKey)
                                        End If
                                    Next
                                    If hasDupes.Count > 0 Then
                                        For Each item In _items.ToList().Where(Function(i) hasDupes.Contains(i.FullPath & i.DeDupeKey))
                                            previousFullPaths.Add(item.Pidl.ToString() & item.DeDupeKey)
                                        Next
                                    End If
                                    Dim newItems As Item() = result.Where(Function(i) Not previousFullPaths.Contains(i.Key)).Select(Function(kv) kv.Value).ToArray()
                                    Dim removedItems As Item() = _items.ToList().Where(Function(i) Not newFullPaths.Contains(If(Not hasDupes.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl.ToString() & i.DeDupeKey))).ToArray()
                                    existingItems = _items.ToList().Where(Function(i) newFullPaths.Contains(If(Not hasDupes.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl.ToString() & i.DeDupeKey))) _
                                            .Select(Function(i) New Tuple(Of Item, Item)(i, result(If(Not hasDupes.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl.ToString() & i.DeDupeKey)))).ToArray()

                                    Dim seq As EqualityComparer(Of String) = EqualityComparer(Of String).Default
                                    For Each item In existingItems
                                        If Not seq.Equals(item.Item1.TreeSortPrefix, item.Item2.TreeSortPrefix) Then
                                            item.Item1.TreeSortPrefix = item.Item2.TreeSortPrefix
                                        End If
                                        If Not seq.Equals(item.Item1.ItemNameDisplaySortValuePrefix, item.Item2.ItemNameDisplaySortValuePrefix) Then
                                            item.Item1.ItemNameDisplaySortValuePrefix = item.Item2.ItemNameDisplaySortValuePrefix
                                        End If
                                    Next

                                    ' add/remove items
                                    _items.UpdateRange(newItems, removedItems)
                                    For Each item In newItems
                                        item.IsProcessingNotifications = True
                                    Next
                                End If
                            Else
                                ' this is for search folders
                                Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                                For Each item In result.Values
                                    _items.InsertSorted(item, c)
                                Next
                                For Each item In result.Values
                                    item.IsProcessingNotifications = True
                                Next
                            End If
                        End If
                    End Sub)

                ' refresh existing items (this never happens for search folders, so it won't ever be
                ' cancelled, except when forcefully disposing items, so basically only on shutdown
                If doRefreshAllExistingItems _
                        AndAlso Not existingItems Is Nothing AndAlso existingItems.Count > 0 _
                        AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested _
                        AndAlso Not cancellationToken.IsCancellationRequested Then

                    Dim size As Integer = Math.Max(1, Math.Min(existingItems.Count / 10, 50))
                    Dim chuncks()() As Tuple(Of Item, Item) = existingItems.Chunk(existingItems.Count / size).ToArray()
                    Dim tcses As List(Of TaskCompletionSource) = New List(Of TaskCompletionSource)()

                    ' threads for refreshing
                    For i = 0 To chuncks.Count - 1
                        Dim j As Integer = i
                        Dim tcs As TaskCompletionSource = New TaskCompletionSource()
                        tcses.Add(tcs)
                        Shell.GlobalThreadPool.Add(
                            Sub()
                                'Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") started for " & Me.FullPath)
                                Dim seq As EqualityComparer(Of String) = EqualityComparer(Of String).Default

                                ' Process tasks from the queue
                                For Each item In chuncks(j)
                                    If cancellationToken.IsCancellationRequested _
                                            OrElse Shell.ShuttingDownToken.IsCancellationRequested Then
                                        Exit For
                                    End If

                                    Try
                                        If item.Item1.IsPinned <> item.Item2.IsPinned Then item.Item1.IsPinned = item.Item2.IsPinned
                                        If item.Item1.CanShowInTree <> item.Item2.CanShowInTree Then item.Item1.CanShowInTree = item.Item2.CanShowInTree
                                        For Each [property] In item.Item1._propertiesByKey.Where(Function(p) p.Value.IsCustom).ToList()
                                            item.Item1._propertiesByKey.Remove([property].Key)
                                        Next
                                        For Each [property] In item.Item2._propertiesByKey.Where(Function(p) p.Value.IsCustom).ToList()
                                            item.Item1._propertiesByKey.Add([property].Key, [property].Value)
                                        Next

                                        If Not item.Item2 Is Nothing Then
                                            item.Item1.Refresh(item.Item2.ShellItem2)
                                            SyncLock item.Item2._shellItemLock
                                                item.Item2._shellItem2 = Nothing
                                            End SyncLock
                                            item.Item2._parent = Nothing
                                        Else
                                            item.Item1.Refresh()
                                        End If

                                        '' preload sort property
                                        'If isSortPropertyByText Then
                                        '    Dim sortValue As Object = item.Item1.PropertiesByKeyAsText(sortPropertyKey)?.Value
                                        'ElseIf isSortPropertyDisplaySortValue Then
                                        '    Dim sortValue As Object = item.Item1.ItemNameDisplaySortValue
                                        'End If

                                        If doRecursive AndAlso TypeOf item.Item1 Is Folder Then
                                            Dim recursiveFolder As Folder = item.Item1
                                            If recursiveFolder._isLoaded Then
                                                recursiveFolder._isEnumerated = False
                                                recursiveFolder._isEnumeratedForTree = False
                                                Dim __ = recursiveFolder.GetItemsAsync(doRefreshAllExistingItems, doRecursive)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        Debug.WriteLine("Exception refreshing " & item.Item1.FullPath & ": " & ex.Message)
                                    End Try
                                Next
                                'Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") finished for " & Me.FullPath)

                                tcs.SetResult()
                            End Sub)
                    Next

                    'Task.WaitAll(tcses.Select(Function(tcs) tcs.Task).ToArray(), cancellationToken)
                End If
            End Sub

        SyncLock _listeningLock
            If Not _isListening AndAlso Not _shellItem2 Is Nothing Then
                Shell.StartListening(Me)
                _isListening = True
            End If
        End SyncLock

        Me.EnumerateItems(Me.ShellItem2, flags, cancellationToken,
            isSortPropertyByText, isSortPropertyDisplaySortValue, sortPropertyKey,
            result, newFullPaths, addItems, threadId)

        If Not cancellationToken.IsCancellationRequested Then
            ' add new items
            addItems()

            _isEnumerated = True
            _isEnumeratedForTree = True
            _isLoaded = True

            UIHelper.OnUIThread(
                Sub()
                    ' set and update HasSubFolders property
                    Me.HasSubFolders = Not Me.Items.ToList().FirstOrDefault(Function(i) i.CanShowInTree) Is Nothing
                End Sub)
        End If
    End Sub

    Protected Overridable Sub EnumerateItems(shellItem2 As IShellItem2, flags As UInt32, cancellationToken As CancellationToken,
        isSortPropertyByText As Boolean, isSortPropertyDisplaySortValue As Boolean, sortPropertyKey As String,
        result As Dictionary(Of String, Item), newFullPaths As HashSet(Of String), addItems As System.Action,
        threadId As Integer?)

        Dim isDebuggerAttached As Boolean = Debugger.IsAttached
        Dim isRootDesktop As Boolean = If(Me.Pidl?.Equals(Shell.Desktop.Pidl), False)
        Dim replacedWithCustomFolders As HashSet(Of String) = New HashSet(Of String)()
        For Each customFolder In Shell.CustomFolders
            replacedWithCustomFolders.Add(customFolder.ReplacesFullPath.ToLower())
        Next
        Dim hasDupes As List(Of Item) = New List(Of Item)

        Dim bindCtx As ComTypes.IBindCtx = Nothing, propertyBag As IPropertyBag = Nothing
        Dim var As PROPVARIANT, enumShellItems As IEnumShellItems = Nothing
        Try
            Functions.CreateBindCtx(0, bindCtx)
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBag)
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 
            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 
            SyncLock _shellItemLock
                CType(shellItem2, IShellItem2ForIEnumShellItems).BindToHandler _
                    (bindCtx, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, enumShellItems)
            End SyncLock

            If Not enumShellItems Is Nothing Then
                Dim celt As Integer = If(TypeOf Me Is SearchFolder, 1, 25000)
                Dim shellItems(celt - 1) As IShellItem, fetched As UInt32 = 1
                Dim lastRefresh As DateTime = DateTime.Now, lastUpdate As DateTime = DateTime.Now
                'Debug.WriteLine("{0:HH:mm:ss.ffff} Fetching first", DateTime.Now)
                If Not cancellationToken.IsCancellationRequested _
                    AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested Then
                    Dim h As HRESULT
                    Do
                        h = enumShellItems.Next(celt, shellItems, fetched)
                        'Debug.WriteLine("{0:HH:mm:ss.ffff} Fetched " & fetched & " items", DateTime.Now)
                        For x = 0 To fetched - 1
                            'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting attributes", DateTime.Now)
                            If Not cancellationToken.IsCancellationRequested _
                                AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested Then
                                Dim attr2 As SFGAO = SFGAO.FOLDER Or SFGAO.LINK
                                shellItems(x).GetAttributes(attr2, attr2)
                                Dim newItem As Item
                                If attr2.HasFlag(SFGAO.FOLDER) Then
                                    newItem = New Folder(shellItems(x), Me, False, False, threadId)

                                    If replacedWithCustomFolders.Contains(newItem.FullPath.ToLower()) Then
                                        newItem.Dispose()
                                        newItem = Folder.FromParsingName(
                                            Shell.CustomFolders.FirstOrDefault(Function(f) _
                                                f.ReplacesFullPath.ToLower().Equals(newItem.FullPath.ToLower()))?.FullPath, Me, False, False)
                                    End If
                                ElseIf attr2.HasFlag(SFGAO.LINK) Then
                                    newItem = New Link(shellItems(x), Me, False, False, threadId)
                                Else
                                    newItem = New Item(shellItems(x), Me, False, False, threadId)
                                End If

                                newItem = Me.InitializeItem(newItem)

                                If Not newItem Is Nothing Then
                                    Dim isAlreadyAdded As Boolean = False
                                    If isDebuggerAttached Then
                                        ' only check if debugger is attached, to avoid exception,
                                        ' otherwise, rely on exception being thrown, for speed
                                        isAlreadyAdded = newFullPaths.Contains(newItem.FullPath & newItem.DeDupeKey)
                                    End If
                                    If Not isAlreadyAdded Then
                                        Try
                                            result.Add(newItem.FullPath & newItem.DeDupeKey, newItem)
                                        Catch ex As Exception
                                            isAlreadyAdded = True
                                        End Try
                                    End If
                                    If isAlreadyAdded Then
                                        hasDupes.Add(newItem)
                                    End If

                                    ' preload sort property 
                                    If isSortPropertyByText Then
                                        Dim sortValue As Object = newItem.PropertiesByKeyAsText(sortPropertyKey)
                                    ElseIf isSortPropertyDisplaySortValue Then
                                        Dim sortValue As Object = newItem.ItemNameDisplaySortValue
                                    End If

                                    If TypeOf Me Is SearchFolder Then
                                        ' preload Content view mode properties because searchfolder is slow
                                        Dim contentViewModeProperties() As [Property] = newItem.ContentViewModeProperties

                                        ' preload System_StorageProviderUIStatus images
                                        Dim System_StorageProviderUIStatus As System_StorageProviderUIStatusProperty _
                                        = newItem.PropertiesByKey(System_StorageProviderUIStatusProperty.Key)
                                        If Not System_StorageProviderUIStatus Is Nothing _
                                        AndAlso System_StorageProviderUIStatus.RawValue.vt <> 0 Then
                                            Dim imgrefs As String() = System_StorageProviderUIStatus.ImageReferences16
                                        End If

                                        ' preload attributes
                                        Dim attributes As SFGAO = newItem.Attributes
                                    End If

                                    newFullPaths.Add(newItem.FullPath & newItem.DeDupeKey)

                                    If isRootDesktop Then
                                        If Not Shell.GetSpecialFolder(SpecialFolders.Home) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Home).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_001" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Gallery) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Gallery).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_002" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Pictures) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Pictures).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_003" _
                                    Else If newItem.FullPath.ToUpper().Equals(Shell.Desktop.FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_004" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Documents) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Documents).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_005" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.OneDrive) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.OneDrive).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_006" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.OneDriveBusiness) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.OneDriveBusiness).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_007" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Downloads) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Downloads).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_008" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Music) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Music).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_009" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Videos) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Videos).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_010" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.UserProfile) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.UserProfile).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_011" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.ThisPc) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.ThisPc).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_012" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Libraries) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Libraries).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_013" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.Network) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.Network).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_014" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.ControlPanel) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.ControlPanel).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_015" _
                                    Else If Not Shell.GetSpecialFolder(SpecialFolders.RecycleBin) Is Nothing AndAlso newItem.FullPath.ToUpper().Equals(Shell.GetSpecialFolder(SpecialFolders.RecycleBin).FullPath.ToUpper()) Then newItem.TreeSortPrefix = "_016" _
                                    Else newItem.TreeSortPrefix = "_100"
                                    End If
                                End If

                                If (DateTime.Now.Subtract(lastUpdate).TotalMilliseconds >= 1000 _
                                        OrElse result.Count >= 10) AndAlso TypeOf Me Is SearchFolder Then
                                    addItems()
                                    result.Clear()
                                    lastUpdate = DateTime.Now
                                End If
                            Else
                                If Not shellItems(x) Is Nothing Then
                                    Marshal.ReleaseComObject(shellItems(x))
                                    shellItems(x) = Nothing
                                End If
                            End If
                        Next
                        'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting next", DateTime.Now)
                    Loop While celt = fetched AndAlso Not cancellationToken.IsCancellationRequested _
                        AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested
                    If Not (h = HRESULT.S_OK OrElse h = HRESULT.S_FALSE OrElse h = HRESULT.ERROR_INVALID_PARAMETER) Then
                        Throw Marshal.GetExceptionForHR(h)
                    End If

                    If hasDupes.Count > 0 Then
                        For Each item In result.Where(Function(kv) hasDupes.Exists(Function(d) d.FullPath = kv.Value.FullPath)).ToList()
                            If Not hasDupes.Exists(Function(d) d.Pidl.Equals(item.Value.Pidl)) Then
                                hasDupes.Add(item.Value)
                                result.Remove(item.Key)
                            Else
                                hasDupes.Remove(hasDupes.FirstOrDefault(Function(d) d.Pidl.Equals(item.Value.Pidl)))
                            End If
                        Next
                        For Each item In hasDupes
                            If Not result.ContainsKey(item.Pidl.ToString()) Then
                                result.Add(item.Pidl.ToString() & item.DeDupeKey, item)
                            End If
                        Next
                    End If
                End If
            End If
            Me.EnumerationException = Nothing
        Catch ex As COMException
            If Not (ex.HResult = HRESULT.S_OK OrElse ex.HResult = HRESULT.S_FALSE OrElse ex.HResult = HRESULT.ERROR_INVALID_PARAMETER) Then
                Me.EnumerationException = ex
            End If
        Catch ex As Exception
            Me.EnumerationException = ex
        Finally
            If Not enumShellItems Is Nothing Then
                Marshal.ReleaseComObject(enumShellItems)
                enumShellItems = Nothing
            End If
            If Not propertyBag Is Nothing Then
                Marshal.ReleaseComObject(propertyBag)
                propertyBag = Nothing
            End If
            If Not bindCtx Is Nothing Then
                Marshal.ReleaseComObject(bindCtx)
                bindCtx = Nothing
            End If
            var.Dispose()
        End Try
    End Sub

    Protected Friend Overridable Function InitializeItem(item As Item) As Item
        Return item
    End Function

    Public Overridable ReadOnly Property CanSort As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overridable ReadOnly Property CanGroupBy As Boolean
        Get
            Return True
        End Get
    End Property

    Public Property ItemsSortPropertyName As String
        Get
            Return If(_initializeItemsSortPropertyName, _itemsSortPropertyName)
        End Get
        Set(value As String)
            If System.Windows.Application.Current.Dispatcher.CheckAccess() Then
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

                    _initializeItemsSortPropertyName = Nothing
                End If
            Else
                _initializeItemsSortPropertyName = value
            End If
        End Set
    End Property

    Public Property ItemsSortDirection As ListSortDirection
        Get
            Return If(_initializeItemsSortDirection <> -1, _initializeItemsSortDirection, _itemsSortDirection)
        End Get
        Set(value As ListSortDirection)
            If System.Windows.Application.Current.Dispatcher.CheckAccess() Then
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

                    _initializeItemsSortDirection = -1
                End If
            Else
                _initializeItemsSortDirection = value
            End If
        End Set
    End Property

    Public Property ItemsGroupByPropertyName As String
        Get
            Return If(_initializeItemsGroupByPropertyName, _itemsGroupByPropertyName)
        End Get
        Set(value As String)
            If System.Windows.Application.Current.Dispatcher.CheckAccess() Then
                Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
                If Not view Is Nothing Then
                    If Not String.IsNullOrWhiteSpace(value) Then
                        Dim groupDescription As PropertyGroupDescription = New PropertyGroupDescription(value)
                        Dim groupSortDesc As SortDescription = New SortDescription() With {
                            .PropertyName = value.Replace(".GroupByText", ".Value"),
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

                    _initializeItemsGroupByPropertyName = Nothing
                End If
            Else
                _initializeItemsGroupByPropertyName = value
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

    Public Sub SubscribeToNotifications(item As IProcessNotifications) Implements INotify.SubscribeToNotifications
        SyncLock _notificationSubscribersLock
            _notificationSubscribers.Add(item)
        End SyncLock
    End Sub

    Public Sub UnsubscribeFromNotifications(item As IProcessNotifications) Implements INotify.UnsubscribeFromNotifications
        SyncLock _notificationSubscribersLock
            _notificationSubscribers.Remove(item)
        End SyncLock
    End Sub

    Protected Friend Overrides Sub ProcessNotification(e As NotificationEventArgs)
        MyBase.ProcessNotification(e)

        If Not disposedValue Then
            If Me.Pidl?.Equals(Shell.GetSpecialFolder(SpecialFolders.RecycleBin).Pidl) Then
                If Me.IsExpanded OrElse Me.IsActiveInFolderView OrElse Me.IsVisibleInAddressBar Then
                    Dim info As SHQUERYRBINFO = New SHQUERYRBINFO()
                    info.cbSize = Marshal.SizeOf(info)
                    Functions.SHQueryRecycleBin(Nothing, info)
                    If info.i64NumItems <> _items.Count Then
                        Debug.WriteLine("RECYCLE BIN=" & info.i64NumItems & " vs " & _items.Count)
                        _isEnumerated = False
                        _isEnumeratedForTree = False
                        Dim __ = Me.GetItemsAsync()
                    End If
                End If
            End If

            Select Case e.Event
                Case SHCNE.CREATE, SHCNE.MKDIR
                    If _isLoaded Then
                        If e.Item1.Parent?.Pidl?.Equals(Me.Pidl) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(Me.FullPath) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(_hookFolderFullPath) Then
                            _wasActivity = True
                            Dim existing As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                        AndAlso (If(i.Pidl?.Equals(e.Item1.Pidl), False) OrElse If(i.FullPath?.Equals(e.Item1.FullPath), False)))
                            If existing Is Nothing Then
                                Me.InitializeItem(e.Item1)
                                e.Item1.LogicalParent = Me
                                e.Item1.IsProcessingNotifications = True
                                e.IsHandled1 = True
                                Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                                UIHelper.OnUIThread(
                                    Sub()
                                        _items.InsertSorted(e.Item1, c)
                                    End Sub)
                                Me.IsEmpty = _items.Count = 0
                                Shell.GlobalThreadPool.Run(
                                    Sub()
                                        e.Item1.Refresh()
                                    End Sub)
                            ElseIf TypeOf existing Is Folder Then
                                Shell.GlobalThreadPool.Run(
                                    Sub()
                                        Dim existingFolder As Folder = existing
                                        existingFolder.Refresh(e.Item1?.ShellItem2, e.Item1?.Pidl?.Clone(), e.Item1?.FullPath)
                                        e.Item1._shellItem2 = Nothing
                                        existingFolder._isEnumerated = False
                                        existingFolder._isEnumeratedForTree = False
                                        Dim __ = existingFolder.GetItemsAsync(True, True)
                                    End Sub)
                            End If
                        End If
                    End If
                Case SHCNE.RMDIR, SHCNE.DELETE
                    If _isLoaded Then
                        If e.Item1.Parent?.Pidl?.Equals(Me.Pidl) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(Me.FullPath) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(_hookFolderFullPath) Then
                            _wasActivity = True
                            Dim existing As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                        AndAlso (If(i.Pidl?.Equals(e.Item1.Pidl), False) OrElse If(i.FullPath?.Equals(e.Item1.FullPath), False)))
                            If Not existing Is Nothing Then
                                existing.Dispose()
                            End If
                        End If
                    End If
                Case SHCNE.DRIVEADD
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        _wasActivity = True
                        If Not _items Is Nothing AndAlso _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                                                                            AndAlso If(i.Pidl?.Equals(e.Item1.Pidl), False)) Is Nothing Then
                            Me.InitializeItem(e.Item1)
                            e.Item1.LogicalParent = Me
                            e.Item1.IsProcessingNotifications = True
                            e.IsHandled1 = True
                            Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                            UIHelper.OnUIThread(
                                Sub()
                                    _items.InsertSorted(e.Item1, c)
                                End Sub)
                            Me.IsEmpty = _items.Count = 0
                        End If
                        Shell.GlobalThreadPool.Run(
                            Sub()
                                e.Item1.Refresh()
                            End Sub)
                    End If
                Case SHCNE.DRIVEREMOVED
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        _wasActivity = True
                        Dim item As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                                                              AndAlso If(i.Pidl?.Equals(e.Item1.Pidl), False))
                        If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                            item.Dispose()
                        End If
                    End If
                Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                    If _isLoaded Then
                        If (Me.Pidl?.Equals(e.Item1?.Pidl) OrElse Me.FullPath?.Equals(e.Item1?.FullPath) OrElse _hookFolderFullPath?.Equals(e.Item1?.FullPath) _
                                OrElse (Shell.Desktop.Pidl.Equals(e.Item1.Pidl) AndAlso _wasActivity)) _
                                AndAlso (e.Event = SHCNE.UPDATEDIR OrElse _wasActivity) Then
                            If (Me.IsExpanded OrElse Me.IsActiveInFolderView OrElse Me.IsVisibleInAddressBar) _
                                AndAlso Not TypeOf Me Is SearchFolder Then
                                _isEnumerated = False
                                _isEnumeratedForTree = False
                                Dim __ = Me.GetItemsAsync()
                            End If
                        End If
                    End If
            End Select

            If _isLoaded Then
                ' notify children
                Dim list As List(Of IProcessNotifications) = Nothing
                SyncLock _notificationSubscribersLock
                    list = _notificationSubscribers.Where(Function(i) If(i?.IsProcessingNotifications, False)).ToList()
                End SyncLock
                If list.Count > 0 Then
                    Dim size As Integer = Math.Max(1, Math.Min(list.Count / 10, 250))
                    Dim chuncks()() As IProcessNotifications = list.Chunk(list.Count / size).ToArray()
                    Dim tcses As List(Of TaskCompletionSource) = New List(Of TaskCompletionSource)()

                    ' threads for refreshing
                    For i = 0 To chuncks.Count - 1
                        Dim j As Integer = i
                        Dim tcs As TaskCompletionSource = New TaskCompletionSource()
                        tcses.Add(tcs)
                        _notificationThreadPool.Add(
                            Sub()
                                ' Process tasks from the queue
                                For Each item In chuncks(j)
                                    If Shell.ShuttingDownToken.IsCancellationRequested Then Exit For
                                    item?.ProcessNotification(e)
                                Next

                                tcs.SetResult()
                            End Sub)
                    Next

                    Task.WaitAll(tcses.Select(Function(tcs) tcs.Task).ToArray(), Shell.ShuttingDownToken)
                End If
            End If
        End If
    End Sub

    Protected Overrides Sub Settings_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
        MyBase.Settings_PropertyChanged(s, e)

        Select Case e.PropertyName
            Case "DoShowProtectedOperatingSystemFiles", "DoShowHiddenFilesAndFolders"
                If _isLoaded AndAlso Not disposedValue Then
                    Me.CancelEnumeration()
                    _isEnumerated = False
                    _isEnumeratedForTree = False
                End If
                Dim __ = Me.GetItemsAsync()
            Case "DoShowDriveLetters"
                If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                    _isEnumerated = False
                    _isEnumeratedForTree = False
                End If
                Dim __ = Me.GetItemsAsync()
        End Select
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        Dim oldShellFolder As IShellFolder = Nothing
        Dim wasDisposed As Boolean = disposedValue

        SyncLock _shellItemLock
            If Not disposedValue Then
                MyBase.Dispose(disposing)

                SyncLock _listeningLock
                    If _isListening Then
                        Shell.StopListening(Me)
                    End If
                End SyncLock

                Me.CancelEnumeration()

                _notificationThreadPool?.Dispose()
                _notificationThreadPool = Nothing

                If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                    ' extensive cleanup, because we're still live:

                    ' switch shellfolder, we'll release it later
                    oldShellFolder = _shellFolder
                    _shellFolder = Nothing
                Else
                    ' quick cleanup, because we're shutting down anyway and it has to go fast:

                    ' release shellfolder
                    If Not _shellFolder Is Nothing Then
                        Marshal.ReleaseComObject(_shellFolder)
                        _shellFolder = Nothing
                    End If
                End If
            End If
        End SyncLock

        If Not Shell.ShuttingDownToken.IsCancellationRequested AndAlso Not wasDisposed Then
            ' dispose outside of the lock because it can take a while
            Shell.GlobalThreadPool.Add(
                Sub()
                    ' dispose of children
                    For Each item In _items.ToList()
                        item.Dispose()
                    Next

                    ' dispose of shellfolder
                    If Not oldShellFolder Is Nothing Then
                        Marshal.ReleaseComObject(oldShellFolder)
                        oldShellFolder = Nothing
                    End If
                End Sub, _livesOnThreadId)
        End If
    End Sub
End Class
