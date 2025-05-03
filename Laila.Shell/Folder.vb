Imports System.Collections.ObjectModel
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

    Private _activeView As Guid?
    Protected _columns As List(Of Column)
    Private _defaultView As Guid
    Protected _enumerationCancellationTokenSource As CancellationTokenSource
    Private _enumerationException As Exception
    Friend _enumerationLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Protected _hasSubFolders As Boolean?
    Protected _hookFolderFullPath As String
    Protected _initializeItemsGroupByPropertyName As String
    Protected _initializeItemsSortDirection As ListSortDirection = -1
    Protected _initializeItemsSortPropertyName As String
    Private _isActiveInFolderView As Boolean
    Private _isEmpty As Boolean
    Friend _isEnumerated As Boolean
    Friend _isEnumeratedForTree As Boolean
    Private _isExpanded As Boolean
    Private _isInHistory As Boolean
    Private _isInitializing As Boolean
    Private _isListening As Boolean
    Protected _isLoaded As Boolean
    Private _isLoading As Boolean
    Private _isRefreshingItems As Boolean
    Friend _items As ItemsCollection(Of Item)
    Private _itemsGroupByPropertyName As String
    Private _itemsSortDirection As ListSortDirection
    Private _itemsSortPropertyName As String
    Private _listeningLock As Object = New Object()
    Private _notificationSubscribers As List(Of IProcessNotifications) = New List(Of IProcessNotifications)()
    Private _notificationSubscribersLock As Object = New Object()
    Private _notificationThreadPool As Helpers.ThreadPool
    Friend _previousFullPaths As HashSet(Of String) = New HashSet(Of String)()
    Friend _previousFullPathsLock As Object = New Object()
    Private _shellFolder As IShellFolder
    Protected _views As List(Of FolderViewRegistration)
    Friend _wasActivity As Boolean

    ''' <summary>
    ''' Creates a Folder object for the Desktop folder (the root).
    ''' </summary>
    ''' <returns>A Folder object representing the Desktop folder (the root)</returns>
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

    ''' <summary>
    ''' Gets an IShellFolder from an IShellItem2.
    ''' </summary>
    ''' <param name="shellItem2">The IShellItem2 object to derive the IShellFolder from</param>
    ''' <returns>An IShellFolder</returns>
    Friend Shared Function GetIShellFolderFromIShellItem2(shellItem2 As IShellItem2) As IShellFolder
        Dim result As IShellFolder = Nothing
        If Not shellItem2 Is Nothing Then
            shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, result)
        End If
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

    ''' <summary>
    ''' Main constructor for the Folder object.
    ''' </summary>
    ''' <param name="shellItem2">The IShellItem2 this Folder object is based on</param>
    ''' <param name="logicalParent">The logical parent for this Folder object</param>
    ''' <param name="doKeepAlive">Wether to keep this object alive (i.e. prevent it from getting disposed after minimum 10 seconds of inactivity).</param>
    ''' <param name="doHookUpdates">Wether or not this Folder should listen to shell notifications</param>
    ''' <param name="threadId">The GlobalThreadPool thread id of the thread this item lives on</param>
    ''' <param name="pidl">The PIDL for this Folder (optional)</param>
    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, threadId As Integer?, Optional pidl As Pidl = Nothing)
        MyBase.New(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, pidl)
        _canShowInTree = True

        Dim items As ItemsCollection(Of Item) = New ItemsCollection(Of Item)()
        BindingOperations.EnableCollectionSynchronization(items, items.Lock)
        _items = items

        _views = New List(Of FolderViewRegistration) From {
            New FolderViewRegistration(New Guid("5a0a0f55-1f79-4b96-8e47-0092f06d73d6"), "Extra large icons", New Uri("pack://application:,,,/Laila.Shell;component/Images/extralargeicons16.png"), GetType(ExtraLargeIconsView)),
            New FolderViewRegistration(New Guid("69785b58-7418-4f1b-ae2b-04f62a5cf090"), "Large icons", New Uri("pack://application:,,,/Laila.Shell;component/Images/largeicons16.png"), GetType(LargeIconsView)),
            New FolderViewRegistration(New Guid("bde20919-1132-4f64-8124-b7b1c5b3ec47"), "Normal icons", New Uri("pack://application:,,,/Laila.Shell;component/Images/normalicons16.png"), GetType(NormalIconsView)),
            New FolderViewRegistration(New Guid("0fc82e4f-844c-4ec4-90ff-fb4f2ec9b3c6"), "Small icons", New Uri("pack://application:,,,/Laila.Shell;component/Images/smallicons16.png"), GetType(SmallIconsView)),
            New FolderViewRegistration(New Guid("d2e1c65f-58b4-49d1-a4c9-5d43cb2554cf"), "List", New Uri("pack://application:,,,/Laila.Shell;component/Images/list16.png"), GetType(Controls.ListView)),
            New FolderViewRegistration(New Guid("ffc28f55-2f56-4698-97d4-f80ff8817713"), "Details", New Uri("pack://application:,,,/Laila.Shell;component/Images/details16.png"), GetType(DetailsView)),
            New FolderViewRegistration(New Guid("3e7ea73f-22d7-4697-b858-cde65c9c9ea7"), "Tiles", New Uri("pack://application:,,,/Laila.Shell;component/Images/tiles16.png"), GetType(TileView)),
            New FolderViewRegistration(New Guid("9d96a6be-4061-4c89-9cb7-d97c6b8dfe41"), "Content", New Uri("pack://application:,,,/Laila.Shell;component/Images/content16.png"), GetType(ContentView))
        }
        Me.DefaultView = New Guid("ffc28f55-2f56-4698-97d4-f80ff8817713")
    End Sub

    ''' <summary>
    ''' Gets or sets wether or not this Folder is processing shell notifications.
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Property IsProcessingNotifications As Boolean
        Get
            Return _isProcessingNotifications _
                AndAlso (If(Me._logicalParent?.IsVisibleInAddressBar, False) OrElse Me.IsVisibleInAddressBar _
                OrElse If(Me._logicalParent?.IsVisibleInTree, False) OrElse Me.IsVisibleInTree _
                OrElse If(Me._logicalParent?.IsActiveInFolderView, False) OrElse Me.IsActiveInFolderView)
        End Get
        Friend Set(value As Boolean)
            MyBase.IsProcessingNotifications = value
        End Set
    End Property

    ''' <summary>
    ''' Makes an IShellFolder for this Folder on the calling thread.
    ''' </summary>
    ''' <returns>A new IShellFolder, created on the calling thread</returns>
    Friend Overridable Function MakeIShellFolderOnCurrentThread() As IShellFolderForIContextMenu
        SyncLock _shellItemLockMakeIShellFolderOnCurrentThread
            If Not disposedValue Then
                Return Folder.GetIShellFolderFromIShellItem2(Me.ShellItem2)
            End If
        End SyncLock
    End Function

    ''' <summary>
    ''' Gets the IShellFolder for this Folder on the thread this Folder lives on.
    ''' </summary>
    ''' <returns>The main IShellFolder for this Folder</returns>
    Public ReadOnly Property ShellFolder As IShellFolder
        Get
            If _shellFolder Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                Shell.GlobalThreadPool.Run(
                    Sub()
                        SyncLock _shellItemLockShellFolder
                            If _shellFolder Is Nothing AndAlso Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                                _shellFolder = Folder.GetIShellFolderFromIShellItem2(Me.ShellItem2)
                            End If
                        End SyncLock
                    End Sub)
            End If

            Return _shellFolder
        End Get
    End Property

    ''' <summary>
    ''' Gets/sets wether or not this Folder turned out to be empty after enumerating it's children.
    ''' </summary>
    ''' <returns>True if this Folder has no children</returns>
    Public Property IsEmpty As Boolean
        Get
            Return _isEmpty
        End Get
        Friend Set(value As Boolean)
            SetValue(_isEmpty, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets wether or not this Folder is a root folder (i.e. is a special folder or lives in the root of the TreeView).
    ''' </summary>
    ''' <returns>True if this Folder is a root folder</returns>
    Protected Friend ReadOnly Property IsRootFolder As Boolean
        Get
            Return Shell.GetSpecialFolders().Values.Contains(Me) OrElse Me.TreeRootIndex <> -1
        End Get
    End Property

    ''' <summary>
    ''' Gets/sets wether or not this Folder is expanded in the TreeView. 
    ''' Setting this property to true enumerates it's children if necessary. 
    ''' Setting it to false will also collapse all it's children recursively.
    ''' </summary>
    ''' <returns>True if this Folder is expanded in the TreeView</returns>
    Public Overrides Property IsExpanded As Boolean
        Get
            Return _isExpanded
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

    ''' <summary>
    ''' Gets/sets wether or not this Folder is opened in a FolderView.
    ''' </summary>
    ''' <returns>True if this Folder is opened in a FolderView</returns>
    Protected Friend Property IsActiveInFolderView As Boolean
        Get
            Return _isActiveInFolderView
        End Get
        Friend Set(value As Boolean)
            SetValue(_isActiveInFolderView, value)

            If Not value AndAlso TypeOf Me Is SearchFolder AndAlso Me.IsLoading Then
                Me.CancelEnumeration()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets wether this Folder is listed in the history of a Navigation control.
    ''' </summary>
    ''' <returns>True if this folder is listed in the history of a Navigation control</returns>
    Protected Friend Property IsInHistory As Boolean
        Get
            Return _isInHistory
        End Get
        Friend Set(value As Boolean)
            SetValue(_isInHistory, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets wether this Folder can be safely disposed of.
    ''' </summary>
    ''' <returns>True when this folder can be safely disposed</returns>
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

    ''' <summary>
    ''' Dispose this Folder when it is ready to be safely disposed. Otherwise, do nothing.
    ''' </summary>
    Protected Friend Overrides Sub MaybeDispose()
        If Me.IsReadyForDispose Then
            Me.Dispose()
        End If
    End Sub

    ''' <summary>
    ''' Gets/sets wether or not this Folder's children are currently being enumerated.
    ''' </summary>
    ''' <returns>True when this Folder is in the process of enumerating it's children</returns>
    Public Overrides Property IsLoading As Boolean
        Get
            Return _isLoading
        End Get
        Set(value As Boolean)
            SetValue(_isLoading, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets wether or not this Folder is currently being refreshed (i.e. by a user who clicked the Refresh button).
    ''' </summary>
    ''' <returns>True when this Folder is currently being refreshed</returns>
    Public Property IsRefreshingItems As Boolean
        Get
            Return _isRefreshingItems
        End Get
        Friend Set(value As Boolean)
            SetValue(_isRefreshingItems, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets the text about this Folder's contents and size to show in the infotip.
    ''' </summary>
    ''' <param name="cancellationToken">Use this to cancel getting the information</param>
    ''' <returns>A string containing information about this Folder's contents and size, ready to be appended to the infotip for this Folder</returns>
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

    ''' <summary>
    ''' Recursively gets the size of this Folder, in bytes.
    ''' </summary>
    ''' <param name="timeout">Timeout after which to stop enumerating recursively and give up, returning null.</param>
    ''' <param name="cancellationToken">Use this token to cancel enumerating this Folder recursively to determine it's size on disk</param>
    ''' <param name="startTime">Contains the time we started enumerating this Folder, in case of recursive iterations</param>
    ''' <param name="shellItem2">Contains the child IShellItem2 for which this recursive iteration is getting the size</param>
    ''' <returns>An UInt64 containing the recursive size of this Folder on disk, in bytes, or null when cancelled or timed out.</returns>
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
                SyncLock _shellItemLockEnumShellItems
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

    ''' <summary>
    ''' Returns the names of a limited number of child items that are enumerated as quickly as possible. This is much quicker and simpler than
    ''' fully iterating the children for this Folder, but returns substantially less information too. This method is used to show an indication of 
    ''' the contents of this Folder in the infotip.
    ''' </summary>
    ''' <param name="flags">SHCONTF flags used in enumeration</param>
    ''' <param name="celt">Maximum number of items to enumerate.</param>
    ''' <returns>A string array containing the dispay names of the enumerated items</returns>
    Protected Function quickEnum(flags As SHCONTF, celt As Integer) As List(Of String)
        Dim bindCtx As ComTypes.IBindCtx = Nothing, propertyBag As IPropertyBag = Nothing
        Dim var As PROPVARIANT, enumShellItems As IEnumShellItems = Nothing
        Try
            Functions.CreateBindCtx(0, bindCtx)
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBag)
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 
            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 
            SyncLock _shellItemLockEnumShellItems
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

    ''' <summary>
    ''' Gets a column of this Folder by it's canonical name.
    ''' </summary>
    ''' <param name="canonicalName">The canonical name of the column to return</param>
    ''' <returns>The column that corresponds to the given canonical name</returns>
    Public ReadOnly Property Columns(canonicalName As String) As Column
        Get
            Return Me.Columns.SingleOrDefault(Function(c) canonicalName.Equals(c.CanonicalName))
        End Get
    End Property

    ''' <summary>
    ''' Returns the IColumnManager for this Folder.
    ''' </summary>
    ''' <returns>The IColumnManager for this Folder</returns>
    Public ReadOnly Property ColumnManager As IColumnManager
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As IShellView
                    SyncLock _shellItemLockColumnManager
                        Dim shellView As IShellView = Nothing
                        If Not disposedValue AndAlso Not Me.ShellFolder Is Nothing Then
                            Me.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, shellView)
                        End If
                        Return shellView
                    End SyncLock
                End Function)
        End Get
    End Property

    ''' <summary>
    ''' Returns the columns for this Folder.
    ''' </summary>
    ''' <returns>An IEnumerable of Column objects, representing the columns in this Folder</returns>
    Public ReadOnly Property Columns As IEnumerable(Of Column)
        Get
            If _columns Is Nothing Then
                ' get columns from shell
                Dim columnManager As IColumnManager = Nothing
                Try
                    columnManager = Me.ColumnManager
                    If Not Me.ColumnManager Is Nothing Then
                        _columns = New List(Of Column)()

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

    ''' <summary>
    ''' Gets/sets wether or not this Folder has subfolders. This is used to determine wether or not to show the expand/collapse chevron in the TreeView.
    ''' </summary>
    ''' <returns>True if we think this Folder has subfolders</returns>
    Public Overrides Property HasSubFolders As Boolean
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As Boolean
                    If _hasSubFolders.HasValue AndAlso (_isEnumerated OrElse _hasSubFolders.Value) Then
                        Return _hasSubFolders.Value
                    ElseIf Not disposedValue Then
                        Dim attr As SFGAO = SFGAO.HASSUBFOLDER
                        SyncLock _shellItemLockHasSubFolders
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

    Public Overridable Sub AddExplorerMenuItems(menu As ExplorerMenu)
        Dim menuItems As List(Of Control) = menu.GetMenuItems()
        Dim lastMenuItem As Control = Nothing
        For Each item In menuItems
            Dim verb As String = If(Not item.Tag Is Nothing, CType(item.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing)
            Select Case verb
                Case Else
                    Dim isNotDoubleSeparator As Boolean = Not (TypeOf item Is Separator AndAlso
                                (Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator))
                    Dim isNotInitialSeparator As Boolean = Not (TypeOf item Is Separator AndAlso menu.Items.Count = 0)
                    If isNotDoubleSeparator AndAlso isNotInitialSeparator Then
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

        If Clipboard.CanCut(menu.SelectedItems) Then menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "cut", Nothing), "Cut"))
        If Clipboard.CanCopy(menu.SelectedItems) Then menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "copy", Nothing), "Copy"))
        If hasPaste Then menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "paste", Nothing), "Paste"))
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 AndAlso menu.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANRENAME)) Then _
                menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "rename", Nothing), "Rename"))
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count > 0 AndAlso menu.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANDELETE)) Then _
                menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "delete", Nothing), "Delete"))
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 Then
            Dim isPinned As Boolean = PinnedItems.GetIsPinned(menu.SelectedItems(0))
            menu.Buttons.Add(menu.MakeToggleButton(New Tuple(Of Integer, String, Object)(-1, "laila.shell.(un)pin", Nothing),
                                                        If(isPinned, "Unpin item", "Pin item"), isPinned))
        End If
    End Sub

    ''' <summary>
    ''' Adds the right-click menu items for this Folder and the items in this folder.
    ''' </summary>
    ''' <param name="menu">The RightClickMenu asking us to add it's items</param>
    Public Overridable Sub AddRightClickMenuItems(menu As RightClickMenu)
        Dim osver As Version = Environment.OSVersion.Version
        Dim isWindows11 As Boolean = osver.Major = 10 AndAlso osver.Minor = 0 AndAlso osver.Build >= 22000

        ' add menu items
        Dim menuItems As List(Of Control) = menu.GetMenuItems()
        Dim lastMenuItem As Control = Nothing
        For Each item In menuItems
            Dim verb As String = If(Not item.Tag Is Nothing, CType(item.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing)
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

        ' add buttons
        Dim hasPaste As Boolean =
                Not menu.IsDefaultOnly _
                AndAlso (menu.SelectedItems Is Nothing OrElse menu.SelectedItems.Count = 0) _
                AndAlso Clipboard.CanPaste(Me)

        Dim menuItem As MenuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing) = "cut")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing) = "copy")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing) = "paste")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", ""))) _
                Else If hasPaste Then menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "paste", Nothing), "Paste"))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing) = "rename")
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 AndAlso menu.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANRENAME)) Then _
                menu.Buttons.Add(menu.MakeButton(New Tuple(Of Integer, String, Object)(-1, "rename", Nothing), "Rename"))
        menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String, Object)).Item2, Nothing) = "delete")
        If Not menuItem Is Nothing Then menu.Buttons.Add(menu.MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count = 1 Then
            Dim isPinned As Boolean = PinnedItems.GetIsPinned(menu.SelectedItems(0))
            menu.Buttons.Add(menu.MakeToggleButton(New Tuple(Of Integer, String, Object)(-1, "laila.shell.(un)pin", Nothing),
                                                        If(isPinned, "Unpin item", "Pin item"), isPinned))
        End If
    End Sub

    ''' <summary>
    ''' Gets/sets the exception that occured while enumeration this Folder's children, if any. Otherwise, null.
    ''' </summary>
    ''' <returns>The exception that occured while enumeration this Folder's children, if any. Otherwise, null</returns>
    Public Property EnumerationException As Exception
        Get
            Return _enumerationException
        End Get
        Set(value As Exception)
            SetValue(_enumerationException, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets the children of this Folder, if they were enumerated already.
    ''' </summary>
    ''' <returns>An ObservableCollection containing the enumerated children for this Folder</returns>
    Public Overridable ReadOnly Property Items As ItemsCollection(Of Item)
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

    ''' <summary>
    ''' Refresh the contents of this folder.
    ''' </summary>
    ''' <returns>An awaitable Task</returns>
    Public Async Function RefreshItemsAsync() As Task
        Using Shell.OverrideCursor(Cursors.AppStarting)
            Me.IsRefreshingItems = True
            Me.CancelEnumeration()
            _isEnumerated = False
            _isEnumeratedForTree = False
            Await GetItemsAsync()
            Dim view As CollectionView = CollectionViewSource.GetDefaultView(Me.Items)
            view.Refresh()
            Me.IsRefreshingItems = False
        End Using
    End Function

    Public Overridable Function GetItems(Optional doRefreshAllExistingItems As Boolean = True, Optional isForTree As Boolean = False) As List(Of Item)
        _enumerationLock.Wait()
        Dim doStartLoading As Boolean =
            (Not _isEnumerated AndAlso Not isForTree) OrElse (Not _isEnumeratedForTree AndAlso isForTree)
        Try
            If doStartLoading Then
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
                Me.IsLoading = False
                _enumerationLock.Release()
            End If
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
                Dim doStartLoading As Boolean =
                    (Not _isEnumerated AndAlso Not isForTree) OrElse (Not _isEnumeratedForTree AndAlso isForTree)
                Dim originalCancellationTokenSource As CancellationTokenSource = Nothing
                Try
                    If doStartLoading Then
                        Me.IsLoading = True
                        _enumerationCancellationTokenSource = New CancellationTokenSource()
                        originalCancellationTokenSource = _enumerationCancellationTokenSource
                        enumerateItems(True, _enumerationCancellationTokenSource.Token, threadId, doRefreshAllExistingItems, doRecursive)
                    End If
                    tcs.SetResult(_items.ToList())
                Catch ex As Exception
                    tcs.SetException(ex)
                Finally
                    If _enumerationLock.CurrentCount = 0 Then
                        If _enumerationCancellationTokenSource.Equals(originalCancellationTokenSource) Then
                            Me.IsLoading = False
                        End If
                        _enumerationLock.Release()
                    End If
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
        Shell.BlockDisposer(True)
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
        Shell.BlockDisposer(False)
    End Sub

    Protected Sub enumerateItems(flags As UInt32, cancellationToken As CancellationToken, threadId As Integer?,
                                 doRefreshAllExistingItems As Boolean, doRecursive As Boolean)
        If disposedValue Then Return

        Dim result As Dictionary(Of String, Item) = New Dictionary(Of String, Item)
        Dim newFullPaths As HashSet(Of String) = New HashSet(Of String)()
        Dim dupes As List(Of Item) = New List(Of Item)

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

                SyncLock _items.Lock
                    ' init collection
                    If Not String.IsNullOrWhiteSpace(_initializeItemsSortPropertyName) Then Me.ItemsSortPropertyName = _initializeItemsSortPropertyName
                    If Not _initializeItemsSortDirection = -1 Then Me.ItemsSortDirection = _initializeItemsSortDirection
                    If Not String.IsNullOrWhiteSpace(_initializeItemsGroupByPropertyName) Then Me.ItemsGroupByPropertyName = _initializeItemsGroupByPropertyName

                    If Not cancellationToken.IsCancellationRequested _
                            AndAlso Not Shell.ShuttingDownToken.IsCancellationRequested Then
                        SyncLock _previousFullPathsLock
                            If Not TypeOf Me Is SearchFolder Then
                                If _items.Count = 0 Then
                                    ' this happens the first time a folder is loaded
                                    _items.UpdateRange(result.Values, Nothing)
                                    For Each item In result.Values
                                        item.IsProcessingNotifications = True
                                    Next
                                Else
                                    ' this happens when a folder is refreshed
                                    Dim dupeFullPaths As HashSet(Of String) = New HashSet(Of String)
                                    For Each i In dupes
                                        dupeFullPaths.Add(i.FullPath)
                                    Next
                                    Dim newItems As Item() = result.Where(Function(i) Not _previousFullPaths.Contains(i.Key)).Select(Function(kv) kv.Value).ToArray()
                                    Dim removedItems As Item() = _items.ToList().Where(Function(i) Not newFullPaths.Contains(If(Not dupeFullPaths.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl?.ToString() & i.DeDupeKey))).ToArray()
                                    existingItems = _items.ToList().Where(Function(i) newFullPaths.Contains(If(Not dupeFullPaths.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl?.ToString() & i.DeDupeKey))) _
                                            .Select(Function(i) New Tuple(Of Item, Item)(i, result(If(Not dupeFullPaths.Contains(i.FullPath), i.FullPath & i.DeDupeKey, i.Pidl?.ToString() & i.DeDupeKey)))).ToArray()

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

                            _previousFullPaths = newFullPaths
                        End SyncLock
                    End If
                End SyncLock

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

                        Dim t As Thread = New Thread(New ParameterizedThreadStart(
                            Sub(tcsObj As Object)
                                Dim tcs2 As TaskCompletionSource = tcsObj

                                Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") started for " & Me.FullPath)
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
                                        If item.Item2._hasCustomProperties Then
                                            For Each [property] In item.Item2._propertiesByKey.Where(Function(p) p.Value.IsCustom).ToList()
                                                item.Item1._propertiesByKey.Remove([property].Key)
                                                item.Item1._propertiesByKey.Add([property].Key, [property].Value)
                                            Next
                                        End If

                                        Dim newShellItem As IShellItem2 = Nothing
                                        If Not item.Item2 Is Nothing Then
                                            SyncLock item.Item2._shellItemLockEnumRefresh
                                                If Not item.Item2.disposedValue Then
                                                    newShellItem = item.Item2.ShellItem2
                                                    ' we've used this shell item in item1 now, so avoid it getting disposed when item2 gets disposed
                                                    'Debug.WriteLine(item.Item1.FullPath)
                                                    item.Item2._shellItem2 = Nothing
                                                    item.Item2.LogicalParent = Nothing
                                                    item.Item1.Refresh(newShellItem,,, item.Item2._livesOnThreadId)
                                                End If
                                            End SyncLock
                                        End If

                                        ' preload sort property
                                        If isSortPropertyByText Then
                                            Dim sortValue As Object = item.Item1.PropertiesByKeyAsText(sortPropertyKey)?.Value
                                        ElseIf isSortPropertyDisplaySortValue Then
                                            Dim sortValue As Object = item.Item1.ItemNameDisplaySortValue
                                        End If

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
                                Debug.WriteLine("Folder refresh thread (" & j + 1 & "/" & chuncks.Count & ") finished for " & Me.FullPath)

                                tcs2.SetResult()
                            End Sub))
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start(tcs)
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

        Dim enumEx As Exception = Nothing
        Dim tries As Integer = 0
        Do
            tries = tries + 1
            enumEx = Me.EnumerateItems(Me.ShellItem2, flags, cancellationToken,
                isSortPropertyByText, isSortPropertyDisplaySortValue, sortPropertyKey,
                result, newFullPaths, addItems, threadId, dupes)
            If Not enumEx Is Nothing Then Thread.Sleep(150)
        Loop While Not enumEx Is Nothing AndAlso tries < 3
        Me.EnumerationException = enumEx

        If Not cancellationToken.IsCancellationRequested Then
            ' add new items
            addItems()

            _isEnumerated = True
            _isEnumeratedForTree = True
            _isLoaded = True

            Me.OnItemsChanged()
        End If
    End Sub

    Protected Overridable Function EnumerateItems(shellItem2 As IShellItem2, flags As UInt32, cancellationToken As CancellationToken,
        isSortPropertyByText As Boolean, isSortPropertyDisplaySortValue As Boolean, sortPropertyKey As String,
        result As Dictionary(Of String, Item), newFullPaths As HashSet(Of String), addItems As System.Action,
        threadId As Integer?, dupes As List(Of Item)) As Exception

        Dim isDebuggerAttached As Boolean = Debugger.IsAttached
        Dim isRootDesktop As Boolean = If(Me.Pidl?.Equals(Shell.Desktop.Pidl), False)
        Dim replacedWithCustomFolders As HashSet(Of String) = New HashSet(Of String)()
        For Each customFolder In Shell.CustomFolders
            replacedWithCustomFolders.Add(customFolder.ReplacesFullPath.ToLower())
        Next

        Dim bindCtx As ComTypes.IBindCtx = Nothing, propertyBag As IPropertyBag = Nothing
        Dim var As PROPVARIANT, enumShellItems As IEnumShellItems = Nothing
        Try
            Functions.CreateBindCtx(0, bindCtx)
            Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBag)
            var.vt = VarEnum.VT_UI4
            var.union.uintVal = flags
            propertyBag.Write("SHCONTF", var) '  STR_ENUM_ITEMS_FLAGS 
            bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag) ' STR_PROPERTYBAG_PARAM 
            SyncLock _shellItemLockEnumShellItems
                If Not shellItem2 Is Nothing Then
                    CType(shellItem2, IShellItem2ForIEnumShellItems).BindToHandler _
                        (bindCtx, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, enumShellItems)
                End If
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
                                            newFullPaths.Add(newItem.FullPath & newItem.DeDupeKey)
                                        Catch ex As Exception
                                            isAlreadyAdded = True
                                        End Try
                                    End If
                                    If isAlreadyAdded Then
                                        Try
                                            result.Add(newItem.Pidl.ToString() & newItem.DeDupeKey, newItem)
                                            newFullPaths.Add(newItem.Pidl.ToString() & newItem.DeDupeKey)
                                            dupes.Add(newItem)
                                        Catch ex As Exception
                                            isAlreadyAdded = True
                                        End Try
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
                End If
            End If
            Return Nothing
        Catch ex As COMException
            If Not (ex.HResult = HRESULT.S_OK OrElse ex.HResult = HRESULT.S_FALSE OrElse ex.HResult = HRESULT.ERROR_INVALID_PARAMETER) Then
                Return ex
            End If
            Return Nothing
        Catch ex As Exception
            Return ex
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
    End Function

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

    Public Property ActiveView As Guid?
        Get
            Return _activeView
        End Get
        Set(value As Guid?)
            SetValue(_activeView, value)
        End Set
    End Property

    Public Property DefaultView As Guid
        Get
            Return _defaultView
        End Get
        Set(value As Guid)
            SetValue(_defaultView, value)
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

    Friend Overridable Sub OnItemsChanged(Optional item As Item = Nothing)
        Me.IsEmpty = _items.Count = 0
        ' set and update HasSubFolders property
        If item Is Nothing OrElse item.CanShowInTree Then
            Me.HasSubFolders = Not Me.Items.ToList().FirstOrDefault(Function(i) i.CanShowInTree) Is Nothing
        End If
    End Sub

    Protected Friend Overrides Sub ProcessNotification(e As NotificationEventArgs)
        If Not disposedValue Then
            Select Case e.Event
                Case SHCNE.RENAMEFOLDER, SHCNE.RENAMEITEM
                    If (Not e.Item2.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso e.Item2.Parent?.Pidl?.Equals(Me.Pidl)) _
                        OrElse (e.Item2.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso IO.Path.GetDirectoryName(e.Item2.FullPath)?.Equals(Me.FullPath)) _
                        OrElse IO.Path.GetDirectoryName(e.Item2.FullPath)?.Equals(_hookFolderFullPath) Then
                        _wasActivity = True
                        Dim existing As Item = Nothing
                        SyncLock _items.Lock
                            Dim newItem As Item = e.Item2.Clone()
                            newItem = Me.InitializeItem(newItem)
                            If Not newItem Is Nothing Then
                                existing = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                    AndAlso ((Not i.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Not i.Pidl Is Nothing AndAlso Not e.Item1.Pidl Is Nothing AndAlso i.Pidl?.Equals(e.Item1.Pidl)) _
                                        OrElse (Not i.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Not i.Pidl Is Nothing AndAlso Not newItem.Pidl Is Nothing AndAlso i.Pidl?.Equals(newItem.Pidl)) _
                                        OrElse ((i.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse i.Pidl Is Nothing OrElse e.Item1.Pidl Is Nothing) AndAlso i.FullPath?.Equals(e.Item1.FullPath)) _
                                        OrElse ((i.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse i.Pidl Is Nothing OrElse newItem.Pidl Is Nothing) AndAlso i.FullPath?.Equals(newItem.FullPath))))
                                If existing Is Nothing Then
                                    newItem.LogicalParent = Me
                                    newItem.IsProcessingNotifications = True
                                    Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                                    _items.InsertSorted(newItem, c)
                                    SyncLock _previousFullPathsLock
                                        If Not _previousFullPaths.Contains(newItem.FullPath) Then
                                            _previousFullPaths.Add(newItem.FullPath)
                                        End If
                                    End SyncLock
                                    Me.OnItemsChanged()
                                End If
                            End If
                        End SyncLock
                    End If
            End Select
        End If

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
                        If ((Not e.Item1.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse Not Me.Attributes.HasFlag(SFGAO.FILESYSTEM)) AndAlso e.Item1.Parent?.Pidl?.Equals(Me.Pidl)) _
                            OrElse (e.Item1.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Me.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(Me.FullPath)) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(_hookFolderFullPath) Then
                            _wasActivity = True
                            Dim existing As Item = Nothing, newItem As Item = Nothing
                            newItem = Me.InitializeItem(e.Item1.Clone())
                            If Not newItem Is Nothing Then
                                SyncLock _items.Lock
                                    existing = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                        AndAlso (Not i.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Not i.Pidl Is Nothing AndAlso Not newItem.Pidl Is Nothing AndAlso i.Pidl?.Equals(newItem.Pidl) _
                                            OrElse ((i.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse i.Pidl Is Nothing OrElse newItem.Pidl Is Nothing) AndAlso i.FullPath?.Equals(newItem.FullPath))))
                                    If existing Is Nothing Then
                                        newItem.LogicalParent = Me
                                        newItem.IsProcessingNotifications = True
                                        Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                                        _items.InsertSorted(newItem, c)
                                        SyncLock _previousFullPathsLock
                                            If Not _previousFullPaths.Contains(newItem.FullPath) Then
                                                _previousFullPaths.Add(newItem.FullPath)
                                            End If
                                        End SyncLock
                                        Me.OnItemsChanged()
                                    End If
                                End SyncLock
                                Shell.GlobalThreadPool.Run(
                                    Sub()
                                        newItem.Refresh()
                                    End Sub)
                            End If
                        End If
                    End If
                Case SHCNE.RMDIR, SHCNE.DELETE
                    If _isLoaded Then
                        If ((Not e.Item1.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse Not Me.Attributes.HasFlag(SFGAO.FILESYSTEM)) AndAlso e.Item1.Parent?.Pidl?.Equals(Me.Pidl)) _
                            OrElse (e.Item1.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Me.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(Me.FullPath)) _
                            OrElse IO.Path.GetDirectoryName(e.Item1.FullPath)?.Equals(_hookFolderFullPath) Then
                            _wasActivity = True
                            SyncLock _items.Lock
                                Dim existing As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                    AndAlso (Not i.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso Not i.Pidl Is Nothing AndAlso Not e.Item1.Pidl Is Nothing AndAlso i.Pidl?.Equals(e.Item1.Pidl) _
                                        OrElse ((i.Attributes.HasFlag(SFGAO.FILESYSTEM) OrElse i.Pidl Is Nothing OrElse e.Item1.Pidl Is Nothing) AndAlso i.FullPath?.Equals(e.Item1.FullPath))))
                                If Not existing Is Nothing Then
                                    existing.Dispose()
                                    Me.OnItemsChanged()
                                End If
                            End SyncLock
                        End If
                    End If
                Case SHCNE.DRIVEADD
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        _wasActivity = True
                        Dim newItem As Item = Nothing
                        newItem = Me.InitializeItem(e.Item1.Clone())
                        If Not newItem Is Nothing Then
                            SyncLock _items.Lock
                                Dim existing As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                    AndAlso (Not i.Pidl Is Nothing AndAlso Not newItem.Pidl Is Nothing AndAlso i.Pidl?.Equals(newItem.Pidl) _
                                        OrElse ((i.Pidl Is Nothing OrElse newItem.Pidl Is Nothing) AndAlso i.FullPath?.Equals(newItem.FullPath))))
                                If existing Is Nothing Then
                                    newItem.LogicalParent = Me
                                    newItem.IsProcessingNotifications = True
                                    Dim c As IComparer = New Helpers.ItemComparer(Me.ItemsGroupByPropertyName, Me.ItemsSortPropertyName, Me.ItemsSortDirection)
                                    _items.InsertSorted(newItem, c)
                                    Me.OnItemsChanged()
                                End If
                            End SyncLock
                            Shell.GlobalThreadPool.Run(
                                Sub()
                                    newItem.Refresh()
                                End Sub)
                        End If
                    End If
                Case SHCNE.DRIVEREMOVED
                    If Me.FullPath.Equals("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") AndAlso _isLoaded Then
                        _wasActivity = True
                        SyncLock _items.Lock
                            Dim existing As Item = _items.ToList().FirstOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                AndAlso (Not i.Pidl Is Nothing AndAlso Not e.Item1.Pidl Is Nothing AndAlso i.Pidl?.Equals(e.Item1.Pidl) _
                                    OrElse ((i.Pidl Is Nothing OrElse e.Item1.Pidl Is Nothing) AndAlso i.FullPath?.Equals(e.Item1.FullPath))))
                            If Not existing Is Nothing AndAlso TypeOf existing Is Folder Then
                                existing.Dispose()
                                Me.OnItemsChanged()
                            End If
                        End SyncLock
                    End If
                Case SHCNE.UPDATEDIR, SHCNE.UPDATEITEM
                    If _isLoaded Then
                        If (Me.Pidl?.Equals(e.Item1?.Pidl) OrElse Me.FullPath?.Equals(e.Item1?.FullPath) OrElse _hookFolderFullPath?.Equals(e.Item1?.FullPath) _
                                OrElse (Shell.Desktop.Pidl.Equals(e.Item1.Pidl) AndAlso _wasActivity)) _
                                AndAlso (e.Event = SHCNE.UPDATEDIR OrElse Not Me.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) _
                                    OrElse _wasActivity OrElse Not _isEnumerated OrElse Not _isEnumeratedForTree) Then
                            If (Me.IsExpanded OrElse Me.IsActiveInFolderView OrElse Me.IsVisibleInAddressBar) _
                                AndAlso Not TypeOf Me Is SearchFolder Then
                                _isEnumerated = False
                                _isEnumeratedForTree = False
                                Dim __ = Me.GetItemsAsync()
                            End If
                        End If
                    End If
            End Select

            If Not _notificationThreadPool Is Nothing Then
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
                                    item.ProcessNotification(e)
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

    Public ReadOnly Property Views As List(Of FolderViewRegistration)
        Get
            Return _views
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        Dim oldShellFolder As IShellFolder = Nothing
        Dim wasDisposed As Boolean = disposedValue

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

        If Not Shell.ShuttingDownToken.IsCancellationRequested AndAlso Not wasDisposed Then
            ' dispose outside of the lock because it can take a while
            Shell.DisposerThreadPool.Add(
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
                End Sub)
        End If
    End Sub

    Public Class FolderViewRegistration
        Public Property Guid As Guid
        Public Property Title As String
        Public Property IconUri As Uri
        Public Property Type As Type

        Public Sub New(guid As Guid, title As String, iconUri As Uri, type As Type)
            Me.Guid = guid
            Me.Title = title
            Me.IconUri = iconUri
            Me.Type = type
        End Sub
    End Class
End Class
