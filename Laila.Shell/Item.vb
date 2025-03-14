Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Public Class Item
    Inherits NotifyPropertyChangedBase
    Implements IDisposable
    Implements IProcessNotifications

    Public Event Refreshed As EventHandler

    Protected Const MAX_PATH_LENGTH As Integer = 260

    Friend _propertiesByKey As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Protected _propertiesByCanonicalName As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Friend _fullPath As String
    Friend disposedValue As Boolean
    Friend _parent As Folder
    Friend _displayName As String
    Private _isPinned As Boolean
    Private _isCut As Boolean
    Private _attributes As SFGAO
    Friend _preloadedAttributes As SFGAO
    Private _preloadedAttributesMask As SFGAO = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANRENAME Or SFGAO.CANDELETE
    Private _treeRootIndex As Long = -1
    Friend _shellItem2 As IShellItem2
    Friend _objectId As Long = -1
    Private Shared _objectCount As Long = 0
    Friend _pidl As Pidl
    Private _isImage As Boolean?
    Private _propertiesLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Friend _shellItemLock As Object = New Object()
    Protected _doKeepAlive As Boolean
    Private _contentViewModeProperties() As [Property]
    Private _isVisibleInAddressBar As Boolean
    Private _treeSortPrefix As String = String.Empty
    Friend _logicalParent As Folder
    Private _itemNameDisplaySortValuePrefix As String
    Protected _canShowInTree As Boolean
    Friend _livesOnThreadId As Integer?
    Private _isProcessingNotifications As Boolean
    Private _notifier As INotify
    Private _isInstantiated As Boolean

    ''' <summary>
    ''' Makes a new Folder/Item/Link object from a parsing name.
    ''' </summary>
    ''' <param name="parsingName">The parsing name of the item</param>
    ''' <param name="logicalParent">The logical parent folder for this item</param>
    ''' <param name="doKeepAlive">Set to 'true' to avoid this item getting disposed after 10 seconds if it isn't visible in the tree, folder view or address bar.</param>
    ''' <param name="doHookUpdates">Set to 'true' if you want this item to listen for update notifications while it's visible in the tree, folder view or address bar.</param>
    ''' <returns>A new Folder/Item/Link object</returns>
    Public Shared Function FromParsingName(parsingName As String, logicalParent As Folder,
                                           Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True) As Item
        Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId() ' select a random thread we're going to live on
        Return Shell.GlobalThreadPool.Run(
            Function() As Item
                ' is there a custom Folder implementation for this parsing name?
                Dim customFolderType As Type = Shell.CustomFolders _
                    .FirstOrDefault(Function(f) f.FullPath.ToLower().Equals(parsingName.ToLower()) _
                                         OrElse ("shell:" & f.FullPath.ToLower()).Equals(parsingName.ToLower()))?.Type
                If customFolderType Is Nothing Then
                    ' no custom implementation
                    parsingName = Environment.ExpandEnvironmentVariables(parsingName) ' expand environment variables
                    Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName) 'get the IShellItem2
                    If Not shellItem2 Is Nothing Then
                        ' read the attributes and decide what type of item we're dealing with
                        Dim attr As SFGAO = SFGAO.FOLDER Or SFGAO.LINK
                        shellItem2.GetAttributes(attr, attr)
                        If attr.HasFlag(SFGAO.FOLDER) Then
                            Return New Folder(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId)
                        ElseIf attr.HasFlag(SFGAO.LINK) Then
                            Return New Link(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId)
                        Else
                            Return New Item(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId)
                        End If
                    Else
                        ' there is no item with this parsing name
                        Return Nothing
                    End If
                Else
                    ' return custom folder implementation (like for example the 'Home' folder)
                    Return CType(Activator.CreateInstance(customFolderType, {logicalParent, doKeepAlive}), Folder)
                End If
            End Function,, threadId)
    End Function

    ''' <summary>
    ''' Makes a new Folder/Item/Link object from a PIDL.
    ''' </summary>
    ''' <param name="pidl">The pidl for this item. It will be cloned so you'll have to dispose this one yourself.</param>
    ''' <param name="logicalParent">The logical parent folder for this item</param>
    ''' <param name="doKeepAlive">Set to 'true' to avoid this item getting disposed after 10 seconds if it isn't visible in the tree, folder view or address bar.</param>
    ''' <param name="doHookUpdates">Set to 'true' if you want this item to listen for update notifications while it's visible in the tree, folder view or address bar.</param>
    ''' <param name="preservePidl">Set to 'true' to keep the given pidl as the value for the 'Pidl' property instead of disposing it and asking the system.</param>
    ''' <returns>A new Folder/Item/Link object</returns>
    Public Shared Function FromPidl(pidl As Pidl, logicalParent As Folder,
                                    Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True, Optional preservePidl As Boolean = False) As Item
        ' check input
        If pidl Is Nothing OrElse pidl.AbsolutePIDL.Equals(IntPtr.Zero) Then
            Throw New ArgumentException("pidl must not be null")
        End If
        Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId() ' get a random thread we're going to live on
        Return Shell.GlobalThreadPool.Run(
            Function() As Item
                Dim pidlClone As Pidl = pidl.Clone() ' clone the input pidl so we don't interfere with the original
                Dim shellItem2 As IShellItem2
                shellItem2 = GetIShellItem2FromPidl(pidlClone.AbsolutePIDL, logicalParent?.ShellFolder) ' get the IShellItem2
                If Not preservePidl Then pidlClone.Dispose() ' dispose the clone if we're not going to keep it
                If Not shellItem2 Is Nothing Then
                    ' read the attributes and decide what type of item we're dealing with
                    Dim attr As SFGAO = SFGAO.FOLDER Or SFGAO.LINK
                    shellItem2.GetAttributes(attr, attr)
                    If attr.HasFlag(SFGAO.FOLDER) Then
                        Return New Folder(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, If(preservePidl, pidlClone, Nothing))
                    ElseIf attr.HasFlag(SFGAO.LINK) Then
                        Return New Link(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, If(preservePidl, pidlClone, Nothing))
                    Else
                        Return New Item(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, If(preservePidl, pidlClone, Nothing))
                    End If
                Else
                    Return Nothing
                End If
            End Function,, threadId)
    End Function

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr, parentShellFolder As IShellFolder) As IShellItem2
        Dim result As IShellItem2 = Nothing
        If parentShellFolder Is Nothing Then
            Functions.SHCreateItemFromIDList(pidl, Guids.IID_IShellItem2, result)
        Else
            Functions.SHCreateItemWithParent(IntPtr.Zero, parentShellFolder, pidl, Guids.IID_IShellItem2, result)
        End If
        Return result
    End Function

    Friend Shared Function GetIShellItem2FromParsingName(parsingName As String) As IShellItem2
        Dim result As IShellItem2 = Nothing
        Functions.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, Guids.IID_IShellItem2, result)
        Return result
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, threadId As Integer?, Optional pidl As Pidl = Nothing)
        _objectCount += 1
        _objectId = _objectCount
        _shellItem2 = shellItem2
        _doKeepAlive = doKeepAlive
        Me.IsProcessingNotifications = doHookUpdates
        Me.LogicalParent = logicalParent
        _pidl = pidl
        _livesOnThreadId = threadId
        If Not shellItem2 Is Nothing Then
            shellItem2.GetAttributes(_preloadedAttributesMask, _preloadedAttributes)
            Dim d As String = Me.DisplayName
            Shell.AddToItemsCache(Me)
        Else
            _fullPath = String.Empty
        End If
        AddHandler Shell.Settings.PropertyChanged, AddressOf Settings_PropertyChanged
        _isInstantiated = True
    End Sub

    ''' <summary>
    ''' This constructor is to support referencing deleted or renamed files through the FileSystemWatcher only.
    ''' </summary>
    ''' <param name="fullPath"></param>
    Friend Sub New(fullPath As String)
        _objectCount += 1
        _objectId = _objectCount
        _fullPath = fullPath
        _displayName = IO.Path.GetFileName(_fullPath)
    End Sub

    ''' <summary>
    ''' The IShellItem2 this object is based on.
    ''' </summary>
    ''' <returns>The IShellItem2 this object is based on.</returns>
    Public ReadOnly Property ShellItem2 As IShellItem2
        Get
            Return _shellItem2
        End Get
    End Property

    ''' <summary>
    ''' Returns the PIDL for this item, or null if this item is disposed.
    ''' </summary>
    ''' <returns>The PIDL for this item, or null if this item is disposed</returns>
    Public ReadOnly Property Pidl As Pidl
        Get
            SyncLock _shellItemLock
                If _pidl Is Nothing AndAlso Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
                    Dim pidlptr As IntPtr
                    Functions.SHGetIDListFromObject(_shellItem2, pidlptr)
                    _pidl = New Pidl(pidlptr)
                End If
                Return _pidl ' return pidl within lock to make sure it's nothing after it's been disposed
            End SyncLock
        End Get
    End Property

    Protected Friend Property IsVisibleInAddressBar As Boolean
        Get
            Return _isVisibleInAddressBar
        End Get
        Set(value As Boolean)
            SetValue(_isVisibleInAddressBar, value)
        End Set
    End Property

    Protected Friend ReadOnly Property IsVisibleInTree As Boolean
        Get
            Return Me.TreeRootIndex <> -1 _
                OrElse (Me.CanShowInTree AndAlso If(If(_logicalParent, _parent)?.IsExpanded, False))
        End Get
    End Property

    Protected Friend Property CanShowInTree As Boolean
        Get
            Return _canShowInTree
        End Get
        Set(value As Boolean)
            SetValue(_canShowInTree, value)
            Me.NotifyOfPropertyChange("IsVisibleInTree")
            Me.NotifyOfPropertyChange("IsReadyForDispose")
        End Set
    End Property

    Protected Friend Overridable ReadOnly Property IsReadyForDispose As Boolean
        Get
            Return Not _doKeepAlive _
                AndAlso (If(_logicalParent, _parent) Is Nothing OrElse (Not If(_logicalParent, _parent).IsActiveInFolderView)) _
                AndAlso Not Me.IsVisibleInTree AndAlso Not Me.IsVisibleInAddressBar
        End Get
    End Property

    Protected Friend Overridable Sub MaybeDispose()
        If Me.IsReadyForDispose Then
            Me.Dispose()
        End If
    End Sub

    Protected Friend Property TreeViewSection As BaseTreeViewSection

    Protected Friend Property TreeRootIndex As Long
        Get
            Return _treeRootIndex
        End Get
        Set(value As Long)
            SetValue(_treeRootIndex, value)
            Me.NotifyOfPropertyChange("TreeSortKey")
        End Set
    End Property

    Protected Friend Property TreeSortPrefix As String
        Get
            Return _treeSortPrefix
        End Get
        Set(value As String)
            SetValue(_treeSortPrefix, value)
            Me.NotifyOfPropertyChange("TreeSortKey")
        End Set
    End Property

    Public ReadOnly Property TreeSortKey As String
        Get
            If _treeRootIndex <> -1 Then
                Return String.Format("{0:0000000000000000000}", _treeRootIndex)
            Else
                Dim itemNameDisplaySortValue As String = Me.ItemNameDisplaySortValue
                If itemNameDisplaySortValue Is Nothing Then itemNameDisplaySortValue = ""
                Return Me.LogicalParent?.TreeSortKey & Me.TreeSortPrefix & itemNameDisplaySortValue & New String(" ", 260 - itemNameDisplaySortValue.Length)
            End If
        End Get
    End Property

    Public ReadOnly Property TreeMargin As Thickness
        Get
            Dim level As Integer = 0
            If Me.TreeRootIndex = -1 Then
                Dim lp As Folder = Me.LogicalParent
                While Not lp Is Nothing
                    level += 1
                    If lp.TreeRootIndex <> -1 Then Exit While
                    lp = lp.LogicalParent
                End While
            End If

            Return New Thickness(level * 16, 0, 0, 0)
        End Get
    End Property

    Public Property DeDupeKey As String = String.Empty

    Protected Overridable Function GetNewShellItem() As IShellItem2
        Dim result As IShellItem2 = Nothing
        If Not Me.Pidl Is Nothing Then
            Functions.SHCreateItemFromIDList(Me.Pidl.AbsolutePIDL, GetType(IShellItem2).GUID, result)
        End If
        If result Is Nothing Then
            Functions.SHCreateItemFromParsingName(Me.FullPath, IntPtr.Zero, GetType(IShellItem2).GUID, result)
        End If
        If Not result Is Nothing Then
            result.Update(IntPtr.Zero)
            Return result
        End If
        Return Nothing
    End Function

    Public ReadOnly Property TypeAsString As String
        Get
            Return Me.GetType().ToString()
        End Get
    End Property

    Public Overridable Sub Refresh(Optional newShellItem As IShellItem2 = Nothing,
                                   Optional newPidl As Pidl = Nothing,
                                   Optional newFullPath As String = Nothing,
                                   Optional doRefreshImage As Boolean = True)
        Dim oldPropertiesByKey As Dictionary(Of String, [Property]) = Nothing
        Dim oldPropertiesByCanonicalName As Dictionary(Of String, [Property]) = Nothing
        Dim oldItemNameDisplaySortValue As String = Nothing
        Dim oldAttr As SFGAO = _attributes

        SyncLock _shellItemLock
            If Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
                oldItemNameDisplaySortValue = Me.ItemNameDisplaySortValue

                If Not _pidl Is Nothing Then
                    _pidl.Dispose()
                    _pidl = Nothing
                End If
                If Not newPidl Is Nothing Then
                    _pidl = newPidl
                End If
                If Not newFullPath Is Nothing Then
                    _fullPath = newFullPath
                Else
                    _fullPath = Nothing
                End If

                Dim oldShellItem As IShellItem2 = _shellItem2
                If Not newShellItem Is Nothing Then
                    _shellItem2 = newShellItem
                Else
                    _shellItem2 = Me.GetNewShellItem()
                End If
                If Not oldShellItem Is Nothing Then
                    Marshal.ReleaseComObject(oldShellItem)
                    oldShellItem = Nothing
                End If

                _propertiesLock.Wait()
                Try
                    oldPropertiesByKey = _propertiesByKey
                    oldPropertiesByCanonicalName = _propertiesByCanonicalName
                    _propertiesByKey = New Dictionary(Of String, [Property])()
                    _propertiesByCanonicalName = New Dictionary(Of String, [Property])()
                    _contentViewModeProperties = Nothing
                    For Each [property] In oldPropertiesByKey.Values
                        If Not [property].IsCustom Then
                            [property].Dispose()
                        Else
                            _propertiesByKey.Add([property].Key.ToString(), [property])
                        End If
                    Next
                    For Each [property] In oldPropertiesByCanonicalName.Values
                        [property].Dispose()
                    Next
                Finally
                    _propertiesLock.Release()
                End Try

                _displayName = Nothing

                If Not Me.ShellItem2 Is Nothing Then
                    _fullPath = Me.FullPath
                    _attributes = 0
                    _attributes = Me.Attributes

                    ' preload System_StorageProviderUIStatus images
                    'Dim System_StorageProviderUIStatus As System_StorageProviderUIStatusProperty _
                    '    = Me.PropertiesByKey(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey)
                    'If Not System_StorageProviderUIStatus Is Nothing _
                    '            AndAlso System_StorageProviderUIStatus.RawValue.vt <> 0 Then
                    '    Dim imgrefs As String() = System_StorageProviderUIStatus.ImageReferences16
                    'End If
                Else
                    Me.Dispose()
                End If
            Else
                Debug.WriteLine(Me.FullPath & "  " & disposedValue)
            End If
        End SyncLock


        If Not _logicalParent Is Nothing _
            AndAlso (oldAttr.HasFlag(SFGAO.FOLDER) <> _attributes.HasFlag(SFGAO.FOLDER) _
                     OrElse oldAttr.HasFlag(SFGAO.LINK) <> _attributes.HasFlag(SFGAO.LINK) _
                     OrElse TypeOf Me Is Folder AndAlso Not _attributes.HasFlag(SFGAO.FOLDER) _
                     OrElse Not TypeOf Me Is Folder AndAlso _attributes.HasFlag(SFGAO.FOLDER)) Then
            Dim newItem As Item = Me.Clone()
            newItem.LogicalParent = _logicalParent
            UIHelper.OnUIThread(
                Sub()
                    Dim c As IComparer = New Helpers.ItemComparer(Me.LogicalParent.ItemsGroupByPropertyName, Me.LogicalParent.ItemsSortPropertyName, Me.LogicalParent.ItemsSortDirection)
                    Me.LogicalParent._items.InsertSorted(newItem, c)
                    Me.LogicalParent.IsEmpty = Me.LogicalParent._items.Count = 0
                End Sub)
            Me.Dispose()
        End If

        If Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
            Me.NotifyOfPropertyChange("DisplayName")
            If Me.ItemNameDisplaySortValue <> oldItemNameDisplaySortValue Then
                Me.NotifyOfPropertyChange("ItemNameDisplaySortValue")
            End If
            If doRefreshImage Then
                Me.NotifyOfPropertyChange("OverlayImageAsync")
                Me.NotifyOfPropertyChange("IconAsync")
                Me.NotifyOfPropertyChange("ImageAsync")
                Me.NotifyOfPropertyChange("HasThumbnailAsync")
                Me.NotifyOfPropertyChange("IsImage")
            End If
            Me.NotifyOfPropertyChange("PropertiesByKeyAsText")
            Me.NotifyOfPropertyChange("IsHidden")
            Me.NotifyOfPropertyChange("IsCompressed")
            Me.NotifyOfPropertyChange("IsEncrypted")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusFirstIcon16Async")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusHasIcon")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusIcons16Async")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusIconWidth12")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusIconWidth16")
            Me.NotifyOfPropertyChange("ContentViewModeProperties")
            Me.NotifyOfPropertyChange("TileViewProperties")
            If Not oldPropertiesByKey Is Nothing Then
                For Each prop In oldPropertiesByKey
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}]", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Text", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].GroupByText", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].FirstIcon16Async", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].HasIconAsync", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Icons16Async", prop.Key))
                Next
            End If
            If Not oldPropertiesByCanonicalName Is Nothing Then
                For Each prop In oldPropertiesByCanonicalName
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}]", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Text", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].GroupByText", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].FirstIcon16Async", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].HasIconAsync", prop.Key))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Icons16Async", prop.Key))
                Next
            End If
        End If

        RaiseEvent Refreshed(Me, New EventArgs())
    End Sub

    ''' <summary>
    ''' The full path a.k.a. parsing name of this object.
    ''' </summary>
    ''' <returns>The full path a.k.a. parsing name of this object.</returns>
    Public Overridable ReadOnly Property FullPath As String
        Get
            SyncLock _shellItemLock
                If String.IsNullOrWhiteSpace(_fullPath) AndAlso Not disposedValue Then
                    Me.ShellItem2.GetDisplayName(SIGDN.DESKTOPABSOLUTEPARSING, _fullPath)
                End If
            End SyncLock

            Return _fullPath
        End Get
    End Property

    ''' <summary>
    ''' Gets/sets the logical parent for this item, which is the parent folder in the UI.
    ''' This is not necessarily the same as the actual parent folder in the file system.
    ''' </summary>
    ''' <returns>The logical parent folder for this item</returns>
    Public Property LogicalParent As Folder
        Get
            Return If(_logicalParent, Me.Parent)
        End Get
        Set(value As Folder)
            If Not _notifier Is Nothing Then
                _notifier.UnsubscribeFromNotifications(Me)
            ElseIf _isInstantiated Then
                Shell.UnsubscribeFromNotifications(Me)
            End If

            SetValue(_logicalParent, value)

            If Not _logicalParent Is Nothing Then
                _notifier = _logicalParent
                _notifier.SubscribeToNotifications(Me)
            Else
                _notifier = Nothing
                Shell.SubscribeToNotifications(Me)
            End If
        End Set
    End Property

    ''' <summary>
    ''' The actual parent folder of this object, as wired in the shell namespace.
    ''' </summary>
    ''' <returns>The actual parent folder of this object</returns>
    Public Overridable ReadOnly Property Parent As Folder
        Get
            ' if we're not disposed and we're not the desktop...
            If Not disposedValue AndAlso Not Me.FullPath?.Equals(Shell.Desktop.FullPath) Then
                If Not _parent Is Nothing Then ' we've still got a parent object
                    SyncLock _parent._shellItemLock
                        ' if still alive...
                        If Not _parent.disposedValue Then
                            ' extend lifetime
                            Shell.RemoveFromItemsCache(_parent)
                            Shell.AddToItemsCache(_parent)
                        Else
                            ' clear so we can start over
                            _parent = Nothing
                        End If
                    End SyncLock
                End If

                ' if we don't have any yet/anymore...
                If _parent Is Nothing Then
                    Dim parentShellItem2 As IShellItem2 = Nothing
                    Dim threadId As Integer = Shell.GlobalThreadPool.GetNextFreeThreadId() ' get an availble thread for this parent object
                    _parent = Shell.GlobalThreadPool.Run(
                        Function() As Folder
                            SyncLock _shellItemLock
                                If Not Me.ShellItem2 Is Nothing Then
                                    Me.ShellItem2.GetParent(parentShellItem2) ' get our parent folder, all objects must be created on a STA thread
                                End If
                            End SyncLock
                            If Not parentShellItem2 Is Nothing Then
                                Return New Folder(parentShellItem2, Nothing, False, True, threadId)
                            End If
                            Return Nothing
                        End Function, 1, threadId)
                End If
            End If
            Return _parent
        End Get
    End Property

    Public Property IsPinned As Boolean
        Get
            Return _isPinned
        End Get
        Set(value As Boolean)
            SetValue(_isPinned, value)
        End Set
    End Property

    ''' <summary>
    ''' Gets the executable path of the application associated with this item. 
    ''' We use this to get to the icon for the application which is shown in certain cases in the folder view.
    ''' </summary>
    ''' <returns>The executable path of the application associated with this item</returns>
    Public Function GetAssociatedApplication() As String
        Dim ext As String = IO.Path.GetExtension(Me.FullPath)

        Dim bufferSize As UInteger = 0
        ' First call to get the required buffer size
        Functions.AssocQueryStringW(AssocF.None, AssocStr.Executable, ext, Nothing, Nothing, bufferSize)

        If bufferSize = 0 Then Return Nothing

        Dim output As String = New String(ChrW(0), bufferSize)
        ' Second call to get the actual executable path
        If Functions.AssocQueryStringW(AssocF.None, AssocStr.Executable, ext, Nothing, output, bufferSize) = 0 Then
            Return output
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Gets the associated application icon for this item.
    ''' </summary>
    ''' <param name="size">The size of the icon in pixels</param>
    ''' <returns>The associated application icon for this item</returns>
    Public ReadOnly Property AssociatedApplicationIcon(size As Integer) As ImageSource
        Get
            If Not disposedValue AndAlso (IO.File.Exists(Me.FullPath) OrElse IO.Directory.Exists(Me.FullPath)) Then
                Dim app As String = Me.GetAssociatedApplication()
                If Not String.IsNullOrWhiteSpace(app) Then
                    Dim hBitmap As IntPtr
                    Try
                        Using icon As System.Drawing.Icon = Icon.ExtractAssociatedIcon(app.Trim(vbNullChar))
                            Using bitmap = icon.ToBitmap()
                                hBitmap = bitmap.GetHbitmap()
                                Dim image As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                image.Freeze()
                                Return image
                            End Using
                        End Using
                    Finally
                        If Not IntPtr.Zero.Equals(hBitmap) Then
                            Functions.DeleteObject(hBitmap)
                            hBitmap = IntPtr.Zero
                        End If
                    End Try
                End If
            End If
            Return Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the associated application icon for this item asynchronously.
    ''' </summary>
    ''' <param name="size">The requested icon size in pixels</param>
    ''' <returns>The associated application icon for this item</returns>
    Public Overridable ReadOnly Property AssociatedApplicationIconAsync(size As Integer) As ImageSource
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As ImageSource
                    Dim result As ImageSource
                    result = Me.AssociatedApplicationIcon(size)
                    If Not result Is Nothing Then result.Freeze()
                    Return result
                End Function, 3)
        End Get
    End Property

    Public ReadOnly Property OverlayIconIndex As Byte
        Get
            If Not disposedValue AndAlso (IO.File.Exists(Me.FullPath) OrElse IO.Directory.Exists(Me.FullPath)) Then
                Dim shFileInfo As New SHFILEINFO()
                Dim result As IntPtr = Functions.SHGetFileInfo(
                    Me.FullPath,
                    0,
                    shFileInfo,
                    Marshal.SizeOf(shFileInfo),
                    SHGFI.SHGFI_ICON Or SHGFI.SHGFI_OVERLAYINDEX
                )
                If Not IntPtr.Zero.Equals(shFileInfo.hIcon) Then
                    Functions.DestroyIcon(shFileInfo.hIcon)
                    shFileInfo.hIcon = IntPtr.Zero
                End If
                If Not IntPtr.Zero.Equals(result) Then
                    Return CByte((shFileInfo.iIcon >> 24) And &HFF)
                End If
            End If
            Return 0
        End Get
    End Property

    Public Overridable ReadOnly Property OverlayImage(size As Integer) As ImageSource
        Get
            Dim overlayIconIndex As Byte = Me.OverlayIconIndex
            If overlayIconIndex > 0 Then
                Return ImageHelper.GetOverlayIcon(overlayIconIndex, size)
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property OverlayImageAsync(size As Integer) As ImageSource
        Get
            Dim overlayIconIndex As Byte? = Shell.GlobalThreadPool.Run(
                Function() As Byte?
                    Return Me.OverlayIconIndex
                End Function, 3)

            If overlayIconIndex.HasValue AndAlso overlayIconIndex > 0 Then
                Return ImageHelper.GetOverlayIcon(overlayIconIndex, size)
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr
            Try
                SyncLock _shellItemLock
                    If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                        Dim h As HRESULT = HRESULT.S_FALSE, result As ImageSource = Nothing
                        If Not Settings.IsWindows8_1OrLower Then
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY Or SIIGBF.SIIGBF_INCACHEONLY, hbitmap)
                            If h = HRESULT.S_OK AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            End If
                        End If
                        If h <> HRESULT.S_OK OrElse IntPtr.Zero.Equals(hbitmap) _
                        OrElse (Not result Is Nothing AndAlso result.Width < size * Settings.DpiScaleX AndAlso result.Height < size * Settings.DpiScaleY) Then
                            If Not IntPtr.Zero.Equals(hbitmap) Then
                                Functions.DeleteObject(hbitmap)
                                hbitmap = IntPtr.Zero
                            End If
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY, hbitmap)
                            If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            Else
                                Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                            End If
                        End If
                        Return result
                    End If
                End SyncLock
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
            End Try
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property IconAsync(size As Integer) As ImageSource
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As ImageSource
                    Dim result As ImageSource
                    result = Me.Icon(size)
                    If Not result Is Nothing Then result.Freeze()
                    Return result
                End Function, 3)
        End Get
    End Property

    Public Overridable ReadOnly Property Image(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr
            Try
                If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                    SyncLock _shellItemLock
                        Dim h As HRESULT = HRESULT.S_FALSE, result As ImageSource = Nothing
                        If Not Settings.IsWindows8_1OrLower Then
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_INCACHEONLY, hbitmap)
                            If h = HRESULT.S_OK AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            End If
                        End If
                        If h <> HRESULT.S_OK OrElse IntPtr.Zero.Equals(hbitmap) _
                        OrElse (Not result Is Nothing AndAlso result.Width < size * Settings.DpiScaleX AndAlso result.Height < size * Settings.DpiScaleY) Then
                            If Not IntPtr.Zero.Equals(hbitmap) Then
                                Functions.DeleteObject(hbitmap)
                                hbitmap = IntPtr.Zero
                            End If
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), 0, hbitmap)
                            If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            Else
                                Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                            End If
                        End If
                        Return result
                    End SyncLock
                End If
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
            End Try
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property ImageAsync(size As Integer) As ImageSource
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As ImageSource
                    Dim result As ImageSource
                    result = Me.Image(size)
                    If Not result Is Nothing Then result.Freeze()
                    Return result
                End Function, 3)
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusIcons16Async As ImageSource()
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As ImageSource()
                    Dim result As ImageSource() = Nothing
                    If Not Me.disposedValue Then
                        result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.Icons16
                    End If
                    Return result
                End Function)
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusFirstIcon16Async As ImageSource
        Get
            Return Shell.GlobalThreadPool.Run(
                Function() As ImageSource
                    Dim result As ImageSource = Nothing
                    If Not Me.disposedValue Then
                        result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.FirstIcon16
                    End If
                    Return result
                End Function)
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusHasIcon As Boolean
        Get
            Return If(Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.HasIcon, False)
        End Get
    End Property

    Public ReadOnly Property StorageProviderUIStatusIconWidth12 As Double
        Get
            Return If(Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.ImageReferences16.Count, 0) * 12
        End Get
    End Property

    Public ReadOnly Property StorageProviderUIStatusIconWidth16 As Double
        Get
            Return If(Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.ImageReferences16.Count, 0) * 16
        End Get
    End Property

    Public Overridable ReadOnly Property HasThumbnail As Boolean
        Get
            Dim hbitmap As IntPtr
            Try
                If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                    Dim h As HRESULT
                    SyncLock _shellItemLock
                        h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(1, 1), SIIGBF.SIIGBF_THUMBNAILONLY, hbitmap)
                    End SyncLock
                    If h = HRESULT.S_OK AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                        Return True
                    End If
                End If
            Catch ex As Exception
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
            End Try
            Return False
        End Get
    End Property

    Public Overridable ReadOnly Property HasThumbnailAsync As Boolean
        Get
            Dim result As Boolean? = Shell.GlobalThreadPool.Run(
                Function() As Boolean?
                    Return Me.HasThumbnail
                End Function, 3)

            If result.HasValue Then
                Return result.Value
            Else
                Return False
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property IsImage As Boolean
        Get
            If Not _isImage.HasValue Then
                _isImage = ImageHelper.IsImage(Me.FullPath)
            End If
            Return _isImage.Value
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayName As String
        Get
            Dim isNew As Boolean

            ' get displayname?
            If String.IsNullOrWhiteSpace(_displayName) AndAlso Not disposedValue Then
                SyncLock _shellItemLock
                    If String.IsNullOrWhiteSpace(_displayName) AndAlso Not disposedValue Then
                        Me.ShellItem2.GetDisplayName(SHGDN.NORMAL, _displayName)
                        isNew = True
                    End If
                End SyncLock
            End If

            ' strip root name from displayname
            If isNew AndAlso Me.FullPath?.StartsWith(IO.Path.DirectorySeparatorChar & IO.Path.DirectorySeparatorChar) Then
                Dim idx As Integer = Me.FullPath.IndexOf(IO.Path.DirectorySeparatorChar, 2)
                If idx >= 0 Then
                    Dim rootDisplayName As String = Me.FullPath.Substring(2, idx - 2)
                    If _displayName.ToLower().EndsWith(String.Format("({0}{0}{1})", IO.Path.DirectorySeparatorChar, rootDisplayName.ToLower())) Then
                        _displayName = _displayName.Substring(0, _displayName.Length - rootDisplayName.Length - 5)
                    End If
                End If
            End If

            Return _displayName
        End Get
    End Property

    Protected Friend Property AddressBarRoot As String

    Protected Friend Property AddressBarDisplayName As String

    Public ReadOnly Property AddressBarDisplayPath As String
        Get
            If String.IsNullOrWhiteSpace(Me.AddressBarRoot) Then
                Dim parent As Item = Me
                Dim path As String
                If parent.IsDrive Then
                    path = If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                  parent.FullPath, parent.AddressBarDisplayName)
                Else
                    path = If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                  parent.DisplayName, parent.AddressBarDisplayName)
                End If
                If Not Shell.GetSpecialFolders().Values.ToList().Exists(Function(f) _
                    If(f.Pidl?.Equals(parent?.Pidl), False) OrElse (f?.Pidl Is Nothing AndAlso parent?.Pidl Is Nothing)) Then
                    parent = parent.LogicalParent
                    If parent Is Nothing Then
                        Dim i As Int16 = 9
                    End If
                    While Not parent Is Nothing
                        If parent.IsDrive Then
                            path = IO.Path.Combine(If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                                          parent.FullPath, parent.AddressBarDisplayName), path)
                        Else
                            path = IO.Path.Combine(If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                                          parent.DisplayName, parent.AddressBarDisplayName), path)
                        End If
                        If Shell.GetSpecialFolders().Values.ToList().Exists(Function(f) _
                            If(f.Pidl?.Equals(parent?.Pidl), False) OrElse (f?.Pidl Is Nothing AndAlso parent?.Pidl Is Nothing)) Then
                            Exit While
                        End If
                        parent = parent.LogicalParent
                    End While
                End If
                Return path
            Else
                Return IO.Path.Combine(Me.AddressBarRoot, If(String.IsNullOrWhiteSpace(Me.AddressBarDisplayName),
                                                              Me.DisplayName, Me.AddressBarDisplayName))
            End If
        End Get
    End Property

    Public ReadOnly Property IsDrive As Boolean
        Get
            Return Not Me.FullPath Is Nothing AndAlso Me.FullPath.Equals(Path.GetPathRoot(Me.FullPath)) AndAlso Not Me.FullPath.StartsWith("\\")
        End Get
    End Property

    Public ReadOnly Property IsFolder As Boolean
        Get
            Return TypeOf Me Is Folder
        End Get
    End Property

    Protected Friend Property ItemNameDisplaySortValuePrefix As String
        Get
            Return _itemNameDisplaySortValuePrefix
        End Get
        Set(value As String)
            SetValue(_itemNameDisplaySortValuePrefix, value)
            Me.NotifyOfPropertyChange("ItemNameDisplaySortValue")
        End Set
    End Property

    Public Overridable ReadOnly Property ItemNameDisplaySortValue As String
        Get
            If Me.IsDrive AndAlso Shell.Settings.DoShowDriveLetters Then
                Return String.Format("{0}{1}", Me.ItemNameDisplaySortValuePrefix, Me.FullPath)
            Else
                Return String.Format("{0}{1}{2}", Me.ItemNameDisplaySortValuePrefix, If(Me.IsFolder AndAlso Me.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), "0", "1"), Me.DisplayName)
            End If
        End Get
    End Property

    Public ReadOnly Property IsExisting As Boolean
        Get
            Dim attr As SFGAO = SFGAO.VALIDATE
            SyncLock _shellItemLock
                Return ShellItem2.GetAttributes(attr, attr) = HRESULT.S_FALSE
            End SyncLock
        End Get
    End Property

    Public Property IsCut As Boolean
        Get
            Return _isCut
        End Get
        Friend Set(value As Boolean)
            SetValue(_isCut, value)
        End Set
    End Property

    Public ReadOnly Property IsHidden As Boolean
        Get
            Return Me.Attributes.HasFlag(SFGAO.HIDDEN)
        End Get
    End Property

    Public ReadOnly Property IsCompressed As Boolean
        Get
            Return Me.Attributes.HasFlag(SFGAO.COMPRESSED)
        End Get
    End Property

    Public ReadOnly Property IsEncrypted As Boolean
        Get
            Return Me.Attributes.HasFlag(SFGAO.ENCRYPTED)
        End Get
    End Property

    Public Overridable ReadOnly Property Attributes As SFGAO
        Get
            If _attributes = 0 AndAlso Not disposedValue Then
                _attributes = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANLINK Or SFGAO.CANRENAME _
                                Or SFGAO.CANDELETE Or SFGAO.DROPTARGET Or SFGAO.ENCRYPTED Or SFGAO.ISSLOW _
                                Or SFGAO.LINK Or SFGAO.SHARE Or SFGAO.RDONLY Or SFGAO.HIDDEN Or SFGAO.FOLDER _
                                Or SFGAO.FILESYSTEM Or SFGAO.COMPRESSED Or SFGAO.STORAGEANCESTOR
                SyncLock _shellItemLock
                    ShellItem2.GetAttributes(_attributes, _attributes)
                End SyncLock
            End If
            Return _attributes
        End Get
    End Property

    Public Overridable Property IsLoading As Boolean
        Get
            Return False
        End Get
        Set(value As Boolean)
        End Set
    End Property

    Public Overridable Property IsExpanded As Boolean
        Get
            Return False
        End Get
        Set(value As Boolean)
        End Set
    End Property

    Public Overridable Property HasSubFolders As Boolean
        Get
            Return False
        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public ReadOnly Property ContentViewModeProperties As [Property]()
        Get
            If Not disposedValue AndAlso _contentViewModeProperties Is Nothing Then
                Dim propList As String
                If _parent Is Nothing OrElse Not TypeOf _parent Is SearchFolder Then
                    propList = Me.PropertiesByCanonicalName("System.PropList.ContentViewModeForBrowse").Text
                Else
                    propList = Me.PropertiesByCanonicalName("System.PropList.ContentViewModeForSearch").Text
                End If
                If Not String.IsNullOrWhiteSpace(propList) Then
                    Dim propertyNames() As String = propList.Substring(5).Split(";")
                    Dim properties As List(Of [Property]) = New List(Of [Property])()
                    For Each propCanonicalName In propertyNames
                        Dim prop As [Property] = Me.PropertiesByCanonicalName(propCanonicalName.TrimStart("~"))
                        If Not prop Is Nothing Then
                            properties.Add(prop)
                        Else
                            properties.Add(Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"))
                        End If
                    Next
                    While properties.Count < 6
                        properties.Add(Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"))
                    End While
                    _contentViewModeProperties = properties.ToArray()
                Else
                    _contentViewModeProperties = {
                        Me.PropertiesByCanonicalName("System.ItemNameDisplay"),
                        Me.PropertiesByCanonicalName("System.ItemTypeText"),
                        Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"),
                        Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"),
                        Me.PropertiesByCanonicalName("System.DateModified"),
                        Me.PropertiesByCanonicalName("System.Size")
                    }
                End If
            End If

            Return _contentViewModeProperties
        End Get
    End Property

    Public ReadOnly Property InfoTip As String
        Get
            If Not disposedValue Then
                Dim PKEY_System_InfoTipText As New PROPERTYKEY() With {
                    .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                    .pid = 4
                }
                Dim system_InfoTipText As String = Me.PropertiesByKey(PKEY_System_InfoTipText)?.Text
                Dim properties() As String
                If Not String.IsNullOrWhiteSpace(system_InfoTipText) Then
                    properties = system_InfoTipText.Substring(5).Split(";")
                Else
                    properties = {"System.ItemTypeText", "System.Size"}
                End If
                Dim text As List(Of String) = New List(Of String)()
                Dim i As Integer = 0
                For Each propCanonicalName In properties
                    Dim prop As [Property] = Me.PropertiesByCanonicalName(propCanonicalName)
                    If Not prop Is Nothing AndAlso Not String.IsNullOrWhiteSpace(prop.Text) Then
                        text.Add(prop.DisplayName & ": " & prop.Text)
                    End If
                    i += 1
                Next
                Dim System_StorageProviderUIStatus As System_StorageProviderUIStatusProperty _
                    = Me.PropertiesByKey(System_StorageProviderUIStatusProperty.Key)
                If System_StorageProviderUIStatus?.RawValue.vt <> 0 Then
                    If Not String.IsNullOrWhiteSpace(System_StorageProviderUIStatus.Text) Then
                        text.Add(System_StorageProviderUIStatus.DisplayName & ": " & System_StorageProviderUIStatus.Text)
                    End If
                    If Not String.IsNullOrWhiteSpace(System_StorageProviderUIStatus.ActivityText) Then
                        text.Add(System_StorageProviderUIStatus.ActivityDisplayName & ": " & System_StorageProviderUIStatus.ActivityText)
                    End If
                End If
                Return String.Join(vbCrLf, text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property TileViewProperties As String
        Get
            If Not disposedValue Then
                Dim PKEY_System_PropList_TileInfo As New PROPERTYKEY() With {
                    .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                    .pid = 3
                }
                Dim system_PropList_TileInfo As String = Me.PropertiesByKey(PKEY_System_PropList_TileInfo).Text
                Dim properties() As String
                If Not String.IsNullOrWhiteSpace(system_PropList_TileInfo) Then
                    properties = system_PropList_TileInfo.Substring(5).Split(";")
                Else
                    properties = {"System.ItemTypeText", "System.Size"}
                End If
                Dim text As List(Of String) = New List(Of String)()
                Dim i As Integer = 0
                For Each propCanonicalName In properties
                    If Not propCanonicalName.StartsWith("*") Then
                        Dim prop As [Property] = Me.PropertiesByCanonicalName(propCanonicalName)
                        If Not prop Is Nothing AndAlso Not String.IsNullOrWhiteSpace(prop.Text) Then
                            text.Add(If(i >= 2, prop.DisplayName & ": ", "") & prop.Text)
                        End If
                    End If
                    i += 1
                Next
                Return String.Join(vbCrLf, text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKeyAsText(propertyKey As String) As [Property]
        Get
            Dim [property] As [Property] = Nothing
            Dim key As PROPERTYKEY = New PROPERTYKEY(propertyKey)
            _propertiesLock.Wait()
            Try
                If Not _propertiesByKey.TryGetValue(key.ToString(), [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromKey(key, Me.ShellItem2)
                        Else
                            [property] = Nothing
                        End If
                    End SyncLock
                    _propertiesLock.Wait()
                    If Not [property] Is Nothing Then
                        If Not _propertiesByKey.ContainsKey(propertyKey) Then
                            _propertiesByKey.Add(key.ToString(), [property])
                        Else
                            [property].Dispose()
                            [property] = _propertiesByKey(key.ToString())
                        End If
                    End If
                End If
            Finally
                If _propertiesLock.CurrentCount = 0 Then
                    _propertiesLock.Release()
                End If
            End Try
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKey(propertyKey As PROPERTYKEY) As [Property]
        Get
            Dim [property] As [Property] = Nothing
            _propertiesLock.Wait()
            Try
                If Not _propertiesByKey.TryGetValue(propertyKey.ToString(), [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromKey(propertyKey, Me.ShellItem2)
                        Else
                            [property] = Nothing
                        End If
                    End SyncLock
                    _propertiesLock.Wait()
                    If Not [property] Is Nothing Then
                        If Not _propertiesByKey.ContainsKey(propertyKey.ToString()) Then
                            _propertiesByKey.Add(propertyKey.ToString(), [property])
                        Else
                            [property].Dispose()
                            [property] = _propertiesByKey(propertyKey.ToString())
                        End If
                    End If
                End If
            Finally
                If _propertiesLock.CurrentCount = 0 Then
                    _propertiesLock.Release()
                End If
            End Try
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByCanonicalName(canonicalName As String) As [Property]
        Get
            Dim [property] As [Property] = Nothing
            _propertiesLock.Wait()
            Try
                If Not _propertiesByCanonicalName.TryGetValue(canonicalName, [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromCanonicalName(canonicalName, Me.ShellItem2)
                        Else
                            [property] = Nothing
                        End If
                    End SyncLock
                    _propertiesLock.Wait()
                    If Not [property] Is Nothing Then
                        If Not _propertiesByCanonicalName.ContainsKey(canonicalName) Then
                            _propertiesByCanonicalName.Add(canonicalName, [property])
                        Else
                            [property].Dispose()
                            [property] = _propertiesByCanonicalName(canonicalName)
                        End If
                    End If
                End If
            Finally
                If _propertiesLock.CurrentCount = 0 Then
                    _propertiesLock.Release()
                End If
            End Try
            Return [property]
        End Get
    End Property

    Public Overridable Property IsProcessingNotifications As Boolean Implements IProcessNotifications.IsProcessingNotifications
        Get
            Return _isProcessingNotifications AndAlso (Me.IsVisibleInAddressBar OrElse Me.IsVisibleInTree OrElse If(Me._logicalParent?.IsActiveInFolderView, False))
        End Get
        Friend Set(value As Boolean)
            SetValue(_isProcessingNotifications, value)
        End Set
    End Property

    Public Shared Async Function FromParsingNameDeepGetAsync(parsingName As String) As Task(Of Item)
        ' resolve environment variable?
        parsingName = Environment.ExpandEnvironmentVariables(parsingName)

        Dim specialFolder As Folder =
            Shell.GetSpecialFolders().Values.ToList().FirstOrDefault(Function(f) f.FullPath = parsingName)
        If Not specialFolder Is Nothing Then
            Return specialFolder
        End If

        Dim path As String = parsingName.Trim()

        ' get parts of path
        Dim parts As List(Of String)
        Dim isNetworkPath As Boolean = path.StartsWith("\\")
        If isNetworkPath Then
            parts = path.Substring(2).Split(IO.Path.DirectorySeparatorChar).ToList()
            If parts.Count = 1 AndAlso parts(0).Length = 0 Then
                parts.RemoveAt(0)
            ElseIf parts.Count > 0 Then
                parts(0) = "\\" & parts(0)
            End If
        Else
            parts = New List(Of String)()
            While Not String.IsNullOrWhiteSpace(path)
                Debug.WriteLine(path)
                If path = IO.Path.GetPathRoot(path) Then
                    parts.Add(path)
                Else
                    parts.Add(IO.Path.GetFileName(path))
                End If
                path = IO.Path.GetDirectoryName(path)
            End While
            parts.Reverse()
        End If

        If parts.Count > 0 Or isNetworkPath Then
            Dim folder As Folder
            Dim j As Integer, start As Integer = 0

            If isNetworkPath Then
                ' network path
                folder = Shell.GetSpecialFolder(SpecialFolders.Network)
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.GetSpecialFolder(SpecialFolders.ThisPc)
            Else
                ' root must be some special folder
                folder = Shell.GetSpecialFolders().Values.FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(0).ToLower() _
                                                                         OrElse f.FullPath.ToLower() = parts(0).ToLower())
                start = 1
            End If

            ' find folder
            If Not folder Is Nothing Then
                For j = start To parts.Count - 1
                    Dim subFolder As Folder
                    If j = 0 Then
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) TypeOf f Is Folder AndAlso IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower() = parts(j).ToLower())
                    Else
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) TypeOf f Is Folder AndAlso IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower() = parts(j).ToLower())
                    End If
                    If subFolder Is Nothing Then
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) TypeOf f Is Folder AndAlso f.DisplayName.ToLower() = parts(j).ToLower())
                    End If
                    folder = subFolder
                    If folder Is Nothing Then Exit For
                Next
            End If

            Return folder
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function FromParsingNameDeepGet(parsingName As String) As Item
        ' resolve environment variable?
        parsingName = Environment.ExpandEnvironmentVariables(parsingName)

        Dim specialFolder As Folder =
            Shell.GetSpecialFolders().Values.ToList().FirstOrDefault(Function(f) f.FullPath = parsingName)
        If Not specialFolder Is Nothing Then
            Return specialFolder
        End If

        Dim path As String = parsingName.Trim()

        ' get parts of path
        Dim parts As List(Of String)
        Dim isNetworkPath As Boolean = path.StartsWith("\\")
        If isNetworkPath Then
            parts = path.Substring(2).Split(IO.Path.DirectorySeparatorChar).ToList()
            If parts.Count = 1 AndAlso parts(0).Length = 0 Then
                parts.RemoveAt(0)
            ElseIf parts.Count > 0 Then
                parts(0) = "\\" & parts(0)
            End If
        Else
            parts = New List(Of String)()
            While Not String.IsNullOrWhiteSpace(path)
                Debug.WriteLine(path)
                If path = IO.Path.GetPathRoot(path) Then
                    parts.Add(path)
                Else
                    parts.Add(IO.Path.GetFileName(path))
                End If
                path = IO.Path.GetDirectoryName(path)
            End While
            parts.Reverse()
        End If

        If parts.Count > 0 Or isNetworkPath Then
            Dim folder As Folder
            Dim j As Integer, start As Integer = 0

            If isNetworkPath Then
                ' network path
                folder = Shell.GetSpecialFolder(SpecialFolders.Network)
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.GetSpecialFolder(SpecialFolders.ThisPc)
            Else
                ' root must be some special folder
                folder = Shell.GetSpecialFolders().Values.FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(0).ToLower() _
                                                                         OrElse f.FullPath.ToLower() = parts(0).ToLower())
                start = 1
            End If

            ' find folder
            If Not folder Is Nothing Then
                For j = start To parts.Count - 1
                    Dim subFolder As Folder
                    If j = 0 Then
                        subFolder = (folder.GetItems()).FirstOrDefault(Function(f) IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower() = parts(j).ToLower())
                    Else
                        subFolder = (folder.GetItems()).FirstOrDefault(Function(f) IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower() = parts(j).ToLower())
                    End If
                    If subFolder Is Nothing Then
                        subFolder = (folder.GetItems()).FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(j).ToLower())
                    End If
                    folder = subFolder
                    If folder Is Nothing Then Exit For
                Next
            End If

            Return folder
        Else
            Return Nothing
        End If
    End Function

    Protected Friend Overridable Sub ProcessNotification(e As NotificationEventArgs) Implements IProcessNotifications.ProcessNotification
        If Not disposedValue Then
            Select Case e.Event
                Case SHCNE.UPDATEITEM, SHCNE.UPDATEDIR ' general update
                    If Me.Pidl?.Equals(e.Item1?.Pidl) Then ' if this is us...
                        Shell.GlobalThreadPool.Run(
                            Sub()
                                Me.Refresh() ' refresh
                            End Sub)
                    End If
                Case SHCNE.FREESPACE ' free space has changed, disk not specified
                    If Me.IsDrive Then
                        Shell.GlobalThreadPool.Run(
                            Sub()
                                Me.Refresh(,,, False) ' don't refresh image because there's times when it does this a lot and we want to avoid flicker
                            End Sub)
                    End If
                Case SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED ' cdrom has been inserted or removed
                    If Me.IsDrive AndAlso Me.Pidl?.Equals(e.Item1?.Pidl) Then
                        Shell.GlobalThreadPool.Run(
                            Sub()
                                Me.Refresh() ' refresh, as the icon and name may have changed
                            End Sub)
                    End If
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    ' all the following conditions are to support file operations in .zip folders
                    If (Not e.Item1?.Pidl Is Nothing AndAlso Me.Pidl?.Equals(e.Item1?.Pidl)) _  ' if this is our pidl
                        OrElse (Me.FullPath?.ToLower().Equals(e.Item1?.FullPath.ToLower())) _   ' or our fullpath
                        AndAlso Not (Me.IsFolder AndAlso Not Me.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) _ ' and we're not a zip folder being renamed...
                                     AndAlso e.Item1.Pidl Is Nothing AndAlso IO.Path.GetFileName(e.Item2?.FullPath)?.Contains("~") _ ' ..to...
                                     AndAlso IO.Path.GetExtension(e.Item2?.FullPath?.ToLower())?.Equals(".tmp")) Then   ' ...a temp file
                        Dim doRefresh As Boolean = True
                        If (e.Event = SHCNE.RENAMEFOLDER AndAlso e.Item2.IsFolder AndAlso Not e.Item2.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) _
                            AndAlso e.Item1.Pidl Is Nothing AndAlso Not _logicalParent Is Nothing) Then ' if we're being renamed to a .zip file
                            Dim existing As Item = Nothing
                            UIHelper.OnUIThread(
                                Sub()
                                    ' check if the new zip folder already exists in the parent
                                    existing = _logicalParent.Items.ToList().FirstOrDefault(Function(i) Not i.disposedValue _
                                                AndAlso (i.Pidl?.Equals(e.Item2.Pidl) OrElse i.FullPath?.Equals(e.Item2.FullPath)))
                                End Sub)
                            If Not existing Is Nothing AndAlso TypeOf existing Is Folder Then ' the folder exists...
                                doRefresh = False ' don't do regular refresh on rename
                                Shell.GlobalThreadPool.Run(
                                    Sub()
                                        Dim existingFolder As Folder = existing
                                        existingFolder.Refresh(e.Item2?.ShellItem2, e.Item2?.Pidl?.Clone(), e.Item2?.FullPath) ' refresh the folder itself
                                        e.Item2._shellItem2 = Nothing
                                        existingFolder._isEnumerated = False
                                        existingFolder.GetItemsAsync(True, True) ' refresh the contents
                                    End Sub)
                            End If
                        End If

                        ' regular refresh on rename
                        If doRefresh Then
                            Shell.GlobalThreadPool.Run(
                                Sub()
                                    Dim oldPidl As Pidl = Me.Pidl?.Clone() ' save old pidl
                                    Me.Refresh(e.Item2?.ShellItem2, e.Item2?.Pidl?.Clone(), e.Item2?.FullPath) ' refresh this item
                                    If Not oldPidl Is Nothing AndAlso Not Me.Pidl Is Nothing Then
                                        ' rename pinned and frequent items with the same pidl
                                        PinnedItems.RenameItem(oldPidl, Me.Pidl)
                                        FrequentFolders.RenameItem(oldPidl, Me.Pidl)
                                    End If
                                    If Not oldPidl Is Nothing Then
                                        oldPidl.Dispose() ' dispose of old pidl
                                    End If
                                    If Not e.Item2 Is Nothing Then
                                        ' we've used this shell item in item1 now, so avoid it getting disposed when item2 gets disposed
                                        e.Item2._shellItem2 = Nothing
                                    End If
                                End Sub)
                        End If
                    End If
            End Select
        End If
    End Sub

    Protected Overridable Sub Settings_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
        Select Case e.PropertyName
            Case "DoHideKnownFileExtensions"
                _displayName = Nothing
                Me.NotifyOfPropertyChange("DisplayName")
            Case "DoShowDriveLetters"
                If Me.IsDrive Then
                    Shell.GlobalThreadPool.Run(
                        Sub()
                            Me.Refresh()
                        End Sub)
                End If
        End Select
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        Dim wasDisposed As Boolean = disposedValue
        SyncLock _shellItemLock
            If Not disposedValue Then
                disposedValue = True
                'Debug.WriteLine("Disposing " & _objectId & ": " & Me.FullPath)

                If disposing Then
                    ' dispose managed state (managed objects):

                    If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                        ' extensive cleanup, because we're still live:

                        ' unsubscribe from notifications
                        If Not _notifier Is Nothing Then
                            _notifier.UnsubscribeFromNotifications(Me)
                        Else
                            Shell.UnsubscribeFromNotifications(Me)
                        End If
                        RemoveHandler Shell.Settings.PropertyChanged, AddressOf Settings_PropertyChanged

                        ' remove from parent collection
                        If Me.IsReadyForDispose Then
                            ' we're being disposed from the disposer thread, so try not to block the ui thread
                            If Not _logicalParent Is Nothing Then
                                _logicalParent._items.RemoveWithoutNotifying(Me)
                                _logicalParent._isEnumerated = False
                                _logicalParent.IsEmpty = _logicalParent._items.Count = 0
                                _logicalParent = Nothing
                            End If
                        Else
                            UIHelper.OnUIThreadAsync(
                            Sub()
                                If Not _logicalParent Is Nothing Then
                                    _logicalParent._items.Remove(Me)
                                    _logicalParent._isEnumerated = False
                                    _logicalParent.IsEmpty = _logicalParent._items.Count = 0
                                    ' don't set logicalparent to nothing, because we might still need it when the treeview removes folder's children
                                End If
                            End Sub)
                        End If

                        ' dispose properties
                        For Each [property] In _propertiesByKey.ToList()
                            [property].Value.Dispose()
                        Next
                        _propertiesByKey.Clear()
                        For Each [property] In _propertiesByCanonicalName.ToList()
                            [property].Value.Dispose()
                        Next
                        _propertiesByCanonicalName.Clear()

                        ' switch shellitem, we'll release it later
                        If Not _shellItem2 Is Nothing Then
                            Marshal.ReleaseComObject(_shellItem2)
                            _shellItem2 = Nothing
                        End If

                        ' remove from cache
                        Shell.RemoveFromItemsCache(Me)
                    Else
                        ' quick cleanup, because we're shutting down anyway and it has to go fast:

                        ' release shellitem
                        If Not _shellItem2 Is Nothing Then
                            Marshal.ReleaseComObject(_shellItem2)
                            _shellItem2 = Nothing
                        End If

                        ' dispose properties
                        For Each [property] In _propertiesByKey.ToList()
                            [property].Value.Dispose()
                        Next
                        For Each [property] In _propertiesByCanonicalName.ToList()
                            [property].Value.Dispose()
                        Next
                    End If

                    ' dispose pidl
                    If Not _pidl Is Nothing Then
                        _pidl.Dispose()
                        _pidl = Nothing
                    End If

                    ' free unmanaged resources (unmanaged objects) and override finalizer
                End If
            End If
        End SyncLock
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '    Dispose(disposing:=False)
    '    MyBase.Finalize()
    'End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Overridable Function Clone() As Item
        Return Item.FromPidl(Me.Pidl, Nothing, _doKeepAlive)
    End Function
End Class
