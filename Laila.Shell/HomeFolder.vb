Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports System.Runtime.InteropServices.ComTypes
Imports Laila.Shell.Events
Imports System.Collections.ObjectModel
Imports System.Windows.Controls
Imports Laila.Shell.Controls
Imports System.Windows.Data
Imports System.ComponentModel

Public Class HomeFolder
    Inherits Folder
    Implements ISupportDragInsert

    Private _cachedHasSubFolders As Boolean?

    Shared Sub New()
        Shell.GlobalThreadPool.Add(
            Sub()
                Shell.StartListening(Shell.GetSpecialFolder(SpecialFolders.Recent))
            End Sub)
    End Sub

    Public Sub New(parent As Folder, doKeepAlive As Boolean)
        MyBase.New(Nothing, parent, doKeepAlive, True, Nothing)

        _hasSubFolders = True
        _hasShellItem = False ' prevent getting disposed on refresh because we ain't got a shellitem
        Me.ItemsGroupByPropertyName = "GroupByText[" & Home_CategoryProperty.Key.ToString() & "]"
        Me.ItemsSortDirection = ComponentModel.ListSortDirection.Descending

        _views = New List(Of FolderViewRegistration) From {
            New FolderViewRegistration(New Guid("5aab6a71-6d79-4bcd-8a85-3ed37a3fdb4d"), My.Resources.View_Tiles, "lailaShell_View_TilesIcon", GetType(HomeFolderTilesView)),
            New FolderViewRegistration(New Guid("fc73e12c-6b88-4a3f-9a2e-c7414bcd1f1f"), My.Resources.View_Details, "lailaShell_View_DetailsIcon", GetType(HomeFolderDetailsView))
        }
        Me.DefaultView = New Guid("5aab6a71-6d79-4bcd-8a85-3ed37a3fdb4d")

        AddHandler PinnedItems.ItemPinned,
            Sub(s2 As Object, e2 As PinnedItemEventArgs)
                If _isLoaded Then
                    _isEnumerated = False
                    _isEnumeratedForTree = False
                    Dim __ = Me.GetItemsAsync()
                End If
            End Sub
        AddHandler PinnedItems.ItemUnpinned,
            Sub(s2 As Object, e2 As PinnedItemEventArgs)
                If _isLoaded Then
                    _isEnumerated = False
                    _isEnumeratedForTree = False
                    Dim __ = Me.GetItemsAsync()
                End If
            End Sub
    End Sub

    Public Overrides ReadOnly Property Pidl As Pidl
        Get
            If _pidl Is Nothing Then
                _pidl = Shell.GetSpecialFolder(SpecialFolders.Recent).Pidl.Clone()
            End If
            Return _pidl
        End Get
    End Property

    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return My.Resources.Folder_Home_DisplayName
        End Get
    End Property

    Public Overrides ReadOnly Property FullPath As String
        Get
            Return "::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}HOME"
        End Get
    End Property

    Public Overrides ReadOnly Property Attributes As SFGAO
        Get
            Return SFGAO.FOLDER
        End Get
    End Property

    Public Overrides ReadOnly Property Parent As Folder
        Get
            Return Shell.Desktop
        End Get
    End Property

    Friend Overrides Function MakeIShellFolderOnCurrentThread() As IShellFolderForIContextMenu
        Return Nothing
    End Function

    Public Overrides Property HasSubFolders As Boolean
        Get
            If _hasSubFolders.HasValue AndAlso (_isEnumerated OrElse _hasSubFolders.Value) Then
                _cachedHasSubFolders = _hasSubFolders
                Return _hasSubFolders.Value
            ElseIf _cachedHasSubFolders.HasValue Then
                Return _cachedHasSubFolders.Value
            Else
                Return False
            End If
        End Get
        Set(value As Boolean)
            MyBase.HasSubFolders = value
        End Set
    End Property

    Public Overrides ReadOnly Property Icon(size As Integer) As BitmapSource
        Get
            Dim hbitmap As IntPtr, result As BitmapSource = Nothing, shellItem As IShellItem2 = Nothing
            Try
                shellItem = Item.GetIShellItem2FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
                Dim h As HRESULT = CType(shellItem, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY, hbitmap)
                If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                    result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Else
                    Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                End If
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
                If Not shellItem Is Nothing Then
                    Marshal.ReleaseComObject(shellItem)
                    shellItem = Nothing
                End If
            End Try
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property Image(size As Integer) As BitmapSource
        Get
            Dim hbitmap As IntPtr, result As BitmapSource = Nothing, shellItem As IShellItem2 = Nothing
            Try
                shellItem = Item.GetIShellItem2FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
                Dim h As HRESULT = CType(shellItem, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), 0, hbitmap)
                If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                    result = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Else
                    Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                End If
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
                If Not shellItem Is Nothing Then
                    Marshal.ReleaseComObject(shellItem)
                    shellItem = Nothing
                End If
            End Try
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property OverlayImage(size As Integer) As BitmapSource
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Sub AddRightClickMenuItems(menu As RightClickMenu)
        ' don't show menu for home folder, except for children
        If Not menu.SelectedItems Is Nothing AndAlso menu.SelectedItems.Count > 0 _
            AndAlso Not (menu.SelectedItems.Count = 1 AndAlso menu.SelectedItems(0).Equals(Me)) Then
            MyBase.AddRightClickMenuItems(menu)
        End If
    End Sub

    Protected Overrides Function EnumerateItems(shellItem2 As IShellItem2, flags As UInteger, cancellationToken As CancellationToken,
        isSortPropertyByText As Boolean, isSortPropertyDisplaySortValue As Boolean, sortPropertyKey As String,
        result As Dictionary(Of String, Item), newFullPaths As HashSet(Of String), addItems As Action, threadId As Integer?,
        dupes As List(Of Item), Optional maxCount As Integer = -1) As Exception

        _hookFolderFullPath = Shell.GetSpecialFolder(SpecialFolders.Recent).FullPath

        ' enumerate pinned items
        Dim count As UInt64 = UInt64.MaxValue
        For Each item In PinnedItems.GetPinnedItems()
            If Not result.ContainsKey(item.FullPath & "_" & item.DisplayName) Then
                item.LogicalParent = Me
                item.TreeSortPrefix = String.Format("{0:00000000000000000000}", UInt64.MaxValue - count)
                item.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)
                item._hasCustomProperties = True

                Dim systemLastAccessedProperty As [Property] = item.PropertiesByKey(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:16"))
                Dim lastAccessedProperty As Home_LastAccessedProperty = New Home_LastAccessedProperty(systemLastAccessedProperty.Value)
                item._propertiesByKey.Add(Home_LastAccessedProperty.Key.ToString(), lastAccessedProperty)

                Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.QUICK_LAUNCH)
                item._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

                item.DeDupeKey = "PINNED"
                result.Add(item.Pidl.ToString() & item.DeDupeKey, item)
                dupes.Add(item)
                newFullPaths.Add(item.Pidl.ToString() & item.DeDupeKey)
                If TypeOf item Is Folder Then Me.HasSubFolders = True

                count -= 1
            End If
        Next

        ' enumerate frequent folders
        For Each item In FrequentFolders.GetMostFrequent()
            item.LogicalParent = Me
            item.TreeSortPrefix = String.Format("{0:00000000000000000000}", UInt64.MaxValue - count)
            item.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)
            item._hasCustomProperties = True

            Dim systemLastAccessedProperty As [Property] = item.PropertiesByKey(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:16"))
            Dim lastAccessedProperty As Home_LastAccessedProperty = New Home_LastAccessedProperty(systemLastAccessedProperty.Value)
            item._propertiesByKey.Add(Home_LastAccessedProperty.Key.ToString(), lastAccessedProperty)

            Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.QUICK_LAUNCH)
            item._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

            item.DeDupeKey = "FREQUENT"
            result.Add(item.Pidl.ToString() & item.DeDupeKey, item)
            dupes.Add(item)
            newFullPaths.Add(item.Pidl.ToString() & item.DeDupeKey)
            If TypeOf item Is Folder Then Me.HasSubFolders = True

            count -= 1
        Next

        Thread.Sleep(200) ' give the disk a chance to update

        ' enumerate recent files
        Return MyBase.EnumerateItems(Shell.GetSpecialFolder(SpecialFolders.Recent).Clone().ShellItem2, flags, cancellationToken, isSortPropertyByText,
            isSortPropertyDisplaySortValue, sortPropertyKey, result, newFullPaths, addItems, threadId, dupes, result.Count + 16)
    End Function

    Friend Overrides Sub OnItemsChanged(Optional item As Item = Nothing)
        MyBase.OnItemsChanged()
    End Sub

    Protected Friend Overrides Function InitializeItem(item As Item) As Item
        If TypeOf item Is Link Then
            'CType(item, Link).Resolve(SLR_FLAGS.NO_UI Or SLR_FLAGS.NOSEARCH)
            Dim target As Item = CType(item, Link).TargetItem
            If Not target Is Nothing _
                AndAlso (Not TypeOf target Is Folder OrElse Not target.Attributes.HasFlag(SFGAO.STORAGEANCESTOR)) _
                AndAlso target.IsExisting _
                AndAlso Not String.IsNullOrWhiteSpace(target.PropertiesByKey(New PROPERTYKEY("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6")).Text) Then

                Dim modifiedProperty As [Property] = item.PropertiesByCanonicalName("System.DateModified")

                Dim shellItem2 As IShellItem2
                shellItem2 = GetIShellItem2FromPidl(item.Pidl.AbsolutePIDL, Nothing) ' get the IShellItem2
                Dim pidl As Pidl = item.Pidl.Clone() ' clone the pidl to prevent it from being disposed
                Shell.DisposerThreadPool.Add(
                    Sub()
                        item.Dispose()
                    End Sub)
                item = New ProxyLink(shellItem2, Me, True, True, Nothing, pidl)
                target = CType(item, ProxyLink).TargetItem

                item.LogicalParent = Me
                item._hasCustomProperties = True
                item.CanShowInTree = False

                Dim locationProperty As [Property] = target.PropertiesByKey(New PROPERTYKEY("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6"))
                locationProperty.IsCustom = True
                item._propertiesByKey.Add("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6".ToLower(), locationProperty)

                Dim storageProviderProperty As [Property] = target.PropertiesByKey(System_StorageProviderUIStatusProperty.Key)
                storageProviderProperty.IsCustom = True
                item._propertiesByKey.Add(System_StorageProviderUIStatusProperty.Key.ToString(), storageProviderProperty)

                item.ItemNameDisplaySortValuePrefix = String.Format("{0:yyyyMMddHHmmssffff}", modifiedProperty.Value)

                Dim lastAccessedProperty As Home_LastAccessedProperty = New Home_LastAccessedProperty(modifiedProperty.Value)
                item._propertiesByKey.Add(Home_LastAccessedProperty.Key.ToString(), lastAccessedProperty)

                Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.RECENT_FILE)
                item._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

                Return item
            End If
        End If

        item._logicalParent = Nothing

        Return Nothing
    End Function

    Public Overrides ReadOnly Property CanSort As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property CanGroupBy As Boolean
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ISupportDragInsert_Items As ObservableCollection(Of Item) Implements ISupportDragInsert.Items
        Get
            Return Me.Items
        End Get
    End Property

    Public Overrides Function Clone() As Item
        Return Shell.GlobalThreadPool.Run(
            Function() As Item
                Return New HomeFolder(Nothing, _doKeepAlive)
            End Function)
    End Function

    Public Function DragInsertBefore(dataObject As IDataObject_PreserveSig, files As List(Of Item), index As Integer, overListBoxItem As ListBoxItem) As HRESULT Implements ISupportDragInsert.DragInsertBefore
        Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(Me.Items)
        Dim canPinItem As Boolean =
            index = 0 _
            OrElse (index > view.Count - 1 AndAlso view(view.Count - 1).IsPinned) _
            OrElse (view(index - 1).IsPinned _
                AndAlso (UIHelper.GetParentOfType(Of BaseFolderView)(overListBoxItem) Is Nothing _
                         OrElse CType(overListBoxItem?.DataContext, Item).IsPinned))
        If canPinItem Then
            WpfDragTargetProxy.SetDropDescription(dataObject, DROPIMAGETYPE.DROPIMAGE_LINK, My.Resources.Folder_Home_DragDrop_PinTo, My.Resources.Folder_Home_QuickLaunch)
            Return HRESULT.S_OK
        Else
            Return HRESULT.S_FALSE
        End If
    End Function

    Public Function Drop(dataObject As IDataObject_PreserveSig, files As List(Of Item), index As Integer) As HRESULT Implements ISupportDragInsert.Drop
        Try
            PinnedItems.IsNotifying = False

            For Each file In files
                Dim unpinnedIndex As Integer = PinnedItems.UnpinItem(file.Pidl)
                If unpinnedIndex <> -1 AndAlso unpinnedIndex < index Then
                    If index <> 0 Then index -= 1
                End If
            Next
            For Each file In files
                PinnedItems.PinItem(file, index)
                If index <> -1 Then index += 1
            Next

            _isEnumerated = False
            _isEnumeratedForTree = False
            Dim __ = Me.GetItemsAsync()

            Return HRESULT.S_OK
        Finally
            PinnedItems.IsNotifying = True
            PinnedItems.NotifyReset()
        End Try
    End Function

    Public Overrides Sub Refresh(Optional newShellItem As IShellItem2 = Nothing,
                                 Optional newPidl As Pidl = Nothing,
                                 Optional doRefreshImage As Boolean = True,
                                 Optional threadId As Integer? = Nothing,
                                 Optional count As Integer = 1)

    End Sub
End Class
