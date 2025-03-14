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

Public Class HomeFolder
    Inherits Folder
    Implements ISupportDragInsert

    Public Sub New(parent As Folder, doKeepAlive As Boolean)
        MyBase.New(Nothing, parent, doKeepAlive, False, Nothing)

        _pidl = Pidl.FromBytes(System.Text.UTF8Encoding.Unicode.GetBytes(Me.FullPath))
        _hasSubFolders = True
        _columns = New List(Of Column)() From {
            New Column(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:10"), New CM_COLUMNINFO(), 0) With {.IsVisible = True},
            New Column(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:16"), New CM_COLUMNINFO(), 0) With {.IsVisible = True},
            New Column(New PROPERTYKEY("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6"), New CM_COLUMNINFO(), 0) With {.IsVisible = True}
        }
        Me.ItemsGroupByPropertyName = "PropertiesByKeyAsText[" & Home_CategoryProperty.Key.ToString() & "].GroupByText"
        Me.ItemsSortDirection = ComponentModel.ListSortDirection.Descending

        AddHandler PinnedItems.ItemPinned,
            Sub(s2 As Object, e2 As PinnedItemEventArgs)
                If _isLoaded Then
                    _isEnumerated = False
                    Me.GetItemsAsync()
                End If
            End Sub
        AddHandler PinnedItems.ItemUnpinned,
            Sub(s2 As Object, e2 As PinnedItemEventArgs)
                If _isLoaded Then
                    _isEnumerated = False
                    Me.GetItemsAsync()
                End If
            End Sub
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return "Home"
        End Get
    End Property

    Public Overrides ReadOnly Property FullPath As String
        Get
            Return "::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}"
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

    Friend Overrides Function GetShellFolderOnCurrentThread() As IShellFolderForIContextMenu
        Return Nothing
    End Function

    Public Overrides ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr, result As ImageSource = Nothing, shellItem As IShellItem2 = Nothing
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

    Public Overrides ReadOnly Property Image(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr, result As ImageSource = Nothing, shellItem As IShellItem2 = Nothing
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

    Public Overrides ReadOnly Property OverlayImage(size As Integer) As ImageSource
        Get
            Return Nothing
        End Get
    End Property

    Protected Overrides Sub EnumerateItems(shellItem2 As IShellItem2, flags As UInteger, cancellationToken As CancellationToken,
        isSortPropertyByText As Boolean, isSortPropertyDisplaySortValue As Boolean, sortPropertyKey As String,
        result As Dictionary(Of String, Item), newFullPaths As HashSet(Of String), addItems As Action, threadId As Integer?)

        ' enumerate pinned items
        Dim count As UInt64 = UInt64.MaxValue
        For Each item In PinnedItems.GetPinnedItems()
            If Not result.ContainsKey(item.FullPath & "_" & item.DisplayName) Then
                item.LogicalParent = Me
                item.TreeSortPrefix = String.Format("{0:00000000000000000000}", UInt64.MaxValue - count)
                item.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)

                Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.PINNED_ITEM)
                item._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

                item.DeDupeKey = "PINNED"
                result.Add(item.FullPath & item.DeDupeKey, item)
                newFullPaths.Add(item.FullPath & item.DeDupeKey)
                If TypeOf item Is Folder Then Me.HasSubFolders = True

                count -= 1
            End If
        Next

        ' enumerate frequent folders
        For Each item In FrequentFolders.GetMostFrequent()
            item.LogicalParent = Me
            item.TreeSortPrefix = String.Format("{0:00000000000000000000}", UInt64.MaxValue - count)
            item.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)

            Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.FREQUENT_FOLDER)
            item._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

            item.DeDupeKey = "FREQUENT"
            result.Add(item.FullPath & item.DeDupeKey, item)
            newFullPaths.Add(item.FullPath & item.DeDupeKey)
            If TypeOf item Is Folder Then Me.HasSubFolders = True

            count -= 1
        Next

        ' enumerate recent files
        MyBase.EnumerateItems(Shell.GetSpecialFolder(SpecialFolders.Recent).Clone().ShellItem2, flags, cancellationToken, isSortPropertyByText,
            isSortPropertyDisplaySortValue, sortPropertyKey, result, newFullPaths, addItems, threadId)
    End Sub

    Protected Overrides Function InitializeItem(item As Item) As Item
        Try
            If TypeOf item Is Link Then
                CType(item, Link).Resolve(SLR_FLAGS.NO_UI Or SLR_FLAGS.NOSEARCH)
                Dim target As Item = CType(item, Link).TargetItem
                If Not TypeOf target Is Folder AndAlso target.IsExisting _
                    AndAlso Not String.IsNullOrWhiteSpace(target.PropertiesByKeyAsText("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6").Text) Then
                    target.LogicalParent = Me

                    Dim modifiedProperty As [Property] = item.PropertiesByKeyAsText("b725f130-47ef-101a-a5f1-02608c9eebac:14")
                    target.ItemNameDisplaySortValuePrefix = String.Format("{0:yyyyMMddHHmmssffff}", modifiedProperty.Value)

                    Dim lastAccessedProperty As [Property] = target.PropertiesByKeyAsText("B725F130-47EF-101A-A5F1-02608C9EEBAC:16")
                    lastAccessedProperty._rawValue.Dispose()
                    lastAccessedProperty._rawValue = New PROPVARIANT()
                    lastAccessedProperty._rawValue.SetValue(CType(modifiedProperty.Value, DateTime))

                    Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.RECENT_FILE)
                    target._propertiesByKey.Add(Home_CategoryProperty.Key.ToString(), categoryProperty)

                    Return target
                Else
                    target.Dispose()
                End If
            End If
        Finally
            item.Dispose()
        End Try

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
            WpfDragTargetProxy.SetDropDescription(dataObject, DROPIMAGETYPE.DROPIMAGE_LINK, "Pin to %1", "Quick access")
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
            Me.GetItemsAsync()

            Return HRESULT.S_OK
        Finally
            PinnedItems.IsNotifying = True
            PinnedItems.NotifyReset()
        End Try
    End Function
End Class
