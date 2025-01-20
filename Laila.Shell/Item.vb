Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class Item
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Protected Const MAX_PATH_LENGTH As Integer = 260

    Protected _propertiesByKey As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Protected _propertiesByCanonicalName As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Friend _fullPath As String
    Friend disposedValue As Boolean
    Friend _parent As Folder
    Friend _displayName As String
    Private _isPinned As Boolean
    Private _isCut As Boolean
    Private _attributes As SFGAO
    Private _treeRootIndex As Long = -1
    Friend _shellItem2 As IShellItem2
    Friend _objectId As Long = -1
    Private Shared _objectCount As Long = 0
    Private _pidl As Pidl
    Private _isImage As Boolean?
    Private _propertiesLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Friend _shellItemLock As Object = New Object()
    Protected _doKeepAlive As Boolean
    Private _contentViewModeProperties() As [Property]
    Private _isVisibleInAddressBar As Boolean
    Private _treeSortPrefix As String = String.Empty

    Public Shared Function FromParsingName(parsingName As String, parent As Folder,
                                           Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True) As Item
        parsingName = Environment.ExpandEnvironmentVariables(parsingName)
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        If Not shellItem2 Is Nothing Then
            Dim attr As SFGAO = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If attr.HasFlag(SFGAO.FOLDER) Then
                Return New Folder(shellItem2, parent, doKeepAlive, doHookUpdates)
            Else
                Return New Item(shellItem2, parent, doKeepAlive, doHookUpdates)
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function FromPidl(pidl As IntPtr, parent As Folder,
                                    Optional doKeepAlive As Boolean = False, Optional doHookUpdates As Boolean = True) As Item
        Dim shellItem2 As IShellItem2, shellFolder As IShellFolder
        shellItem2 = GetIShellItem2FromPidl(pidl, parent?.ShellFolder)
        If Not shellItem2 Is Nothing Then
            Dim attr As SFGAO = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If attr.HasFlag(SFGAO.FOLDER) Then
                Return New Folder(shellItem2, parent, doKeepAlive, doHookUpdates, pidl)
            Else
                Return New Item(shellItem2, parent, doKeepAlive, doHookUpdates, pidl)
            End If
        Else
            Return Nothing
        End If
    End Function

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr, parentShellFolder As IShellFolder) As IShellItem2
        Dim result As IShellItem2
        If parentShellFolder Is Nothing Then
            Functions.SHCreateItemFromIDList(pidl, Guids.IID_IShellItem2, result)
        Else
            Functions.SHCreateItemWithParent(IntPtr.Zero, parentShellFolder, pidl, Guids.IID_IShellItem2, result)
        End If
        Return result
    End Function

    Friend Shared Function GetIShellItem2FromParsingName(parsingName As String) As IShellItem2
        Dim result As IShellItem2
        Functions.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, Guids.IID_IShellItem2, result)
        Return result
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, Optional pidl As IntPtr? = Nothing)
        _objectCount += 1
        _objectId = _objectCount
        _shellItem2 = shellItem2
        _doKeepAlive = doKeepAlive
        If pidl.HasValue AndAlso Not IntPtr.Zero.Equals(pidl.Value) Then
            _pidl = New Pidl(pidl.Value).Clone()
        End If
        If Not shellItem2 Is Nothing Then
            If doHookUpdates Then Me.HookUpdates()
            Shell.AddToItemsCache(Me)
        Else
            _fullPath = String.Empty
        End If
        If Not logicalParent Is Nothing Then
            _parent = logicalParent
        Else
            Me.IsLogicalOrphan = True
        End If
        AddHandler Shell.Settings.PropertyChanged, AddressOf Settings_PropertyChanged
    End Sub

    Friend Property IsLogicalOrphan As Boolean

    Public Sub HookUpdates()
        AddHandler Shell.Notification, AddressOf shell_Notification
    End Sub

    Public ReadOnly Property ShellItem2 As IShellItem2
        Get
            Return _shellItem2
        End Get
    End Property

    Public ReadOnly Property Pidl As Pidl
        Get
            SyncLock _shellItemLock
                If _pidl Is Nothing AndAlso Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
                    Dim pidlptr As IntPtr
                    Functions.SHGetIDListFromObject(_shellItem2, pidlptr)
                    _pidl = New Pidl(pidlptr)
                End If
            End SyncLock
            Return _pidl
        End Get
    End Property

    Public Property IsVisibleInAddressBar As Boolean
        Get
            Return _isVisibleInAddressBar
        End Get
        Set(value As Boolean)
            SetValue(_isVisibleInAddressBar, value)
        End Set
    End Property

    Public ReadOnly Property IsVisibleInTree As Boolean
        Get
            Return Me.TreeRootIndex <> -1 OrElse (Not _parent Is Nothing AndAlso _parent.IsExpanded)
        End Get
    End Property

    Public Overridable ReadOnly Property IsReadyForDispose As Boolean
        Get
            Return Not _doKeepAlive _
                AndAlso (_parent Is Nothing OrElse (Not _parent.IsActiveInFolderView)) _
                AndAlso Not Me.IsVisibleInTree AndAlso Not Me.IsVisibleInAddressBar
        End Get
    End Property

    Public Overridable Sub MaybeDispose()
        If Me.IsReadyForDispose Then
            Me.Dispose()
        End If
    End Sub

    Public Property TreeRootIndex As Long
        Get
            Return _treeRootIndex
        End Get
        Set(value As Long)
            SetValue(_treeRootIndex, value)
            Me.NotifyOfPropertyChange("TreeSortKey")
        End Set
    End Property

    Friend Property TreeSortPrefix As String
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
                Return Me.Parent?.TreeSortKey & Me.TreeSortPrefix & itemNameDisplaySortValue & New String(" ", 260 - itemNameDisplaySortValue.Length)
            End If
        End Get
    End Property

    Public ReadOnly Property TreeMargin As Thickness
        Get
            Dim level As Integer = 0
            If Me.TreeRootIndex = -1 Then
                Dim lp As Folder = Me.Parent
                While Not lp Is Nothing
                    level += 1
                    If lp.TreeRootIndex <> -1 Then Exit While
                    lp = lp.Parent
                End While
            End If

            Return New Thickness(level * 16, 0, 0, 0)
        End Get
    End Property

    Protected Overridable Function GetNewShellItem() As IShellItem2
        Dim result As IShellItem2
        Functions.SHCreateItemFromIDList(Me.Pidl.AbsolutePIDL, GetType(IShellItem2).GUID, result)
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

    Public Overridable Sub Refresh(Optional newShellItem As IShellItem2 = Nothing)
        SyncLock _shellItemLock
            If Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
                Dim oldItemNameDisplaySortValue As String = Me.ItemNameDisplaySortValue

                Dim oldShellItem As IShellItem2 = _shellItem2
                If Not newShellItem Is Nothing Then
                    _shellItem2 = newShellItem
                ElseIf Not _shellItem2 Is Nothing Then
                    _shellItem2 = Me.GetNewShellItem()
                End If
                Marshal.ReleaseComObject(oldShellItem)
                If Not _pidl Is Nothing Then
                    _pidl.Dispose()
                    _pidl = Nothing
                End If

                Dim oldPropertiesByKey As Dictionary(Of String, [Property])
                Dim oldPropertiesByCanonicalName As Dictionary(Of String, [Property])
                oldPropertiesByKey = _propertiesByKey
                oldPropertiesByCanonicalName = _propertiesByCanonicalName
                _propertiesByKey = New Dictionary(Of String, [Property])()
                _propertiesByCanonicalName = New Dictionary(Of String, [Property])()
                For Each [property] In oldPropertiesByKey.Values
                    [property].Dispose()
                Next
                For Each [property] In oldPropertiesByCanonicalName.Values
                    [property].Dispose()
                Next

                _displayName = Nothing

                If Not Me.ShellItem2 Is Nothing Then
                    _fullPath = Nothing
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

                    Me.NotifyOfPropertyChange("DisplayName")
                    If Me.ItemNameDisplaySortValue <> oldItemNameDisplaySortValue Then
                        Me.NotifyOfPropertyChange("ItemNameDisplaySortValue")
                    End If
                    Me.NotifyOfPropertyChange("OverlayImageAsync")
                    Me.NotifyOfPropertyChange("IconAsync")
                    Me.NotifyOfPropertyChange("ImageAsync")
                    Me.NotifyOfPropertyChange("HasThumbnailAsync")
                    Me.NotifyOfPropertyChange("PropertiesByKeyAsText")
                    Me.NotifyOfPropertyChange("IsImage")
                    Me.NotifyOfPropertyChange("IsHidden")
                    Me.NotifyOfPropertyChange("IsCompressed")
                    Me.NotifyOfPropertyChange("IsEncrypted")
                    Me.NotifyOfPropertyChange("StorageProviderUIStatusFirstIcon16Async")
                    Me.NotifyOfPropertyChange("StorageProviderUIStatusHasIcon")
                    Me.NotifyOfPropertyChange("StorageProviderUIStatusIcons16Async")
                    Me.NotifyOfPropertyChange("StorageProviderUIStatusIconWidth12")
                    Me.NotifyOfPropertyChange("StorageProviderUIStatusIconWidth16")
                    For Each prop In oldPropertiesByKey
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}]", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Text", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].FirstIcon16Async", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].HasIconAsync", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Icons16Async", prop.Key))
                    Next
                    For Each prop In oldPropertiesByCanonicalName
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}]", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Text", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].FirstIcon16Async", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].HasIconAsync", prop.Key))
                        Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Icons16Async", prop.Key))
                    Next
                Else
                    Me.Dispose()
                End If
            Else
                Debug.WriteLine(Me.FullPath & "  " & disposedValue)
            End If
        End SyncLock
    End Sub

    Public ReadOnly Property FullPath As String
        Get
            SyncLock _shellItemLock
                If String.IsNullOrWhiteSpace(_fullPath) AndAlso Not disposedValue Then
                    Me.ShellItem2.GetDisplayName(SIGDN.DESKTOPABSOLUTEPARSING, _fullPath)
                End If
            End SyncLock

            Return _fullPath
        End Get
    End Property

    Public ReadOnly Property Parent As Folder
        Get
            If Not disposedValue _
                AndAlso (_parent Is Nothing OrElse _parent.disposedValue) _
                AndAlso Not Me.FullPath?.Equals(Shell.Desktop.FullPath) Then
                Dim parentShellItem2 As IShellItem2
                _parent = Shell.RunOnSTAThread(
                    Sub(tcs As TaskCompletionSource(Of Folder))
                        SyncLock _shellItemLock
                            If Not Me.ShellItem2 Is Nothing Then
                                Me.ShellItem2.GetParent(parentShellItem2)
                            End If
                        End SyncLock
                        If Not parentShellItem2 Is Nothing Then
                            _parent = New Folder(parentShellItem2, Nothing, False, True)
                        End If
                        tcs.SetResult(_parent)
                    End Sub, 1)
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

    Public ReadOnly Property AssociatedApplicationIcon(size As Integer) As ImageSource
        Get
            If Not disposedValue AndAlso (IO.File.Exists(Me.FullPath) OrElse IO.Directory.Exists(Me.FullPath)) Then
                Dim app As String = Me.GetAssociatedApplication()
                If Not String.IsNullOrWhiteSpace(app) Then
                    Dim shFileInfo As New SHFILEINFO()
                    Try
                        Dim result As IntPtr = Functions.SHGetFileInfo(
                            app,
                            0,
                            shFileInfo,
                            Marshal.SizeOf(shFileInfo),
                            SHGFI.SHGFI_ICON Or If(size > 16, SHGFI.SHGFI_LARGEICON, SHGFI.SHGFI_SMALLICON)
                        )
                        If Not IntPtr.Zero.Equals(result) Then
                            Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(shFileInfo.hIcon)
                                Using bitmap = icon.ToBitmap()
                                    Dim hBitmap As IntPtr = bitmap.GetHbitmap()
                                    Dim image As BitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                    image.Freeze()
                                    Return image
                                End Using
                            End Using
                        End If
                    Finally
                        If Not IntPtr.Zero.Equals(shFileInfo.hIcon) Then
                            Functions.DestroyIcon(shFileInfo.hIcon)
                        End If
                    End Try
                End If
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property AssociatedApplicationIconAsync(size As Integer) As ImageSource
        Get
            Return Shell.RunOnSTAThread(
                Sub(tcs As TaskCompletionSource(Of ImageSource))
                    Dim result As ImageSource
                    result = Me.AssociatedApplicationIcon(size)
                    If Not result Is Nothing Then result.Freeze()
                    tcs.SetResult(result)
                End Sub, 3)
        End Get
    End Property

    Public ReadOnly Property OverlayIconIndex As Byte
        Get
            If Not disposedValue AndAlso (IO.File.Exists(Me.FullPath) OrElse IO.Directory.Exists(Me.FullPath)) Then
                Dim shFileInfo As New SHFILEINFO()
                Try
                    Dim result As IntPtr = Functions.SHGetFileInfo(
                        Me.FullPath,
                        0,
                        shFileInfo,
                        Marshal.SizeOf(shFileInfo),
                        SHGFI.SHGFI_ICON Or SHGFI.SHGFI_OVERLAYINDEX
                    )
                    If Not IntPtr.Zero.Equals(result) Then
                        Return CByte((shFileInfo.iIcon >> 24) And &HFF)
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(shFileInfo.hIcon) Then
                        Functions.DestroyIcon(shFileInfo.hIcon)
                    End If
                End Try
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
            Dim overlayIconIndex As Byte? = Shell.RunOnSTAThread(
                Sub(tcs As TaskCompletionSource(Of Byte?))
                    tcs.SetResult(Me.OverlayIconIndex)
                End Sub, 3)

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
                        Dim h As HRESULT, result As ImageSource
                        h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY Or SIIGBF.SIIGBF_INCACHEONLY, hbitmap)
                        If h = HRESULT.S_OK AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                            result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        End If
                        If h <> HRESULT.S_OK OrElse IntPtr.Zero.Equals(hbitmap) _
                        OrElse (Not result Is Nothing AndAlso result.Width < size * Settings.DpiScaleX AndAlso result.Height < size * Settings.DpiScaleY) Then
                            If Not IntPtr.Zero.Equals(hbitmap) Then Functions.DeleteObject(hbitmap)
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY, hbitmap)
                            If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            Else
                                Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                            End If
                        End If
                        Return result
                    End If
                End SyncLock
            Finally
                Functions.DeleteObject(hbitmap)
            End Try
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property IconAsync(size As Integer) As ImageSource
        Get
            Return Shell.RunOnSTAThread(
                Sub(tcs As TaskCompletionSource(Of ImageSource))
                    Dim result As ImageSource
                    result = Me.Icon(size)
                    If Not result Is Nothing Then result.Freeze()
                    tcs.SetResult(result)
                End Sub, 3)
        End Get
    End Property

    Public Overridable ReadOnly Property Image(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr
            Try
                If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                    SyncLock _shellItemLock
                        Dim h As HRESULT, result As ImageSource
                        h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_INCACHEONLY, hbitmap)
                        If h = HRESULT.S_OK AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                            result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        End If
                        If h <> HRESULT.S_OK OrElse IntPtr.Zero.Equals(hbitmap) _
                        OrElse (Not result Is Nothing AndAlso result.Width < size * Settings.DpiScaleX AndAlso result.Height < size * Settings.DpiScaleY) Then
                            If Not IntPtr.Zero.Equals(hbitmap) Then Functions.DeleteObject(hbitmap)
                            h = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), 0, hbitmap)
                            If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                                result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            Else
                                Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                            End If
                        End If
                        Return result
                    End SyncLock
                End If
            Finally
                Functions.DeleteObject(hbitmap)
            End Try
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property ImageAsync(size As Integer) As ImageSource
        Get
            Return Shell.RunOnSTAThread(
                Sub(tcs As TaskCompletionSource(Of ImageSource))
                    Dim result As ImageSource
                    result = Me.Image(size)
                    If Not result Is Nothing Then result.Freeze()
                    tcs.SetResult(result)
                End Sub, 3)
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusIcons16Async As ImageSource()
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource())

            Shell.STATaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource()
                        If Not Me.disposedValue Then
                            result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.Icons16
                        End If
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            tcs.Task.Wait(Shell.ShuttingDownToken)
            If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                Return tcs.Task.Result
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusFirstIcon16Async As ImageSource
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource)

            Shell.STATaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource
                        If Not Me.disposedValue Then
                            result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2")?.FirstIcon16
                        End If
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            tcs.Task.Wait(Shell.ShuttingDownToken)
            If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                Return tcs.Task.Result
            Else
                Return Nothing
            End If
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
                Functions.DeleteObject(hbitmap)
            End Try
            Return False
        End Get
    End Property

    Public Overridable ReadOnly Property HasThumbnailAsync As Boolean
        Get
            Dim result As Boolean? = Shell.RunOnSTAThread(
                Sub(tcs As TaskCompletionSource(Of Boolean?))
                    tcs.SetResult(Me.HasThumbnail)
                End Sub, 3)

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
                _isImage = IO.File.Exists(Me.FullPath) AndAlso ImageHelper.IsImage(Me.FullPath)
            End If
            Return _isImage.Value
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayName As String
        Get
            SyncLock _shellItemLock
                If String.IsNullOrWhiteSpace(_displayName) AndAlso Not disposedValue Then
                    Me.ShellItem2.GetDisplayName(SHGDN.NORMAL, _displayName)
                End If
            End SyncLock

            Return _displayName
        End Get
    End Property

    Public Property AddressBarRoot As String

    Public Property AddressBarDisplayName As String

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
                If Not Shell.GetSpecialFolders().Values.ToList().Exists(Function(f) f.Pidl.Equals(parent.Pidl)) Then
                    parent = parent.Parent
                    While Not parent Is Nothing
                        If parent.IsDrive Then
                            path = IO.Path.Combine(If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                                          parent.FullPath, parent.AddressBarDisplayName), path)
                        Else
                            path = IO.Path.Combine(If(String.IsNullOrWhiteSpace(parent.AddressBarDisplayName),
                                                          parent.DisplayName, parent.AddressBarDisplayName), path)
                        End If
                        If Shell.GetSpecialFolders().Values.ToList().Exists(Function(f) f.Pidl.Equals(parent.Pidl)) Then
                            Exit While
                        End If
                        parent = parent.Parent
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

    Public Overridable ReadOnly Property ItemNameDisplaySortValue As String
        Get
            If Me.IsDrive AndAlso Shell.Settings.DoShowDriveLetters Then
                Return Me.FullPath
            Else
                Return If(Me.IsFolder AndAlso Me.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), "0", "1") & Me.DisplayName
            End If
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

    Public ReadOnly Property Attributes As SFGAO
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

    Public Overridable ReadOnly Property HasSubFolders As Boolean
        Get
            Return False
        End Get
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
                Dim system_InfoTipText As String = Me.PropertiesByKey(PKEY_System_InfoTipText).Text
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
                    = Me.PropertiesByKey(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey)
                If System_StorageProviderUIStatus.RawValue.vt <> 0 Then
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
            Dim [property] As [Property]
            Dim key As PROPERTYKEY = New PROPERTYKEY(propertyKey)
            _propertiesLock.Wait()
            Try
                If Not _propertiesByKey.TryGetValue(propertyKey, [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromKey(key, Me.ShellItem2)
                        End If
                    End SyncLock
                    _propertiesLock.Wait()
                    If Not [property] Is Nothing Then
                        If Not _propertiesByKey.ContainsKey(propertyKey) Then
                            _propertiesByKey.Add(propertyKey, [property])
                        Else
                            [property].Dispose()
                            [property] = _propertiesByKey(propertyKey)
                        End If
                    End If
                End If
            Finally
                _propertiesLock.Release()
            End Try
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKey(propertyKey As PROPERTYKEY) As [Property]
        Get
            Dim [property] As [Property]
            _propertiesLock.Wait()
            Try
                If Not _propertiesByKey.TryGetValue(propertyKey.ToString(), [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromKey(propertyKey, Me.ShellItem2)
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
                _propertiesLock.Release()
            End Try
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByCanonicalName(canonicalName As String) As [Property]
        Get
            Dim [property] As [Property]
            _propertiesLock.Wait()
            Try
                If Not _propertiesByCanonicalName.TryGetValue(canonicalName, [property]) AndAlso Not disposedValue Then
                    _propertiesLock.Release()
                    SyncLock _shellItemLock
                        If Not disposedValue AndAlso Not Me.ShellItem2 Is Nothing Then
                            [property] = [Property].FromCanonicalName(canonicalName, Me.ShellItem2)
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
                _propertiesLock.Release()
            End Try
            Return [property]
        End Get
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
                folder = Shell.GetSpecialFolder("Network")
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.GetSpecialFolder("This pc")
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
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) IO.Path.TrimEndingDirectorySeparator(f.FullPath).ToLower() = parts(j).ToLower())
                    Else
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) IO.Path.GetFileName(IO.Path.TrimEndingDirectorySeparator(f.FullPath)).ToLower() = parts(j).ToLower())
                    End If
                    If subFolder Is Nothing Then
                        subFolder = (Await folder.GetItemsAsync()).FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(j).ToLower())
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
                folder = Shell.GetSpecialFolder("Network")
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.GetSpecialFolder("This pc")
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

    Protected Overridable Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        If Not disposedValue Then
            Select Case e.Event
                Case SHCNE.UPDATEITEM, SHCNE.UPDATEDIR
                    If Me.Pidl?.Equals(e.Item1.Pidl) Then
                        Me.Refresh()
                    End If
                Case SHCNE.FREESPACE
                    If Me.IsDrive AndAlso Me.Pidl?.Equals(e.Item1.Pidl) Then
                        Me.Refresh()
                    End If
                Case SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED
                    If Me.IsDrive AndAlso Me.Pidl?.Equals(e.Item1?.Pidl) Then
                        Me.Refresh()
                    End If
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    If Me.Pidl?.Equals(e.Item1?.Pidl) Then
                        Dim oldPidl As Pidl = Me.Pidl
                        _pidl = e.Item2.Pidl.Clone()
                        PinnedItems.RenameItem(oldPidl, _pidl)
                        FrequentFolders.RenameItem(oldPidl, _pidl)
                        oldPidl.Dispose()
                        Me.Refresh(e.Item2.ShellItem2)
                        e.Item2._shellItem2 = Nothing
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
                    Shell.STATaskQueue.Add(
                        Sub()
                            Me.Refresh()
                        End Sub)
                End If
        End Select
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        SyncLock _shellItemLock
            If Not disposedValue Then
                disposedValue = True
                'Debug.WriteLine("Disposing " & _objectId & ": " & Me.FullPath)

                If disposing Then
                    ' dispose managed state (managed objects):

                    ' unsubscribe from notifications
                    RemoveHandler Shell.Notification, AddressOf shell_Notification
                    RemoveHandler Shell.Settings.PropertyChanged, AddressOf Settings_PropertyChanged

                    ' remove from parent collection
                    UIHelper.OnUIThreadAsync(
                        Sub()
                            If Not _parent Is Nothing Then
                                _parent._items.Remove(Me)
                                _parent._isEnumerated = False
                            End If
                        End Sub)

                    ' dispose properties
                    For Each [property] In _propertiesByKey
                        [property].Value.Dispose()
                    Next
                    _propertiesByKey.Clear()
                    For Each [property] In _propertiesByCanonicalName
                        [property].Value.Dispose()
                    Next
                    _propertiesByCanonicalName.Clear()

                    ' dispose shellitem
                    If Not _shellItem2 Is Nothing Then
                        Marshal.ReleaseComObject(_shellItem2)
                        _shellItem2 = Nothing
                    End If

                    ' remove from cache
                    Shell.RemoveFromItemsCache(Me)

                    ' dispose pidl
                    If Not _pidl Is Nothing Then
                        _pidl.Dispose()
                    End If
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
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

    Public Function Clone() As Item
        Return Item.FromPidl(Me.Pidl.AbsolutePIDL, Nothing, _doKeepAlive)
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
