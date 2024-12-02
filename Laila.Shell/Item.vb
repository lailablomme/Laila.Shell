Imports System.Drawing
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

    Protected _properties As HashSet(Of [Property]) = New HashSet(Of [Property])
    Protected _fullPath As String
    Friend disposedValue As Boolean
    Friend _logicalParent As Folder
    Protected _displayName As String
    Private _isPinned As Boolean
    Private _isCut As Boolean
    Private _attributes As SFGAO
    Private _treeRootIndex As Long = -1
    Private _shellItem2 As IShellItem2
    Private _objectId As Long = -1
    Private Shared _objectCount As Long = 0
    Private _expiredShellItem2 As List(Of IShellItem2) = New List(Of IShellItem2)
    Private _pidl As Pidl
    Private _isImage As Boolean?
    Private _propertiesLock As SemaphoreSlim = New SemaphoreSlim(1, 1)

    Public Shared Function FromParsingName(parsingName As String, parent As Folder) As Item
        parsingName = Environment.ExpandEnvironmentVariables(parsingName)
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        If Not shellItem2 Is Nothing Then
            Dim attr As SFGAO = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If attr.HasFlag(SFGAO.FOLDER) Then
                Return New Folder(shellItem2, parent)
            Else
                Return New Item(shellItem2, parent)
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function FromPidl(pidl As IntPtr, parent As Folder) As Item
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromPidl(pidl)
        If Not shellItem2 Is Nothing Then
            Dim attr As SFGAO = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If attr.HasFlag(SFGAO.FOLDER) Then
                Return New Folder(shellItem2, parent)
            Else
                Return New Item(shellItem2, parent)
            End If
        Else
            Return Nothing
        End If
    End Function

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr) As IShellItem2
        Dim ptr As IntPtr
        Try
            Functions.SHCreateItemFromIDList(pidl, Guids.IID_IShellItem2, ptr)
            Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellItem2))
        Finally
            If Not IntPtr.Zero.Equals(ptr) Then
                Marshal.Release(ptr)
            End If
        End Try
    End Function

    Friend Shared Function GetIShellItem2FromParsingName(parsingName As String) As IShellItem2
        Dim ptr As IntPtr
        Try
            Functions.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, Guids.IID_IShellItem2, ptr)
            Return If(Not IntPtr.Zero.Equals(ptr), Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellItem2)), Nothing)
        Finally
            If Not IntPtr.Zero.Equals(ptr) Then
                Marshal.Release(ptr)
            End If
        End Try
    End Function

    Friend Shared Function GetFullPathFromShellItem2(shellItem2 As IShellItem2) As String
        Dim fullPath As String
        shellItem2.GetDisplayName(SHGDN.FORPARSING, fullPath)
        If Not fullPath Is Nothing AndAlso fullPath.StartsWith("::{") AndAlso fullPath.EndsWith("}") Then
            fullPath = "shell:" & fullPath
        End If
        Return fullPath
    End Function

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder)
        _objectCount += 1
        _objectId = _objectCount
        _shellItem2 = shellItem2
        If Not shellItem2 Is Nothing Then
            'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting PIDL", DateTime.Now)
            Dim ptr As IntPtr, pidl As IntPtr
            Try
                ptr = Marshal.GetIUnknownForObject(_shellItem2)
                Functions.SHGetIDListFromObject(ptr, pidl)
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
            End Try

            _pidl = New Pidl(pidl)
            'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting full path", DateTime.Now)
            _fullPath = Item.GetFullPathFromShellItem2(shellItem2)
            'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting display name", DateTime.Now)
            Dim dn As String = Me.DisplayName
            'Debug.WriteLine("{0:HH:mm:ss.ffff} Getting attributes", DateTime.Now)
            _attributes = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANLINK Or SFGAO.CANRENAME _
            Or SFGAO.CANDELETE Or SFGAO.DROPTARGET Or SFGAO.ENCRYPTED Or SFGAO.ISSLOW _
            Or SFGAO.LINK Or SFGAO.SHARE Or SFGAO.RDONLY Or SFGAO.HIDDEN Or SFGAO.FOLDER _
            Or SFGAO.FILESYSTEM Or SFGAO.COMPRESSED
            shellItem2.GetAttributes(_attributes, _attributes)
            AddHandler Shell.Notification, AddressOf shell_Notification
            'Debug.WriteLine("{0:HH:mm:ss.ffff} Adding to cache", DateTime.Now)
            Shell.AddToItemsCache(Me)
        Else
            _fullPath = String.Empty
        End If
        If Not logicalParent Is Nothing Then
            _logicalParent = logicalParent
        End If
    End Sub

    Public ReadOnly Property ShellItem2 As IShellItem2
        Get
            If Not disposedValue AndAlso _shellItem2 Is Nothing AndAlso Not Me.Pidl Is Nothing Then
                Dim ptr As IntPtr
                Try
                    Functions.SHCreateItemFromParsingName(Me.FullPath, IntPtr.Zero, GetType(IShellItem2).GUID, ptr)
                    If IntPtr.Zero.Equals(ptr) Then
                        Functions.SHCreateItemFromIDList(Me.Pidl.AbsolutePIDL, GetType(IShellItem2).GUID, ptr)
                    End If
                    If Not IntPtr.Zero.Equals(ptr) Then
                        _shellItem2 = Marshal.GetObjectForIUnknown(ptr)
                        _shellItem2.Update(IntPtr.Zero)
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(ptr) Then
                        Marshal.Release(ptr)
                    End If
                End Try
            End If
            Return _shellItem2
        End Get
    End Property

    Public ReadOnly Property Pidl As Pidl
        Get
            Return _pidl
        End Get
    End Property

    Public ReadOnly Property IsVisibleInTree As Boolean
        Get
            Return Me.TreeRootIndex <> -1 OrElse (Not _logicalParent Is Nothing AndAlso _logicalParent.IsExpanded)
        End Get
    End Property

    Public Overridable Sub MaybeDispose()
        If Not _logicalParent.IsActiveInFolderView AndAlso Not Me.IsVisibleInTree Then
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
            Me.MaybeDispose()
        End Set
    End Property

    Public ReadOnly Property TreeSortKey As String
        Get
            If _treeRootIndex <> -1 Then
                Return String.Format("{0:0000000000000000000}", _treeRootIndex)
            Else
                Dim itemNameDisplaySortValue As String = Me.ItemNameDisplaySortValue
                If itemNameDisplaySortValue Is Nothing Then itemNameDisplaySortValue = ""
                Return _logicalParent.TreeSortKey & itemNameDisplaySortValue & New String(" ", 260 - itemNameDisplaySortValue.Length)
            End If
        End Get
    End Property

    Public ReadOnly Property TreeMargin As Thickness
        Get
            Dim level As Integer = 0
            If Me.TreeRootIndex = -1 Then
                Dim lp As Folder = _logicalParent
                While Not lp Is Nothing
                    level += 1
                    If lp.TreeRootIndex <> -1 Then Exit While
                    lp = lp._logicalParent
                End While
            End If

            Return New Thickness(level * 16, 0, 0, 0)
        End Get
    End Property

    Public Overridable Sub ClearCache()
        If Not _shellItem2 Is Nothing Then
            _expiredShellItem2.Add(_shellItem2)
            _shellItem2 = Nothing
        End If

        For Each [property] In _properties
            [property].Dispose()
        Next
        _properties = New HashSet(Of [Property])()
        _displayName = Nothing
    End Sub

    Public Overridable Sub Refresh()
        Debug.WriteLine("Refreshing " & Me.DisplayName)
        Dim oldProperties As HashSet(Of [Property]) = _properties
        Dim oldDisplayName As String = Me.DisplayName
        Dim oldItemNameDisplaySortValue As String = Me.ItemNameDisplaySortValue
        Me.ClearCache()

        If Not Me.ShellItem2 Is Nothing Then
            _fullPath = Item.GetFullPathFromShellItem2(Me.ShellItem2)
            _attributes = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANLINK Or SFGAO.CANRENAME _
                Or SFGAO.CANDELETE Or SFGAO.DROPTARGET Or SFGAO.ENCRYPTED Or SFGAO.ISSLOW _
                Or SFGAO.LINK Or SFGAO.SHARE Or SFGAO.RDONLY Or SFGAO.HIDDEN Or SFGAO.FOLDER _
                Or SFGAO.FILESYSTEM Or SFGAO.HASSUBFOLDER Or SFGAO.COMPRESSED
            Me.ShellItem2.GetAttributes(_attributes, _attributes)
            If Me.DisplayName <> oldDisplayName Then
                Me.NotifyOfPropertyChange("DisplayName")
            End If
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
            Me.NotifyOfPropertyChange("StorageProviderUIStatusIcons16Async")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusFirstIcon16Async")
            Me.NotifyOfPropertyChange("StorageProviderUIStatusHasIconAsync")
            For Each prop In oldProperties
                Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}]", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].HasIcon", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Text", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Icons16Async", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].FirstIcon16Async", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}]", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].HasIcon", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Text", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Icons16Async", prop.Key.ToString()))
                Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].FirstIcon16Async", prop.Key.ToString()))
            Next
        Else
            Me.Dispose()
        End If
    End Sub

    Public ReadOnly Property FullPath As String
        Get
            Return _fullPath
        End Get
    End Property

    Public Function GetParent() As Folder
        If Not Me.FullPath.Equals(Shell.Desktop.FullPath) Then
            Dim parentShellItem2 As IShellItem2
            If Not Me.ShellItem2 Is Nothing Then
                Me.ShellItem2.GetParent(parentShellItem2)
            End If
            If Not parentShellItem2 Is Nothing Then
                Return New Folder(parentShellItem2, Nothing)
            End If
        End If

        Return Nothing
    End Function

    Public Property IsPinned As Boolean
        Get
            Return _isPinned
        End Get
        Set(value As Boolean)
            SetValue(_isPinned, value)
        End Set
    End Property

    Public ReadOnly Property OverlayIconIndex As Byte
        Get
            If IO.File.Exists(Me.FullPath) OrElse IO.Directory.Exists(Me.FullPath) Then
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
            Dim tcs As New TaskCompletionSource(Of Byte)

            Shell.PriorityTaskQueue.Add(
                Sub()
                    Try
                        tcs.SetResult(Me.OverlayIconIndex)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Dim overlayIconIndex As Byte = tcs.Task.Result
            If overlayIconIndex > 0 Then
                Return ImageHelper.GetOverlayIcon(overlayIconIndex, size)
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            If Not disposedValue Then
                Dim hbitmap As IntPtr
                Try
                    Dim h As HRESULT = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size, size), SIIGBF.SIIGBF_ICONONLY, hbitmap)
                    If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                        Return Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    End If
                Catch ex As Exception
                Finally
                    Functions.DeleteObject(hbitmap)
                End Try
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property IconAsync(size As Integer) As ImageSource
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource)

            Shell.PriorityTaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource
                        result = Me.Icon(size)
                        If Not result Is Nothing Then result.Freeze()
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Return tcs.Task.Result
        End Get
    End Property

    Public Overridable ReadOnly Property Image(size As Integer) As ImageSource
        Get
            If Not disposedValue Then
                Dim hbitmap As IntPtr
                Try
                    Dim h As HRESULT = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size, size), 0, hbitmap)
                    If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                        Return Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    End If
                Catch ex As Exception
                Finally
                    Functions.DeleteObject(hbitmap)
                End Try
            End If
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property ImageAsync(size As Integer) As ImageSource
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource)

            Shell.PriorityTaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource
                        result = Me.Image(size)
                        If Not result Is Nothing Then result.Freeze()
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Return tcs.Task.Result
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusIcons16Async As ImageSource()
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource())

            Shell.SlowTaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource()
                        result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2").Icons16Async
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Return tcs.Task.Result
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusFirstIcon16Async As ImageSource
        Get
            Dim tcs As New TaskCompletionSource(Of ImageSource)

            Shell.SlowTaskQueue.Add(
                Sub()
                    Try
                        Dim result As ImageSource
                        result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2").FirstIcon16Async
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Return tcs.Task.Result
        End Get
    End Property

    Public Overridable ReadOnly Property StorageProviderUIStatusHasIconAsync As Boolean
        Get
            Dim tcs As New TaskCompletionSource(Of Boolean)

            Shell.SlowTaskQueue.Add(
                Sub()
                    Try
                        Dim result As Boolean
                        result = Me.PropertiesByKeyAsText("e77e90df-6271-4f5b-834f-2dd1f245dda4:2").HasIcon
                        tcs.SetResult(result)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            Return tcs.Task.Result
        End Get
    End Property

    Public Overridable ReadOnly Property HasThumbnail As Boolean
        Get
            If Not disposedValue Then
                Dim hbitmap As IntPtr
                Try
                    Dim h As HRESULT = CType(Me.ShellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(1, 1), SIIGBF.SIIGBF_THUMBNAILONLY, hbitmap)
                    If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                        Return True
                    End If
                Catch ex As Exception
                Finally
                    Functions.DeleteObject(hbitmap)
                End Try
            End If
            Return False
        End Get
    End Property

    Public Overridable ReadOnly Property HasThumbnailAsync As Boolean
        Get
            Dim tcs As New TaskCompletionSource(Of Boolean)

            Shell.SlowTaskQueue.Add(
                Sub()
                    Try
                        tcs.SetResult(Me.HasThumbnail)
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)

            ' Wait for the result
            Return tcs.Task.Result
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
            Try
                'Debug.WriteLine("GetDisplayName for " & Me.FullPath)
                If String.IsNullOrWhiteSpace(_displayName) AndAlso Not disposedValue Then
                    Me.ShellItem2.GetDisplayName(SHGDN.NORMAL, _displayName)
                End If
                'Debug.WriteLine(Me.PropertiesByKeyAsText("fceff153-e839-4cf3-a9e7-ea22832094b8:123").Text)
            Catch ex As Exception
                ' sometimes the treeview will try to sort us just as we're in the process of disposing
            End Try

            Return _displayName
        End Get
    End Property

    Public ReadOnly Property AddressBarDisplayName As String
        Get
            If Not Shell.SpecialFolders.Values.FirstOrDefault(Function(f) f.FullPath = Me.FullPath) Is Nothing Then
                Return Me.DisplayName
            Else
                Dim specialFolderAsRoot As Folder = Shell.SpecialFolders.Values.FirstOrDefault(Function(f) Me.FullPath.StartsWith(f.FullPath))
                If Not specialFolderAsRoot Is Nothing Then
                    Return specialFolderAsRoot.DisplayName & Me.FullPath.Substring(specialFolderAsRoot.FullPath.Length)
                Else
                    Return Me.FullPath.TrimEnd(IO.Path.DirectorySeparatorChar)
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property IsDrive As Boolean
        Get
            Return Me.FullPath.Equals(Path.GetPathRoot(Me.FullPath)) AndAlso Not Me.FullPath.StartsWith("\\")
        End Get
    End Property

    Public ReadOnly Property IsFolder As Boolean
        Get
            Return TypeOf Me Is Folder
        End Get
    End Property

    Public Overridable ReadOnly Property ItemNameDisplaySortValue As String
        Get
            If Me.IsDrive Then
                Return Me.FullPath
            Else
                Return Me.DisplayName
            End If
        End Get
    End Property

    Public ReadOnly Property RelativePath As String
        Get
            Dim result As String
            Me.ShellItem2.GetDisplayName(SHGDN.FORPARSING Or SHGDN.INFOLDER, result)
            Return result
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

    Public ReadOnly Property Attributes As SFGAO
        Get
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

    'Public Overridable ReadOnly Property Properties(key As PROPERTYKEY) As [Property]
    '    Get
    '        Dim propertyDescription As IPropertyDescription
    '        Try
    '            Functions.PSGetPropertyDescription(key, GetType(IPropertyDescription).GUID, propertyDescription)
    '            If Not propertyDescription Is Nothing Then
    '                Dim canonicalName As String
    '                propertyDescription.GetCanonicalName(canonicalName)
    '                Return Me.Properties(canonicalName)
    '            Else
    '                Throw New Exception(String.Format("Property '{0}, {1}' not found.", key.fmtid.ToString(), key.pid))
    '            End If
    '        Finally
    '            If Not propertyDescription Is Nothing Then
    '                Marshal.ReleaseComObject(propertyDescription)
    '            End If
    '        End Try
    '    End Get
    'End Property

    Public ReadOnly Property ContentViewModeForBrowseProperties As [Property]()
        Get
            Dim PKEY_System_PropList_ContentViewModeForBrowse As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 13
            }
            Dim system_PropList_ContentViewModeForBrowse As String = Me.PropertiesByKey(PKEY_System_PropList_ContentViewModeForBrowse).Text
            If Not String.IsNullOrWhiteSpace(system_PropList_ContentViewModeForBrowse) Then
                Dim propertyNames() As String = system_PropList_ContentViewModeForBrowse.Substring(5).Split(";")
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
                Return properties.ToArray()
            Else
                Return {
                    Me.PropertiesByCanonicalName("System.ItemNameDisplay"),
                    Me.PropertiesByCanonicalName("System.ItemTypeText"),
                    Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"),
                    Me.PropertiesByCanonicalName("System.LayoutPattern.PlaceHolder"),
                    Me.PropertiesByCanonicalName("System.DateModified"),
                    Me.PropertiesByCanonicalName("System.Size")
                }
            End If
        End Get
    End Property

    Public ReadOnly Property InfoTip As String
        Get
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
            Dim System_StorageProviderUIStatus As System_StorageProviderUIStatusProperty
            Try
                System_StorageProviderUIStatus =
                    [Property].FromKey(System_StorageProviderUIStatusProperty.System_StorageProviderUIStatusKey, Me.PropertyStore)
                If System_StorageProviderUIStatus.RawValue.vt <> 0 Then
                    If Not String.IsNullOrWhiteSpace(System_StorageProviderUIStatus.Text) Then
                        text.Add(System_StorageProviderUIStatus.DisplayName & ": " & System_StorageProviderUIStatus.Text)
                    End If
                    If Not String.IsNullOrWhiteSpace(System_StorageProviderUIStatus.ActivityText) Then
                        text.Add(System_StorageProviderUIStatus.ActivityDisplayName & ": " & System_StorageProviderUIStatus.ActivityText)
                    End If
                End If
            Finally
                If Not System_StorageProviderUIStatus Is Nothing Then
                    System_StorageProviderUIStatus.Dispose()
                End If
            End Try
            Return String.Join(vbCrLf, text)
        End Get
    End Property

    Public ReadOnly Property TileViewProperties As String
        Get
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
        End Get
    End Property

    Public ReadOnly Property PropertyStore As IPropertyStore
        Get
            Dim ptr As IntPtr
            Try
                Me.ShellItem2.GetPropertyStore(0, GetType(IPropertyStore).GUID, ptr)
                If Not IntPtr.Zero.Equals(ptr) Then
                    Return Marshal.GetObjectForIUnknown(ptr)
                End If
            Finally
                If Not IntPtr.Zero.Equals(ptr) Then
                    Marshal.Release(ptr)
                End If
            End Try
            Return Nothing
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKeyAsText(propertyKey As String) As [Property]
        Get
            _propertiesLock.Wait()
            Try
                Dim [property] As [Property]
                Dim key As PROPERTYKEY = New PROPERTYKEY(propertyKey)
                [property] = _properties.FirstOrDefault(Function(p) p.Key.Equals(key))
                If [property] Is Nothing Then
                    [property] = [Property].FromKey(key, Me.ShellItem2)
                    If Not [property] Is Nothing Then _properties.Add([property])
                End If
                Return [property]
            Finally
                _propertiesLock.Release()
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKey(propertyKey As PROPERTYKEY) As [Property]
        Get
            _propertiesLock.Wait()
            Try
                Dim [property] As [Property] = _properties.FirstOrDefault(Function(p) p.Key.Equals(propertyKey))
                If [property] Is Nothing Then
                    [property] = [Property].FromKey(propertyKey, Me.ShellItem2)
                    If Not [property] Is Nothing Then _properties.Add([property])
                End If
                Return [property]
            Finally
                _propertiesLock.Release()
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByCanonicalName(canonicalName As String) As [Property]
        Get
            _propertiesLock.Wait()
            Try
                Dim [property] As [Property] = _properties.FirstOrDefault(Function(p) p.CanonicalName = canonicalName)
                If [property] Is Nothing Then
                    [property] = [Property].FromCanonicalName(canonicalName, Me.ShellItem2)
                    If Not [property] Is Nothing Then _properties.Add([property])
                End If
                Return [property]
            Finally
                _propertiesLock.Release()
            End Try
        End Get
    End Property

    Public Shared Async Function FromParsingNameDeepGetAsync(parsingName As String) As Task(Of Item)
        ' resolve environment variable?
        parsingName = Environment.ExpandEnvironmentVariables(parsingName)

        Dim specialFolder As Folder =
            Shell.SpecialFolders.Values.ToList().FirstOrDefault(Function(f) f.FullPath = parsingName)?.Clone()
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
                folder = Shell.SpecialFolders("Network").Clone()
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.SpecialFolders("This computer").Clone()
            Else
                ' root must be some special folder
                folder = Shell.SpecialFolders.Values.FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(0).ToLower() _
                                                                         OrElse f.FullPath.ToLower() = parts(0).ToLower())?.Clone()
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
            Shell.SpecialFolders.Values.ToList().FirstOrDefault(Function(f) f.FullPath = parsingName)?.Clone()
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
                folder = Shell.SpecialFolders("Network").Clone()
            ElseIf parts(0) = IO.Path.GetPathRoot(parsingName) Then
                ' this is a path on disk
                folder = Shell.SpecialFolders("This computer").Clone()
            Else
                ' root must be some special folder
                folder = Shell.SpecialFolders.Values.FirstOrDefault(Function(f) f.DisplayName.ToLower() = parts(0).ToLower() _
                                                                         OrElse f.FullPath.ToLower() = parts(0).ToLower())?.Clone()
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
                Case SHCNE.UPDATEITEM, SHCNE.FREESPACE, SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED
                    Using item1 = Item.FromPidl(e.Item1Pidl.AbsolutePIDL, Nothing)
                        If Me.Pidl.Equals(e.Item1Pidl) OrElse Me.FullPath?.Equals(item1.FullPath) Then
                            Me.Refresh()
                        End If
                    End Using
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    If Me.Pidl.Equals(e.Item1Pidl) Then
                        Dim oldPidl As Pidl = Me.Pidl
                        _pidl = e.Item2Pidl.Clone()
                        oldPidl.Dispose()
                        Me.Refresh()
                    End If
            End Select
        End If
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            disposedValue = True

            If disposing Then
                ' dispose managed state (managed objects)
                RemoveHandler Shell.Notification, AddressOf shell_Notification

                Me.ClearCache()
                For Each item In _expiredShellItem2
                    Marshal.ReleaseComObject(item)
                Next

                'If Me.FullPath = "C:\" AndAlso Not _logicalParent Is Nothing Then
                '    Dim i = 9
                'End If

                UIHelper.OnUIThreadAsync(
                    Sub()
                        If Not _logicalParent Is Nothing Then
                            _logicalParent._items.Remove(Me)
                            _logicalParent._isEnumerated = False
                        End If
                    End Sub)

                Shell.RemoveFromItemsCache(Me)
                Me.Pidl.Dispose()
            End If
            ' free unmanaged resources (unmanaged objects) and override finalizer
        End If
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
        If TypeOf Me Is Folder Then
            Return New Folder( Me.ShellItem2, _logicalParent) With {._shellItem2 = Nothing}
        Else
            Return New Item(Me.ShellItem2, _logicalParent) With {._shellItem2 = Nothing}
        End If
    End Function
End Class
