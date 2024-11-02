Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers

Public Class Item
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Protected Const MAX_PATH_LENGTH As Integer = 260

    Private _imageFactory As IShellItemImageFactory
    Protected _properties As List(Of [Property]) = New List(Of [Property])
    Protected _fullPath As String
    Friend disposedValue As Boolean
    Protected _parent As Folder
    Protected _displayName As String
    Friend _shellItem2 As IShellItem2
    Private _isPinned As Boolean
    Private _isCut As Boolean
    Private _attributes As SFGAO
    Private _overlaySmall As ImageSource
    Private _overlayLarge As ImageSource
    Private _overlayIconIndex As Integer?
    Private _treeRootIndex As Long = -1

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

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr, parentShellFolder As IShellFolder) As IShellItem2
        Dim ptr As IntPtr
        Try
            If parentShellFolder Is Nothing Then
                Functions.SHCreateItemFromIDList(pidl, Guids.IID_IShellItem2, ptr)
            Else
                Functions.SHCreateItemWithParent(IntPtr.Zero, parentShellFolder, pidl, Guids.IID_IShellItem2, ptr)
            End If
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
            parsingName = Environment.ExpandEnvironmentVariables(parsingName)
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
        If fullPath.StartsWith("::{") AndAlso fullPath.EndsWith("}") Then
            fullPath = "shell:" & fullPath
        End If
        Return fullPath
    End Function

    Public Sub New(shellItem2 As IShellItem2, parent As Folder)
        _shellItem2 = shellItem2
        If Not shellItem2 Is Nothing Then
            _fullPath = GetFullPathFromShellItem2(shellItem2)
            _attributes = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANLINK Or SFGAO.CANRENAME _
                Or SFGAO.CANDELETE Or SFGAO.DROPTARGET Or SFGAO.ENCRYPTED Or SFGAO.ISSLOW _
                Or SFGAO.LINK Or SFGAO.SHARE Or SFGAO.RDONLY Or SFGAO.HIDDEN Or SFGAO.FOLDER _
                Or SFGAO.FILESYSTEM Or SFGAO.HASSUBFOLDER Or SFGAO.COMPRESSED
            _shellItem2.GetAttributes(_attributes, _attributes)
        Else
            _fullPath = String.Empty
        End If
        _parent = parent
        AddHandler Shell.Notification, AddressOf shell_Notification
    End Sub

    Public ReadOnly Property IsVisibleInTree As Boolean
        Get
            Return Me.TreeRootIndex <> -1 OrElse (Not Me.Parent Is Nothing AndAlso Me.Parent.IsExpanded)
        End Get
    End Property

    Public Property TreeRootIndex As Long
        Get
            Return _treeRootIndex
        End Get
        Set(value As Long)
            SetValue(_treeRootIndex, value)
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
                Return _parent.TreeSortKey & itemNameDisplaySortValue & New String(" ", 260 - itemNameDisplaySortValue.Length)
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

    Public Overridable Sub ClearCache()
        If Not _imageFactory Is Nothing Then
            Marshal.ReleaseComObject(_imageFactory)
            _imageFactory = Nothing
        End If
        For Each [property] In _properties
            [property].Dispose()
        Next
        _properties = New List(Of [Property])()
        _displayName = Nothing
        _overlaySmall = Nothing
        _overlayLarge = Nothing
    End Sub

    Public Overridable Sub Refresh()
        Dim oldProperties As List(Of [Property]) = _properties
        Me.ClearCache()
        Dim bindCtxPtr As IntPtr
        Try
            Functions.CreateBindCtx(0, bindCtxPtr)
            _shellItem2.Update(bindCtxPtr)
        Finally
            If Not IntPtr.Zero.Equals(bindCtxPtr) Then
                Marshal.Release(bindCtxPtr)
            End If
        End Try

        _attributes = SFGAO.CANCOPY Or SFGAO.CANMOVE Or SFGAO.CANLINK Or SFGAO.CANRENAME _
                Or SFGAO.CANDELETE Or SFGAO.DROPTARGET Or SFGAO.ENCRYPTED Or SFGAO.ISSLOW _
                Or SFGAO.LINK Or SFGAO.SHARE Or SFGAO.RDONLY Or SFGAO.HIDDEN Or SFGAO.FOLDER _
                Or SFGAO.FILESYSTEM Or SFGAO.HASSUBFOLDER Or SFGAO.COMPRESSED
        _shellItem2.GetAttributes(_attributes, _attributes)
        For Each prop In Me.GetType().GetProperties()
            UIHelper.OnUIThread(
                Sub()
                    Me.NotifyOfPropertyChange(prop.Name)
                End Sub)
        Next
        For Each prop In oldProperties
            UIHelper.OnUIThread(
                Sub()
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByKeyAsText[{0}].Text", prop.Key.ToString()))
                    Me.NotifyOfPropertyChange(String.Format("PropertiesByCanonicalName[{0}].Text", prop.CanonicalName))
                End Sub)
        Next
    End Sub

    Public ReadOnly Property FullPath As String
        Get
            Return _fullPath
        End Get
    End Property

    Public ReadOnly Property Parent As Folder
        Get
            If Not Me.FullPath.Equals(Shell.Desktop.FullPath) Then
                If _parent Is Nothing AndAlso Not disposedValue AndAlso Not _shellItem2 Is Nothing Then
                    ' this is bound to the desktop
                    Dim parentShellItem2 As IShellItem2
                    Dim d As Boolean = disposedValue
                    _shellItem2.GetParent(parentShellItem2)
                    If Not parentShellItem2 Is Nothing Then
                        _parent = New Folder(parentShellItem2, Nothing)
                    End If
                End If

                Return _parent
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property PrimarySort As Integer
        Get
            Return 0
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

    Public Overridable ReadOnly Property OverlaySmall As ImageSource
        Get
            If _overlaySmall Is Nothing Then
                _overlaySmall = getOverlay(False)
            End If
            Return _overlaySmall
        End Get
    End Property

    Public Overridable ReadOnly Property OverlayLarge As ImageSource
        Get
            If _overlayLarge Is Nothing Then
                _overlayLarge = getOverlay(True)
            End If
            Return _overlayLarge
        End Get
    End Property

    Public Overridable ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            If Not disposedValue Then
                Dim ptr As IntPtr
                Try
                    CType(_shellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size, size), SIIGBF.SIIGBF_ICONONLY, ptr)
                    Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Finally
                    Functions.DeleteObject(ptr)
                End Try
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property Image(size As Integer) As ImageSource
        Get
            If Not disposedValue Then
                Dim ptr As IntPtr
                Try
                    CType(_shellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size, size), 0, ptr)
                    Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Finally
                    Functions.DeleteObject(ptr)
                End Try
            End If
        End Get
    End Property

    Protected Overridable Function getOverlay(isLarge As Boolean) As ImageSource
        If Not _overlayIconIndex.HasValue AndAlso Not disposedValue Then
            Dim pidl As IntPtr, lastpidl As IntPtr, ptr As IntPtr
            ptr = Marshal.GetIUnknownForObject(_shellItem2)
            Functions.SHGetIDListFromObject(ptr, pidl)
            lastpidl = Functions.ILFindLastID(pidl)

            Dim shellFolder As IShellFolder = If(Not Me.Parent Is Nothing, Me.Parent._shellFolder, Shell.Desktop._shellFolder)
            Try
                ptr = Marshal.GetIUnknownForObject(shellFolder)
                Dim ptr2 As IntPtr, shellIconOverlay As IShellIconOverlay
                Try
                    Marshal.QueryInterface(ptr, GetType(IShellIconOverlay).GUID, ptr2)
                    If Not IntPtr.Zero.Equals(ptr2) Then
                        Dim iconIndex As Integer

                        shellIconOverlay = Marshal.GetObjectForIUnknown(ptr2)
                        shellIconOverlay.GetOverlayIconIndex(lastpidl, iconIndex)

                        _overlayIconIndex = iconIndex
                    Else
                        Return Nothing
                    End If
                Finally
                    If Not IntPtr.Zero.Equals(ptr2) Then
                        Marshal.Release(ptr2)
                    End If
                    If Not shellIconOverlay Is Nothing Then
                        Marshal.ReleaseComObject(shellIconOverlay)
                    End If
                End Try
            Finally
                Marshal.Release(ptr)
            End Try
        End If

        If _overlayIconIndex.HasValue AndAlso _overlayIconIndex > 0 Then
            ' Get the system image list
            Dim hImageListLarge As IntPtr
            Dim hImageListSmall As IntPtr
            Functions.Shell_GetImageLists(hImageListLarge, hImageListSmall)

            ' Retrieve the overlay icon
            Dim hIcon As IntPtr
            Try
                If isLarge Then
                    hIcon = Functions.ImageList_GetIcon(hImageListLarge, _overlayIconIndex, 0)
                Else
                    hIcon = Functions.ImageList_GetIcon(hImageListSmall, _overlayIconIndex, 0)
                End If
                If hIcon <> IntPtr.Zero Then
                    Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                        Return Interop.Imaging.CreateBitmapSourceFromHBitmap(icon.ToBitmap().GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    End Using
                Else
                    Return Nothing
                End If
            Finally
                If hIcon <> IntPtr.Zero Then
                    Functions.DestroyIcon(hIcon)
                End If
            End Try
        Else
            Return Nothing
        End If
    End Function

    Public Overridable ReadOnly Property DisplayName As String
        Get
            Try
                If String.IsNullOrWhiteSpace(_displayName) AndAlso Not disposedValue Then
                    _shellItem2.GetDisplayName(SHGDN.NORMAL, _displayName)
                End If
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
            Return Me.FullPath.Equals(Path.GetPathRoot(Me.FullPath))
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
            _shellItem2.GetDisplayName(SHGDN.FORPARSING Or SHGDN.INFOLDER, result)
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

    Public ReadOnly Property InfoTip As String
        Get
            Dim PKEY_System_InfoTipText As New PROPERTYKEY() With {
                        .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                        .pid = 4
                    }
            Dim properties() As String = Me.PropertiesByKey(PKEY_System_InfoTipText).Text.Substring(5).Split(";")
            Dim text As List(Of String) = New List(Of String)()
            Dim i As Integer = 0
            For Each propCanonicalName In properties
                Dim prop As [Property] = Me.PropertiesByCanonicalName(propCanonicalName)
                If Not prop Is Nothing AndAlso Not String.IsNullOrWhiteSpace(prop.Text) Then
                    text.Add(prop.DescriptionDisplayName & ": " & prop.Text)
                End If
                i += 1
            Next
            Return String.Join(vbCrLf, text)
        End Get
    End Property

    Public ReadOnly Property TileViewProperties As String
        Get
            Dim PKEY_System_PropList_TileInfo As New PROPERTYKEY() With {
                .fmtid = New Guid("C9944A21-A406-48FE-8225-AEC7E24C211B"),
                .pid = 3
            }
            Dim properties() As String = Me.PropertiesByKey(PKEY_System_PropList_TileInfo).Text.Substring(5).Split(";")
            Dim text As List(Of String) = New List(Of String)()
            Dim i As Integer = 0
            For Each propCanonicalName In properties
                If Not propCanonicalName.StartsWith("*") Then
                    Dim prop As [Property] = Me.PropertiesByCanonicalName(propCanonicalName)
                    If Not prop Is Nothing AndAlso Not String.IsNullOrWhiteSpace(prop.Text) Then
                        text.Add(If(i >= 2, prop.DescriptionDisplayName & ": ", "") & prop.Text)
                    End If
                End If
                i += 1
            Next
            Return String.Join(vbCrLf, text)
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKeyAsText(propertyKey As String) As [Property]
        Get
            Dim parts() As String = propertyKey.Split(":")
            Dim key As PROPERTYKEY
            key.fmtid = New Guid(parts(0))
            key.pid = parts(1)

            Dim [property] As [Property] = _properties.FirstOrDefault(Function(p) p.Key.Equals(key))
            If [property] Is Nothing Then
                [property] = [Property].FromKey(key, Me)
                _properties.Add([property])
            End If
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByKey(propertyKey As PROPERTYKEY) As [Property]
        Get
            Dim [property] As [Property] = _properties.FirstOrDefault(Function(p) p.Key.Equals(propertyKey))
            If [property] Is Nothing Then
                [property] = [Property].FromKey(propertyKey, Me)
                _properties.Add([property])
            End If
            Return [property]
        End Get
    End Property

    Public Overridable ReadOnly Property PropertiesByCanonicalName(canonicalName As String) As [Property]
        Get
            Dim [property] As [Property] = _properties.FirstOrDefault(Function(p) p.CanonicalName = canonicalName)
            If [property] Is Nothing Then
                [property] = [Property].FromCanonicalName(canonicalName, Me)
                If Not [property] Is Nothing Then _properties.Add([property])
            End If
            Return [property]
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
        If Not _shellItem2 Is Nothing AndAlso Not disposedValue Then
            Select Case e.Event
                Case SHCNE.UPDATEITEM, SHCNE.FREESPACE, SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED
                    If Me.FullPath.Equals(e.Item1Path) Then
                        Me.Refresh()
                    End If
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    If Me.FullPath.Equals(e.Item1Path) Then
                        _fullPath = e.Item2Path
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
            End If
            If Me.FullPath.ToLower().Contains("office") Then
                Dim i As Int16 = 9
            End If
            ' free unmanaged resources (unmanaged objects) and override finalizer
            If Not _imageFactory Is Nothing Then
                Marshal.ReleaseComObject(_imageFactory)
                _imageFactory = Nothing
            End If
            If Not _shellItem2 Is Nothing Then
                Marshal.ReleaseComObject(_shellItem2)
                _shellItem2 = Nothing
            End If
            'If Not _list Is Nothing Then
            '    _list.Remove(Me)
            'End If
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
            Return New Folder(Folder.GetIShellItem2FromParsingName(Me.FullPath), _parent)
        Else
            Return New Item(Item.GetIShellItem2FromParsingName(Me.FullPath), _parent)
        End If
    End Function
End Class
