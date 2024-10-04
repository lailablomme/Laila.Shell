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

    Private _parent As Folder
    Private _imageFactory As IShellItemImageFactory
    Protected _properties As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Protected _fullPath As String
    Friend disposedValue As Boolean
    Protected _logicalParent As Folder
    Protected _displayName As String
    Friend _shellItem2 As IShellItem2
    Private _isPinned As Boolean
    Private _isCut As Boolean
    Private _attributes As SFGAO
    Private _icon As Dictionary(Of Integer, ImageSource) = New Dictionary(Of Integer, ImageSource)()
    Private _overlaySmall As ImageSource
    Private _overlayLarge As ImageSource
    Protected _cachedIconSize As Integer

    Public Shared Function FromParsingName(parsingName As String, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean),
                                           cachedIconSize As Integer) As Item
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        If Not shellItem2 Is Nothing Then
            Dim attr As Integer = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If CBool(attr And SFGAO.FOLDER) Then
                Return New Folder(shellItem2, logicalParent, setIsLoadingAction, cachedIconSize)
            Else
                Return New Item(shellItem2, logicalParent, cachedIconSize)
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

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, cachedIconSize As Integer)
        _shellItem2 = shellItem2
        If Not shellItem2 Is Nothing Then
            _fullPath = GetFullPathFromShellItem2(shellItem2)
            _attributes = SFGAO.HIDDEN Or SFGAO.COMPRESSED Or SFGAO.CANCOPY Or SFGAO.CANMOVE _
                Or SFGAO.CANLINK Or SFGAO.HASSUBFOLDER Or SFGAO.ISSLOW
            _shellItem2.GetAttributes(_attributes, _attributes)
        Else
            _fullPath = String.Empty
        End If
        UIHelper.OnUIThread(
            Sub()
                If cachedIconSize <> 0 Then
                    _cachedIconSize = cachedIconSize
                    _icon(cachedIconSize) = Me.Icon(cachedIconSize)
                    If cachedIconSize < 32 Then
                        _overlaySmall = Me.OverlaySmall
                    Else
                        _overlayLarge = Me.OverlayLarge
                    End If
                End If
            End Sub)
        _logicalParent = logicalParent
        AddHandler Shell.Notification, AddressOf shell_Notification
    End Sub

    Public ReadOnly Property FullPath As String
        Get
            Return _fullPath
        End Get
    End Property

    Public ReadOnly Property LogicalParent As Folder
        Get
            Return _logicalParent
        End Get
    End Property

    Public ReadOnly Property Parent As Folder
        Get
            If Not Me.FullPath.Equals(Shell.Desktop.FullPath) Then
                If _parent Is Nothing Then
                    ' this is bound to the desktop
                    Dim parentShellItem2 As IShellItem2
                    _shellItem2.GetParent(parentShellItem2)
                    If Not parentShellItem2 Is Nothing Then
                        _parent = New Folder(parentShellItem2, Nothing, Nothing, 0)
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
            If Not _icon.ContainsKey(size) Then
                Dim ptr As IntPtr
                Try
                    CType(_shellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(size, size), SIIGBF.SIIGBF_ICONONLY, ptr)
                    _icon.Add(size, Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()))
                Finally
                    Functions.DeleteObject(ptr)
                End Try
            End If
            Return _icon(size)
        End Get
    End Property

    Protected Overridable Function getOverlay(isLarge As Boolean) As ImageSource
        Dim pidl As IntPtr, lastpidl As IntPtr, ptr As IntPtr
        ptr = Marshal.GetIUnknownForObject(_shellItem2)
        Functions.SHGetIDListFromObject(ptr, Pidl)
        lastpidl = Functions.ILFindLastID(Pidl)

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

                    If iconIndex > 0 Then
                        ' Get the system image list
                        Dim hImageListLarge As IntPtr
                        Dim hImageListSmall As IntPtr
                        Functions.Shell_GetImageLists(hImageListLarge, hImageListSmall)

                        ' Retrieve the overlay icon
                        Dim hIcon As IntPtr
                        Try
                            If isLarge Then
                                hIcon = Functions.ImageList_GetIcon(hImageListLarge, iconIndex, 0)
                            Else
                                hIcon = Functions.ImageList_GetIcon(hImageListSmall, iconIndex, 0)
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
    End Function

    Public Overridable ReadOnly Property DisplayName As String
        Get
            If String.IsNullOrWhiteSpace(_displayName) Then
                _shellItem2.GetDisplayName(SHGDN.NORMAL, _displayName)
            End If
            'Debug.WriteLine(_displayName)
            Return _displayName
        End Get
    End Property

    Public Overridable ReadOnly Property ItemNameDisplaySortValue As String
        Get
            If Me.FullPath.Equals(Path.GetPathRoot(Me.FullPath)) Then
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

    Public ReadOnly Property IsExecutable As Boolean
        Get
            Dim executableExtensions As List(Of String) = New List(Of String)() From {
                ".exe", ".bat", ".cmd", ".com", ".msi"
            }
            Return executableExtensions.Contains(IO.Path.GetExtension(Me.FullPath).ToLower())
        End Get
    End Property

    Public Sub Execute(Optional arguments As String = Nothing)
        Dim psi As ProcessStartInfo = New ProcessStartInfo() With {
            .FileName = Me.FullPath,
            .Arguments = arguments,
            .UseShellExecute = True
        }
        Process.Start(psi)
    End Sub

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

    Public Overridable ReadOnly Property Properties(canonicalName As String) As [Property]
        Get
            Dim kv As KeyValuePair(Of String, [Property]) = _properties.FirstOrDefault(Function(kvp) kvp.Key.Equals(canonicalName))
            If Not String.IsNullOrEmpty(kv.Key) Then
                Return kv.Value
            Else
                Dim [property] As [Property] = [Property].FromCanonicalName(canonicalName, Me)
                _properties.Add(canonicalName, [property])
                Return [property]
            End If
        End Get
    End Property

    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is Item Then
            Return EqualityComparer(Of String).Default.Equals(Me.FullPath, CType(obj, Item).FullPath)
        Else
            Return False
        End If
    End Function

    Protected Overridable Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        If Not _shellItem2 Is Nothing AndAlso Not disposedValue Then
            Select Case e.Event
                Case SHCNE.UPDATEITEM, SHCNE.FREESPACE, SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED
                    If Me.FullPath.Equals(e.Item1Path) Then
                        Marshal.ReleaseComObject(_shellItem2)
                        _shellItem2 = Item.GetIShellItem2FromParsingName(_fullPath)
                        _properties = New Dictionary(Of String, [Property])()
                        _displayName = Nothing
                        For Each prop In Me.GetType().GetProperties()
                            Me.NotifyOfPropertyChange(prop.Name)
                        Next
                        If Not Me.Parent Is Nothing Then
                            For Each column In Me.Parent.Columns
                                Me.NotifyOfPropertyChange(String.Format("Properties[{0}].Text", column.CanonicalName))
                            Next
                        End If
                    End If
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    If Me.FullPath.Equals(e.Item1Path) Then
                        _fullPath = e.Item2Path
                        Marshal.ReleaseComObject(_shellItem2)
                        _shellItem2 = Item.GetIShellItem2FromParsingName(_fullPath)
                        _properties = New Dictionary(Of String, [Property])()
                        _displayName = Nothing
                        For Each prop In Me.GetType().GetProperties()
                            Me.NotifyOfPropertyChange(prop.Name)
                        Next
                        If Not Me.Parent Is Nothing Then
                            For Each column In Me.Parent.Columns
                                Me.NotifyOfPropertyChange(String.Format("Properties[{0}].Text", column.CanonicalName))
                            Next
                        End If
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
End Class
