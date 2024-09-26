Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class Item
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Protected Const MAX_PATH_LENGTH As Integer = 260

    Private _parent As Folder
    Private _imageFactory As IShellItemImageFactory
    Protected _properties As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Protected _fullPath As String
    Protected _setIsLoadingAction As Action(Of Boolean)
    Protected disposedValue As Boolean
    Protected _logicalParent As Folder
    Protected _displayName As String
    Friend _shellItem2 As IShellItem2

    Public Shared Function FromParsingName(parsingName As String, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean)) As Item
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        If Not shellItem2 Is Nothing Then
            Dim attr As Integer = SFGAO.FOLDER
            shellItem2.GetAttributes(attr, attr)
            If CBool(attr And SFGAO.FOLDER) Then
                Return New Folder(shellItem2, logicalParent, setIsLoadingAction)
            Else
                Return New Item(shellItem2, logicalParent, setIsLoadingAction)
            End If
        Else
            Return Nothing
        End If
    End Function

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr, bindingParent As Folder) As IShellItem2
        Dim ptr As IntPtr
        Try
            Functions.SHCreateItemWithParent(IntPtr.Zero, bindingParent._shellFolder, pidl, Guids.IID_IShellItem2, ptr)
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

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        _shellItem2 = shellItem2
        If Not shellItem2 Is Nothing Then
            _fullPath = GetFullPathFromShellItem2(shellItem2)
        Else
            _fullPath = String.Empty
        End If
        _setIsLoadingAction = setIsLoadingAction
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
                        _parent = New Folder(parentShellItem2, Nothing, _setIsLoadingAction)
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

    Public ReadOnly Property IsPinned As Boolean
        Get
            If Not _shellItem2 Is Nothing Then
                Using val = Me.Properties("System.IsPinnedToNameSpaceTree").RawValue
                    Return val.union.boolVal
                End Using
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property Icon16 As ImageSource
        Get
            Dim ptr As IntPtr
            Try
                CType(_shellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(16, 16), SIIGBF.SIIGBF_ICONONLY, ptr)
                Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            Finally
                Functions.DeleteObject(ptr)
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property Overlay16 As ImageSource
        Get
            Return getOverlay(False)
        End Get
    End Property

    Public Overridable ReadOnly Property Icon32 As ImageSource
        Get
            Dim ptr As IntPtr
            Try
                CType(_shellItem2, IShellItemImageFactory).GetImage(New System.Drawing.Size(32, 32), SIIGBF.SIIGBF_ICONONLY, ptr)
                Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            Finally
                Functions.DeleteObject(ptr)
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property Overlay32 As ImageSource
        Get
            Return getOverlay(True)
        End Get
    End Property

    Protected Overridable Function getOverlay(isLarge As Boolean) As ImageSource
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

    Public ReadOnly Property IsHidden As Boolean
        Get
            Dim attr As SFGAO = SFGAO.HIDDEN
            _shellItem2.GetAttributes(attr, attr)
            Return attr.HasFlag(SFGAO.HIDDEN)
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

    Public Overridable ReadOnly Property Properties(canonicalName As String) As [Property]
        Get
            If _properties.Keys.Contains(canonicalName) Then
                Return _properties(canonicalName)
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
        Select Case e.Event
            Case SHCNE.UPDATEITEM, SHCNE.FREESPACE, SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED
                If Me.FullPath.Equals(e.Item1Path) Then
                    _properties = New Dictionary(Of String, [Property])()
                    _displayName = Nothing
                    For Each prop In Me.GetType().GetProperties()
                        Me.NotifyOfPropertyChange(prop.Name)
                    Next
                End If
            Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                If Me.FullPath.Equals(e.Item1Path) Then
                    _fullPath = e.Item2Path
                    _properties = New Dictionary(Of String, [Property])()
                    _displayName = Nothing
                    For Each prop In Me.GetType().GetProperties()
                        Me.NotifyOfPropertyChange(prop.Name)
                    Next
                End If
        End Select
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' dispose managed state (managed objects)
                RemoveHandler Shell.Notification, AddressOf shell_Notification
            End If

            ' free unmanaged resources (unmanaged objects) and override finalizer
            If Not _imageFactory Is Nothing Then
                Marshal.ReleaseComObject(_imageFactory)
            End If
            If Not _shellItem2 Is Nothing Then
                Marshal.ReleaseComObject(_shellItem2)
            End If
            disposedValue = True
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    Protected Overrides Sub Finalize()
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=False)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
