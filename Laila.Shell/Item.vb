Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class Item
    Inherits NotifyPropertyChangedBase
    Implements IDisposable

    Protected Const MAX_PATH_LENGTH As Integer = 260

    Protected _bindingParent As Folder
    Private _parent As Folder
    Friend _pidl As IntPtr
    Private _interface As IShellItem2
    Private _imageFactory As IShellItemImageFactory
    Protected _properties As Dictionary(Of String, [Property]) = New Dictionary(Of String, [Property])
    Protected _fullPath As String
    Protected _setIsLoadingAction As Action(Of Boolean)
    Protected disposedValue As Boolean

    Public Shared Function FromParsingName(parsingName As String, setIsLoadingAction As Action(Of Boolean)) As Item
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        Dim attr As Integer = SFGAO.FOLDER
        shellItem2.GetAttributes(attr, attr)
        If CBool(attr And SFGAO.FOLDER) Then
            Return New Folder(Folder.GetIShellFolderFromIShellItem2(shellItem2), shellItem2, setIsLoadingAction)
        Else
            Return New Item(shellItem2, setIsLoadingAction)
        End If
    End Function

    Friend Shared Function GetIShellItem2FromPidl(pidl As IntPtr, bindingParent As Folder) As IShellItem2
        Dim ptr As IntPtr
        Functions.SHCreateItemWithParent(IntPtr.Zero, bindingParent.ShellFolder, pidl, Guids.IID_IShellItem2, ptr)
        Return Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellItem2))
    End Function

    Friend Shared Function GetIShellItem2FromParsingName(parsingName As String) As IShellItem2
        Dim ptr As IntPtr
        Functions.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, Guids.IID_IShellItem2, ptr)
        Return If(Not IntPtr.Zero.Equals(ptr), Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellItem2)), Nothing)
    End Function

    Public Sub New(shellItem2 As IShellItem2, setIsLoadingAction As Action(Of Boolean))
        _interface = shellItem2
        _setIsLoadingAction = setIsLoadingAction
        Dim s As String = Me.FullPath
        AddHandler Shell.Notification, AddressOf shell_Notification
    End Sub

    Public Sub New(bindingParent As Folder, pidl As IntPtr, setIsLoadingAction As Action(Of Boolean))
        _bindingParent = bindingParent
        _pidl = pidl
        _setIsLoadingAction = setIsLoadingAction
        Dim s As String = Me.FullPath
        AddHandler Shell.Notification, AddressOf shell_Notification
    End Sub

    Friend ReadOnly Property ShellItem2 As IShellItem2
        Get
            If _interface Is Nothing Then
                If _pidl.Equals(IntPtr.Zero) Then
                    Dim result As IShellItem
                    Functions.SHGetKnownFolderItem(Guids.KnownFolder_Desktop, 0, IntPtr.Zero, GetType(IShellItem).GUID, result)
                    _interface = result
                Else
                    _interface = GetIShellItem2FromPidl(_pidl, _bindingParent)
                    Marshal.FreeCoTaskMem(_pidl)
                End If
            End If

            Return _interface
        End Get
    End Property

    Friend ReadOnly Property ImageFactory As IShellItemImageFactory
        Get
            If _imageFactory Is Nothing Then
                Dim ptr3 As IntPtr
                Functions.SHCreateItemFromParsingName(Me.FullPath, IntPtr.Zero, GetType(IShellItemImageFactory).GUID, ptr3)
                _imageFactory = Marshal.GetTypedObjectForIUnknown(ptr3, GetType(IShellItemImageFactory))
            End If

            Return _imageFactory
        End Get
    End Property

    Public ReadOnly Property Parent As Folder
        Get
            If Not Me.FullPath.Equals(Shell.Desktop.FullPath) Then
                If _parent Is Nothing Then
                    ' this is bound to the desktop
                    Dim parentShellItem2 As IShellItem2
                    Me.ShellItem2.GetParent(parentShellItem2)
                    If Not parentShellItem2 Is Nothing Then
                        Dim shellFolder As IShellFolder = Folder.GetIShellFolderFromIShellItem2(parentShellItem2)
                        _parent = New Folder(shellFolder, parentShellItem2, _setIsLoadingAction)
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
            Return If(TypeOf Me Is LevelUpFolder, 0, 1)
        End Get
    End Property

    Public Overridable ReadOnly Property Icon16 As ImageSource
        Get
            Dim ptr As IntPtr
            Try
                Me.ImageFactory.GetImage(New System.Drawing.Size(16, 16), SIIGBF.SIIGBF_ICONONLY, ptr)
                Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            Finally
                Functions.DeleteObject(ptr)
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property Overlay16 As ImageSource
        Get
            Dim pidl As IntPtr, lastpidl As IntPtr
            Functions.SHGetIDListFromObject(Marshal.GetIUnknownForObject(Me.ShellItem2), pidl)
            lastpidl = Functions.ILFindLastID(pidl)

            Dim shellIconOverlay As IShellIconOverlay = If(Not Me.Parent Is Nothing, Me.Parent.ShellFolder, Shell.Desktop.ShellFolder)
            Dim iconIndex As Integer
            shellIconOverlay.GetOverlayIconIndex(lastpidl, iconIndex)

            If iconIndex >= 0 Then
                ' Get the system image list
                Dim hImageListLarge As IntPtr
                Dim hImageListSmall As IntPtr
                Functions.Shell_GetImageLists(hImageListLarge, hImageListSmall)

                ' Retrieve the overlay icon
                Dim hIcon As IntPtr = Functions.ImageList_GetIcon(hImageListSmall, iconIndex, 0)
                If hIcon <> IntPtr.Zero Then
                    Try
                        Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                            Return Interop.Imaging.CreateBitmapSourceFromHBitmap(icon.ToBitmap().GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        End Using
                    Finally
                        Functions.DestroyIcon(hIcon)
                    End Try
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property Icon32 As ImageSource
        Get
            Dim ptr As IntPtr
            Try
                Me.ImageFactory.GetImage(New System.Drawing.Size(32, 32), SIIGBF.SIIGBF_ICONONLY, ptr)
                Return Interop.Imaging.CreateBitmapSourceFromHBitmap(ptr, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
            Finally
                Functions.DeleteObject(ptr)
            End Try
        End Get
    End Property

    Public Overridable ReadOnly Property Overlay32 As ImageSource
        Get
            Dim pidl As IntPtr, lastpidl As IntPtr
            Functions.SHGetIDListFromObject(Marshal.GetIUnknownForObject(Me.ShellItem2), pidl)
            lastpidl = Functions.ILFindLastID(pidl)

            Dim shellIconOverlay As IShellIconOverlay = Me.Parent.ShellFolder
            Dim iconIndex As Integer
            shellIconOverlay.GetOverlayIconIndex(lastpidl, iconIndex)

            If iconIndex >= 0 Then
                ' Get the system image list
                Dim hImageListLarge As IntPtr
                Dim hImageListSmall As IntPtr
                Functions.Shell_GetImageLists(hImageListLarge, hImageListSmall)

                ' Retrieve the overlay icon
                Dim hIcon As IntPtr = Functions.ImageList_GetIcon(hImageListLarge, iconIndex, 0)
                If hIcon <> IntPtr.Zero Then
                    Try
                        Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                            Return Interop.Imaging.CreateBitmapSourceFromHBitmap(icon.ToBitmap().GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        End Using
                    Finally
                        Functions.DestroyIcon(hIcon)
                    End Try
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayName As String
        Get
            Dim result As String
            Me.ShellItem2.GetDisplayName(SHGDN.NORMAL, result)
            Return result
        End Get
    End Property

    Public Overridable ReadOnly Property ItemNameDisplaySortValue As String
        Get
            Dim fullPath As String = Me.FullPath
            If fullPath.Equals(Path.GetPathRoot(fullPath)) Then
                Return Me.FullPath
            Else
                Return Me.DisplayName
            End If
        End Get
    End Property

    Public ReadOnly Property FullPath As String
        Get
            If _fullPath Is Nothing Then
                Me.ShellItem2.GetDisplayName(SHGDN.FORPARSING, _fullPath)
            End If

            Return _fullPath
        End Get
    End Property

    Public ReadOnly Property RelativePath As String
        Get
            Dim result As String
            Me.ShellItem2.GetDisplayName(SHGDN.FORPARSING Or SHGDN.INFOLDER, result)
            Return result
        End Get
    End Property

    Public ReadOnly Property IsHidden As Boolean
        Get
            Dim attr As Integer = SFGAO.HIDDEN
            Me.ShellItem2.GetAttributes(attr, attr)
            Return CBool(attr And SFGAO.HIDDEN)
        End Get
    End Property

    Public Overridable ReadOnly Property Properties(key As PROPERTYKEY) As [Property]
        Get
            Dim propertyDescription As IPropertyDescription
            Functions.PSGetPropertyDescription(key, GetType(IPropertyDescription).GUID, propertyDescription)
            If Not propertyDescription Is Nothing Then
                Dim canonicalName As String
                propertyDescription.GetCanonicalName(canonicalName)
                Return Me.Properties(canonicalName)
            Else
                Throw New Exception(String.Format("Property '{0}, {1}' not found.", key.fmtid.ToString(), key.pid))
            End If
        End Get
    End Property

    Public Overridable ReadOnly Property Properties(canonicalName As String) As [Property]
        Get
            If _properties.Keys.Contains(canonicalName) Then
                Return _properties(canonicalName)
            Else
                Dim [property] As [Property] = [Property].FromCanonicalName(canonicalName, Me)
                '[property]._innerPropertyKey = Me.Parent.Columns(canonicalName).PROPERTYKEY
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
                    Me.ShellItem2.Update(IntPtr.Zero)
                    _fullPath = Nothing
                    _properties = New Dictionary(Of String, [Property])()
                    Me.NotifyOfPropertyChange("Properties")
                End If
            Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                If Me.FullPath.Equals(e.Item1Path) Then
                    _interface = Item.GetIShellItem2FromParsingName(e.Item2Path)
                    _fullPath = Nothing
                    _properties = New Dictionary(Of String, [Property])()
                    Me.NotifyOfPropertyChange("Properties")
                End If
        End Select
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
                RemoveHandler Shell.Notification, AddressOf shell_Notification

                If Not _parent Is Nothing Then
                    _parent.Dispose()
                End If
            End If

            Marshal.ReleaseComObject(Me.ShellItem2)
            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    'Protected Function getPropertyValue(propVariant As PROPVARIANT, canonicalName As String) As Object
    '    Select Case propVariant.vt
    '        Case VarEnum.VT_EMPTY
    '            Return String.Empty
    '        Case VarEnum.VT_BSTR
    '            Return Marshal.PtrToStringBSTR(propVariant.union.ptr)
    '        Case VarEnum.VT_LPWSTR
    '            Return Marshal.PtrToStringUni(propVariant.union.ptr)
    '        Case VarEnum.VT_FILETIME
    '            Return DateTime.FromFileTime(propVariant.union.hVal)
    '        Case VarEnum.VT_UI4
    '            Return propVariant.union.uintVal
    '        Case VarEnum.VT_UI8
    '            Return propVariant.union.uhVal
    '        Case VarEnum.VT_BLOB
    '            Select Case canonicalName
    '                Case "System.StorageProviderUIStatus"
    '                    Dim ptr As IntPtr, propertyStore As IPropertyStore, psps As IPersistSerializedPropStorage
    '                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyStore).GUID, ptr)
    '                    propertyStore = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyStore))
    '                    'Me.ShellItem2.GetPropertyStore(0, GetType(IPropertyStore).GUID, ptr)
    '                    'persistStream = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPersistStream))
    '                    'propertyStore = persistStream
    '                    psps = propertyStore
    '                    'Debug.WriteLine("first uint32,,," & Convert.ToUInt32(Marshal.ReadInt32(propVariant.union.bstrblobVal.pData)))
    '                    'Dim ptr2 As IntPtr = Functions.SHCreateMemStream(propVariant.union.bstrblobVal.pData, propVariant.union.bstrblobVal.cbSize)
    '                    'Dim stream As IStream = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IStream))
    '                    'stream.Seek(0, 0, Nothing)
    '                    'Dim b(0) As Byte, bytes(propVariant.union.bstrblobVal.cbSize - 1) As Byte, r As UInteger
    '                    'bytes(0) = BitConverter.GetBytes(propVariant.union.bstrblobVal.cbSize)(0)
    '                    'bytes(1) = BitConverter.GetBytes(propVariant.union.bstrblobVal.cbSize)(1)
    '                    'bytes(2) = BitConverter.GetBytes(propVariant.union.bstrblobVal.cbSize)(2)
    '                    'bytes(3) = BitConverter.GetBytes(propVariant.union.bstrblobVal.cbSize)(3)
    '                    'For y As Integer = 0 To propVariant.union.bstrblobVal.cbSize - 1
    '                    '    '    stream.Read(b, 1, r)
    '                    '    bytes(y) = Marshal.ReadByte(propVariant.union.bstrblobVal.pData, y)
    '                    '    '    Debug.WriteLine(b(0) & ",,," & r)
    '                    'Next
    '                    'File.WriteAllBytes("c:\repos\test2.txt", bytes)
    '                    'Dim pinnedArray As GCHandle = GCHandle.Alloc(bytes, GCHandleType.Pinned)
    '                    'Dim p As IntPtr = pinnedArray.AddrOfPinnedObject()
    '                    'psps.SetPropertyStorage(p, propVariant.union.bstrblobVal.cbSize + 4)
    '                    'stream.Seek(0, 0, Nothing)
    '                    'Dim st As System.Runtime.InteropServices.ComTypes.STATSTG
    '                    'stream.Stat(st, 0)
    '                    'Dim l As Integer = propertyStore.Load(stream)
    '                    Dim pc As UInt32
    '                    'For z = 0 To 30
    '                    psps.SetFlags(0)
    '                    psps.SetPropertyStorage(propVariant.union.bstrblobVal.pData,
    '                                            propVariant.union.bstrblobVal.cbSize)
    '                    propertyStore.GetCount(pc)
    '                    '    If pc <> 0 Then
    '                    '        Debug.WriteLine("gotit")
    '                    '    End If
    '                    '    Debug.WriteLine(pc)
    '                    'Next
    '                    'Dim s = SystemProperties.GetPropertyDescription("System.StorageProviderState")
    '                    'For Each ent In s.PropertyEnumTypes
    '                    '    Debug.WriteLine(ent.DisplayText)
    '                    '    Debug.WriteLine(ent.EnumType.ToString())
    '                    'Next

    '                    Dim desc As IPropertyDescription
    '                    Dim k As PROPVARIANT
    '                    Dim dt2 As PropertyDisplayType
    '                    Dim etl As IPropertyEnumTypeList, etlptr As IntPtr
    '                    'Debug.WriteLine("System.StorageProviderState")
    '                    'Functions.PSGetPropertyDescriptionByName("System.StorageProviderState", GetType(IPropertyDescription).GUID, desc)
    '                    'Dim kk As PROPERTYKEY
    '                    'desc.GetPropertyKey(kk)
    '                    'Dim so = ShellObject.FromParsingName(Me.FullPath)
    '                    'Debug.WriteLine(so.Properties.GetProperty("System.StorageProviderState").IconReference.ToString())
    '                    ''Debug.WriteLine("val" & Me.PropertyValues("System.PropList.StatusIcons").ToString())
    '                    Dim ppptr As IntPtr, ppp As IPropertyStore
    '                    Me.ShellItem2.GetPropertyStore(0, GetType(IPropertyStore).GUID, ppptr)
    '                    ppp = Marshal.GetTypedObjectForIUnknown(ppptr, GetType(IPropertyStore))
    '                    'ppp.GetValue(kk, k)
    '                    'Debug.WriteLine("val " & getPropertyValue(k, "System.StorageProviderState"))
    '                    'Dim pd22 As IPropertyDescription2 = desc, img2 As String
    '                    'pd22.GetImageReferenceForValue(k, img2)
    '                    'Debug.WriteLine("img" & img2)
    '                    'desc.GetDisplayType(dt2)
    '                    'Debug.WriteLine(dt2.ToString())
    '                    'Dim h As HResult = desc.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, etl)
    '                    ''If Not etlptr.Equals(IntPtr.Zero) Then
    '                    ''    etl = Marshal.GetTypedObjectForIUnknown(etlptr, GetType(IPropertyEnumTypeList))
    '                    ''End If
    '                    'Debug.WriteLine(If(etl Is Nothing, "No ", "") & "List")
    '                    Dim etlcount As UInt32
    '                    Dim pet As IPropertyEnumType, dtex As String, pet2 As IPropertyEnumType2, img As String
    '                    'etl.GetCount(etlcount)
    '                    'For a = 0 To etlcount - 1
    '                    '    etl.GetAt(a, GetType(IPropertyEnumType).GUID, pet)
    '                    '    pet.GetDisplayText(dtex)
    '                    '    pet.GetValue(k)

    '                    '    pet2 = pet
    '                    '    Dim h2 = pet2.GetImageReference(img)
    '                    '    Debug.WriteLine(img & h2.ToString())
    '                    '    Debug.WriteLine(dtex)
    '                    '    Debug.WriteLine("val " & getPropertyValue(k, "System.OfflineStatus"))
    '                    'Next
    '                    'Debug.WriteLine("System.StorageProviderCustomStates")
    '                    'Functions.PSGetPropertyDescriptionByName("System.StorageProviderCustomStates", GetType(IPropertyDescription).GUID, desc)
    '                    'desc.GetDisplayType(dt2)
    '                    'Debug.WriteLine(dt2.ToString())

    '                    For x = 0 To pc - 1
    '                        Dim pk2 As PROPERTYKEY
    '                        propertyStore.GetAt(x, pk2)

    '                        propertyStore.GetValue(pk2, k)
    '                        'propertyStore.get
    '                        'Debug.WriteLine(getPropertyValue(k, "unk").ToString())
    '                        Functions.PSGetPropertyDescription(
    '                            pk2,
    '                            GetType(IPropertyDescription).GUID, ptr)
    '                        desc = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IPropertyDescription2))
    '                        Dim cname As String, dispv As String
    '                        desc.GetCanonicalName(cname)
    '                        Debug.WriteLine(cname, getPropertyValue(k, cname).ToString())
    '                        desc.FormatForDisplay(k, PropertyDescriptionFormatOptions.None, dispv)
    '                        Debug.WriteLine(dispv)
    '                        desc.GetDisplayType(dt2)
    '                        Debug.WriteLine(dt2.ToString())
    '                        'If cname = "System.StorageProviderState" Then
    '                        desc.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, etl)
    '                        'If Not etlptr.Equals(IntPtr.Zero) Then
    '                        '    etl = Marshal.GetTypedObjectForIUnknown(etlptr, GetType(IPropertyEnumTypeList))
    '                        'End If
    '                        Debug.WriteLine(If(etl Is Nothing, "No ", "") & "List")
    '                        Dim pd2 As IPropertyDescription2 = desc
    '                        If dt2 = PropertyDisplayType.Enumerated Then
    '                            pd2.GetImageReferenceForValue(k, img)
    '                            Debug.WriteLine("img" & img)
    '                        End If
    '                        etl.GetCount(etlcount)
    '                        For a = 0 To etlcount - 1
    '                            etl.GetAt(a, GetType(IPropertyEnumType).GUID, pet)
    '                            pet.GetDisplayText(dtex)
    '                            pet2 = pet
    '                            pet2.GetImageReference(img)
    '                            Debug.WriteLine(img)
    '                            Debug.WriteLine(dtex)
    '                        Next
    '                        Dim ecount As Integer = 1
    '                        'etl.GetCount(ecount)
    '                        'End If
    '                    Next
    '                    'Dim pss As IPropertyStore, psptr As IntPtr
    '                    'Functions.PSGetPropertyStoreForFile("c:\repos\test.jpg", 0, pss)
    '                    '' Functions.SHLoadPropertyStoreFromBlob(k, pss)
    '                    'pss = Marshal.GetTypedObjectForIUnknown(psptr, GetType(IPropertyStore))
    '                    'CType(desc, IPropertyDescription2).GetImageReferenceForValue(k, imageRes)
    '                    ''ps.gets
    '                    'Me.ShellItem2.GetString(If(Not Me.Parent Is Nothing, Me.Parent, Shell.Desktop).Columns(canonicalName).PROPERTYKEY, imageRes)

    '                    'desc.FormatForDisplay(propVariant, PropertyDescriptionFormatOptions.None, imageRes)
    '                    Dim dt As PropertyDisplayType
    '                    'desc.GetDisplayType(dt)
    '                    'Debug.WriteLine(desc.GetEnumTypeList(GetType(IPropertyEnumTypeList).GUID, etl))
    '                    'Dim count As UInt32
    '                    'etl.GetCount(count)
    '                    'Functions.PSGetImageReferenceForValue(
    '                    '    If(Not Me.Parent Is Nothing, Me.Parent, Shell.Desktop).Columns(canonicalName).PROPERTYKEY,
    '                    '    propVariant, ptr)
    '                    'imageRes = Marshal.PtrToStringUni(ptr)
    '                    Return Nothing
    '                Case Else
    '                    Return String.Format("Column data type {0}/{1} is not implemented.", propVariant.vt, canonicalName)
    '            End Select

    '        Case Else
    '            Return String.Format("Column data type {0}/{1} is not implemented.", propVariant.vt, canonicalName)
    '    End Select
    'End Function
End Class
