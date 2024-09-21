Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport>
<Guid("7E9FB0D3-919F-4307-AB2E-9B1860310C93")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IShellItem22
    ' Inherits IShellItem

    ' IShellItem methods
    Sub BindToHandler(ByVal pbc As IBindCtx, ByRef bhid As Guid, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    Sub GetParent(<MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
    Sub GetDisplayName(ByVal sigdnName As SIGDN, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String)
    Sub GetAttributes(ByVal sfgaoMask As UInteger, ByRef psfgaoAttribs As UInteger)
    Sub Compare(ByVal psi As IShellItem, ByVal hint As UInteger, ByRef piOrder As Integer)

    ' IShellItem2 methods
    Sub GetPropertyStore(ByVal flags As GETPROPERTYSTOREFLAGS, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    Sub GetPropertyStoreWithCreateObject(ByVal flags As GETPROPERTYSTOREFLAGS, <MarshalAs(UnmanagedType.Interface)> ByVal punkCreateObject As Object, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    Sub GetPropertyStoreForKeys(<MarshalAs(UnmanagedType.LPArray)> ByVal rgKeys As PROPERTYKEY(), ByVal cKeys As UInteger, ByVal flags As GETPROPERTYSTOREFLAGS, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    Sub GetPropertyDescriptionList(ByRef keyType As PROPERTYKEY, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    Sub Update()
    Sub GetProperty(ByRef key As PROPERTYKEY, ByRef pv As PROPVARIANT)
    Sub GetCLSID(ByRef key As PROPERTYKEY, ByRef pclsid As Guid)
    Sub GetFileTime(ByRef key As PROPERTYKEY, ByRef pft As FILETIME)
    Sub GetInt32(ByRef key As PROPERTYKEY, ByRef pi As Integer)
    Sub GetString(ByRef key As PROPERTYKEY, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppsz As String)
    Sub GetUInt32(ByRef key As PROPERTYKEY, ByRef pui As UInteger)
    Sub GetUInt64(ByRef key As PROPERTYKEY, ByRef pull As ULong)
    Sub GetBool(ByRef key As PROPERTYKEY, ByRef pf As Integer)
End Interface

'0E2498