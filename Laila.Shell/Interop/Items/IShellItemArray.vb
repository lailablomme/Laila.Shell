Imports System.Runtime.InteropServices

<ComImport>
<Guid("b63ea76d-1f85-456f-a19c-48159efa858b")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IShellItemArray
    Sub BindToHandler(<[In]> pbc As IntPtr, <[In], MarshalAs(UnmanagedType.LPStruct)> ByRef bhid As Guid, <[In], MarshalAs(UnmanagedType.LPStruct)> ByRef riid As Guid, <Out> ByRef ppvOut As IntPtr)
    Sub GetPropertyStore(<[In]> flags As Integer, <[In], MarshalAs(UnmanagedType.LPStruct)> ByRef riid As Guid, <Out> ByRef ppv As IntPtr)
    Sub GetPropertyDescriptionList(<[In], MarshalAs(UnmanagedType.LPStruct)> ByRef keyType As Guid, <[In], MarshalAs(UnmanagedType.LPStruct)> ByRef riid As Guid, <Out> ByRef ppv As IntPtr)
    Sub GetAttributes(<[In]> AttribFlags As Integer, <[Out]> ByRef sfgao As Integer)
    Sub GetCount(<Out> ByRef pdwNumItems As Integer)
    Sub GetItemAt(<[In]> dwIndex As Integer, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
    Sub EnumItems(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppenumShellItems As IEnumShellItems)
End Interface