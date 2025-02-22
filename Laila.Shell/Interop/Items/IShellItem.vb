Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport()>
    <Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellItem
        Sub BindToHandler(ByVal pbc As IntPtr,
        <MarshalAs(UnmanagedType.LPStruct)> ByVal bhid As Guid,
        <MarshalAs(UnmanagedType.LPStruct)> ByVal riid As Guid,
        ByRef ppv As IntPtr)

        Sub GetParent(ByRef ppsi As IShellItem)

        Sub GetDisplayName(ByVal sigdnName As SIGDN, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String)

        Sub GetAttributes(ByVal sfgaoMask As UInt32, ByRef psfgaoAttribs As UInt32)

        Sub Compare(ByVal psi As IShellItem, ByVal hint As UInt32, ByRef piOrder As Integer)
    End Interface
End Namespace
