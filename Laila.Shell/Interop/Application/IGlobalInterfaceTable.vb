Imports System.Runtime.InteropServices

Namespace Interop.Application
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("00000146-0000-0000-C000-000000000046")>
    Public Interface IGlobalInterfaceTable
        ''' <summary>
        ''' Registers an interface in the Global Interface Table (GIT).
        ''' </summary>
        ''' <param name="pUnk">The interface to register.</param>
        ''' <param name="riid">The GUID of the interface.</param>
        ''' <param name="pdwCookie">A cookie that identifies the registered interface.</param>
        Sub RegisterInterfaceInGlobal(
        <MarshalAs(UnmanagedType.Interface)> ByVal pUnk As Object,
        <[In]> ByRef riid As Guid,
        <Out> ByRef pdwCookie As UInteger)

        ''' <summary>
        ''' Revokes an interface from the Global Interface Table (GIT).
        ''' </summary>
        ''' <param name="dwCookie">The cookie of the registered interface to revoke.</param>
        Sub RevokeInterfaceFromGlobal(ByVal dwCookie As UInteger)

        ''' <summary>
        ''' Retrieves an interface from the Global Interface Table (GIT).
        ''' </summary>
        ''' <param name="dwCookie">The cookie of the interface to retrieve.</param>
        ''' <param name="riid">The GUID of the interface to retrieve.</param>
        ''' <param name="ppv">The retrieved interface.</param>
        Sub GetInterfaceFromGlobal(
        ByVal dwCookie As UInteger,
        <[In]> ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)
    End Interface
End Namespace
