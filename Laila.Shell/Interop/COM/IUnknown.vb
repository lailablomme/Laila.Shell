Imports System.Runtime.InteropServices

Namespace Interop.COM
    <ComImport(), Guid("00000000-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IUnknown
        <PreserveSig>
        Function QueryInterface(ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer

        <PreserveSig>
        Function AddRef() As Integer

        <PreserveSig>
        Function Release() As Integer
    End Interface
End Namespace
