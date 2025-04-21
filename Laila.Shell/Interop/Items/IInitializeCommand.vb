Imports Laila.Shell.Interop.Properties
Imports System.Runtime.InteropServices

Namespace Interop.Items

    <ComImport>
    <Guid("85075ACF-231F-40EA-9610-D26B7B58F638")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IInitializeCommand
        <PreserveSig>
        Function Initialize(
            <MarshalAs(UnmanagedType.LPWStr)> pszCommandName As String,
            <MarshalAs(UnmanagedType.Interface)> ppb As IPropertyBag
        ) As Integer
    End Interface
End Namespace