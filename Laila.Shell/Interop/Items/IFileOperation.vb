Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport>
<Guid("947AAB5F-0A5C-4C13-B4D6-4BF7836FC9F8")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IFileOperation
    <PreserveSig>
    Function Advise(<[In], MarshalAs(UnmanagedType.Interface)> pfops As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function Unadvise(dwCookie As UInteger) As HRESULT

    <PreserveSig>
    Function SetOperationFlags(dwOperationFlags As UInteger) As HRESULT

    <PreserveSig>
    Function SetProgressMessage(<[In], MarshalAs(UnmanagedType.LPWStr)> pszMessage As String) As HRESULT

    <PreserveSig>
    Function SetProgressDialog(<[In], MarshalAs(UnmanagedType.Interface)> popd As IOperationsProgressDialog) As HRESULT

    <PreserveSig>
    Function SetProperties(<[In]> pproparray As IntPtr) As HRESULT

    <PreserveSig>
    Function SetOwnerWindow(hwndOwner As IntPtr) As HRESULT

    <PreserveSig>
    Function ApplyPropertiesToItem(<[In], MarshalAs(UnmanagedType.Interface)> psiItem As IShellItem) As HRESULT

    <PreserveSig>
    Function ApplyPropertiesToItems(<[In], MarshalAs(UnmanagedType.Interface)> punkItems As IUnknown) As HRESULT

    <PreserveSig>
    Function RenameItem(<[In], MarshalAs(UnmanagedType.Interface)> psiItem As IShellItem, <[In], MarshalAs(UnmanagedType.LPWStr)> pszNewName As String, <[In], MarshalAs(UnmanagedType.Interface)> pfopsItem As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function RenameItems(<[In], MarshalAs(UnmanagedType.Interface)> pUnkItems As IUnknown, <[In], MarshalAs(UnmanagedType.LPWStr)> pszNewName As String) As HRESULT

    <PreserveSig>
    Function MoveItem(<[In], MarshalAs(UnmanagedType.Interface)> psiItem As IShellItem, <[In], MarshalAs(UnmanagedType.Interface)> psiDestinationFolder As IShellItem, <[In], MarshalAs(UnmanagedType.LPWStr)> pszNewName As String, <[In], MarshalAs(UnmanagedType.Interface)> pfopsItem As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function MoveItems(<[In], MarshalAs(UnmanagedType.Interface)> punkItems As IShellItemArray, <[In], MarshalAs(UnmanagedType.Interface)> psiDestinationFolder As IShellItem) As HRESULT

    <PreserveSig>
    Function CopyItem(<[In], MarshalAs(UnmanagedType.Interface)> psiItem As IShellItem, <[In], MarshalAs(UnmanagedType.Interface)> psiDestinationFolder As IShellItem, <[In], MarshalAs(UnmanagedType.LPWStr)> pszCopyName As String, <[In], MarshalAs(UnmanagedType.Interface)> pfopsItem As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function CopyItems(<[In], MarshalAs(UnmanagedType.Interface)> punkItems As IShellItemArray, <[In], MarshalAs(UnmanagedType.Interface)> psiDestinationFolder As IShellItem) As HRESULT

    <PreserveSig>
    Function DeleteItem(<[In], MarshalAs(UnmanagedType.Interface)> psiItem As IShellItem, <[In], MarshalAs(UnmanagedType.Interface)> pfopsItem As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function DeleteItems(<[In], MarshalAs(UnmanagedType.Interface)> punkItems As IUnknown) As HRESULT

    <PreserveSig>
    Function NewItem(<[In], MarshalAs(UnmanagedType.Interface)> psiDestinationFolder As IShellItem, dwFileAttributes As UInteger, <[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String, <[In], MarshalAs(UnmanagedType.LPWStr)> pszTemplateName As String, <[In], MarshalAs(UnmanagedType.Interface)> pfopsItem As IFileOperationProgressSink) As HRESULT

    <PreserveSig>
    Function PerformOperations() As HRESULT

    <PreserveSig>
    Function GetAnyOperationsAborted(<Out, MarshalAs(UnmanagedType.Bool)> ByRef pfAnyOperationsAborted As Boolean) As HRESULT
End Interface