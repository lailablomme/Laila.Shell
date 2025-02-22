Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport>
    <Guid("0C9FB851-E5C9-43EB-A370-F0677B13874C")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IOperationsProgressDialog
        <PreserveSig>
        Function StartProgressDialog(hwndOwner As IntPtr, flags As UInteger) As HRESULT

        <PreserveSig>
        Function StopProgressDialog() As HRESULT

        <PreserveSig>
        Function SetOperation(operation As UInteger) As HRESULT

        <PreserveSig>
        Function SetMode(mode As UInteger) As HRESULT

        <PreserveSig>
        Function UpdateProgress(ullPointsCurrent As ULong, ullPointsTotal As ULong, ullSizeCurrent As ULong, ullSizeTotal As ULong, ullItemsCurrent As ULong, ullItemsTotal As ULong) As HRESULT

        <PreserveSig>
        Function UpdateLocations(psiSource As IShellItem, psiTarget As IShellItem, psiItem As IShellItem) As HRESULT

        <PreserveSig>
        Function ResetTimer() As HRESULT

        <PreserveSig>
        Function PauseTimer() As HRESULT

        <PreserveSig>
        Function ResumeTimer() As HRESULT

        <PreserveSig>
        Function GetMilliseconds(ullElapsed As ULong, ullRemaining As ULong) As HRESULT

        <PreserveSig>
        Function GetOperationStatus() As HRESULT
    End Interface
End Namespace
