Imports System.Runtime.InteropServices

<ComImport>
<Guid("0C9FB851-E5C9-43EB-A370-F0677B13874C")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IOperationsProgressDialog
    <PreserveSig>
    Function StartProgressDialog(hwndOwner As IntPtr, flags As UInteger) As HResult

    <PreserveSig>
    Function StopProgressDialog() As HResult

    <PreserveSig>
    Function SetOperation(operation As UInteger) As HResult

    <PreserveSig>
    Function SetMode(mode As UInteger) As HResult

    <PreserveSig>
    Function UpdateProgress(ullPointsCurrent As ULong, ullPointsTotal As ULong, ullSizeCurrent As ULong, ullSizeTotal As ULong, ullItemsCurrent As ULong, ullItemsTotal As ULong) As HResult

    <PreserveSig>
    Function UpdateLocations(psiSource As IShellItem, psiTarget As IShellItem, psiItem As IShellItem) As HResult

    <PreserveSig>
    Function ResetTimer() As HResult

    <PreserveSig>
    Function PauseTimer() As HResult

    <PreserveSig>
    Function ResumeTimer() As HResult

    <PreserveSig>
    Function GetMilliseconds(ullElapsed As ULong, ullRemaining As ULong) As HResult

    <PreserveSig>
    Function GetOperationStatus() As HResult
End Interface