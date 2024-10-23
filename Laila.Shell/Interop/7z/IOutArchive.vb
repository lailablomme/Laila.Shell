Imports System.Runtime.InteropServices

Namespace SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000600A00000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IOutArchive

        ' Updates archive items
        <PreserveSig>
        Function UpdateItems(
        <MarshalAs(UnmanagedType.Interface)> stream As ISequentialOutStream,
        <MarshalAs(UnmanagedType.U4)> count As UInteger,
        <MarshalAs(UnmanagedType.Interface)> callback As IArchiveUpdateCallback
    ) As Integer

        ' Gets file time type
        Sub GetFileTimeType(type As IntPtr)

    End Interface
End Namespace