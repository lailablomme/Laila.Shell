Imports System.Runtime.InteropServices

Namespace SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000300010000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ISequentialInStream
        <PreserveSig>
        Function Read(data As IntPtr, size As UInt32, ByRef processedSize As UInt32) As Integer
    End Interface
End Namespace