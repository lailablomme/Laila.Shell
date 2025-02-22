Imports System.Runtime.InteropServices

Namespace Interop.SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000600030000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ISetProperties
        <PreserveSig>
        Function SetProperties(
            <[In], MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPWStr, SizeParamIndex:=2)> names As String(),
            values As IntPtr,
            numProperties As UInteger
        ) As Integer
    End Interface
End Namespace
