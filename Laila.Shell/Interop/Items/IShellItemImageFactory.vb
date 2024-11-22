Imports System.Drawing
Imports System.Runtime.InteropServices

<ComImport, Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IShellItemImageFactory
    <PreserveSig>
    Function GetImage(<[In]> size As Size, <[In]> flags As SIIGBF, <Out> ByRef phbm As IntPtr) As HRESULT
End Interface
