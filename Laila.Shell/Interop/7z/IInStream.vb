Imports System.IO
Imports System.Runtime.InteropServices

<ComImport>
<Guid("23170F69-40C1-278A-0000-000300030000")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IInStream
    Function Read(<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal data() As Byte, ByVal size As UInteger) As Integer
    Sub Seek(ByVal offset As Long, ByVal origin As SeekOrigin, ByVal newpos As IntPtr)
End Interface