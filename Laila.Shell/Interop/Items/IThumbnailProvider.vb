Imports System.Runtime.InteropServices

<ComImport>
<Guid("e357fccd-a995-4576-b01f-234630154e96")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IThumbnailProvider
    <PreserveSig>
    Function GetThumbnail(
            <[In]> cx As UInteger,
            <Out> ByRef hBitmap As IntPtr,
            <Out> ByRef pdwAlpha As WTS_ALPHATYPE) As Integer
End Interface
Public Enum WTS_ALPHATYPE
    WTSAT_UNKNOWN = 0
    WTSAT_RGB = 1
    WTSAT_ARGB = 2
End Enum