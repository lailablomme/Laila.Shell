Imports System.Runtime.InteropServices

<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e357fccd-a995-4576-b01f-234630154e96")>
Public Interface IThumbnailProvider
    Sub GetThumbnail(cx As UInteger, ByRef hBitmap As IntPtr, ByRef bitmapType As WTS_ALPHATYPE)
End Interface

Public Enum WTS_ALPHATYPE
    WTSAT_UNKNOWN = 0
    WTSAT_RGB = 1
    WTSAT_ARGB = 2
End Enum