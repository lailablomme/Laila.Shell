Imports System.Runtime.InteropServices

<ComImportAttribute(), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown), Guid("e318ad57-0aa0-450f-aca5-6fab7103d917")>
Public Interface IPersistSerializedPropStorage
    Function SetFlags(
  <[In]> flags As UInt32
) As HRESULT
    Function SetPropertyStorage(<[In]> psps As IntPtr, <[In]> cb As UInt32) As HRESULT
End Interface