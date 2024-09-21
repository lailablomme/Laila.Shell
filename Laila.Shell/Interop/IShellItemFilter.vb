Imports System.Runtime.InteropServices

<ComImport>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
<Guid("2659b475-eeb8-48b7-8f07-b378810f48cf")>
Public Interface IShellItemFilter
    Function GetEnumFlagsForItem(<[In], MarshalAs(UnmanagedType.Interface)> ByRef psi As IShellItem, <Out> ByRef pgrfFlags As SHCONTF) As HResult
    Function IncludeItem(<[In], MarshalAs(UnmanagedType.Interface)> ByRef psi As IShellItem) As HResult
End Interface
