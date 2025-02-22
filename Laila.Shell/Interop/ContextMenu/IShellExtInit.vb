Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Interop.ContextMenu
    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
                  GuidAttribute("000214e8-0000-0000-c000-000000000046")>
    Public Interface IShellExtInit
        <PreserveSig()>
        Function Initialize(ByVal pidlFolder As IntPtr,
                            ByVal lpdobj As ComTypes.IDataObject,
                            ByVal hKeyProgID As IntPtr) As Integer    'Treat all HANDLEs as IntPtr
    End Interface
End Namespace
