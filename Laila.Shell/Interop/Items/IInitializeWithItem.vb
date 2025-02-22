Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport, Guid("7f73be3f-fb79-493c-a6c7-7ee14e245841"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IInitializeWithItem
        ''' <summary>
        ''' Initializes the handler with a Shell item.
        ''' </summary>
        ''' <param name="psi">The Shell item to initialize with.</param>
        ''' <param name="grfMode">The access mode (from the STGM enumeration).</param>
        <PreserveSig>
        Function Initialize(psi As IShellItem, grfMode As UInteger) As Integer
    End Interface
End Namespace
