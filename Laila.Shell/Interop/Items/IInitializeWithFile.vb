Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport, Guid("b7d14566-0509-4cce-a71f-0a554233bd9b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IInitializeWithFile
        ''' <summary>
        ''' Initializes the handler with a file path.
        ''' </summary>
        ''' <param name="pszFilePath">The file path to initialize with.</param>
        ''' <param name="grfMode">The access mode (from the STGM enumeration).</param>
        <PreserveSig>
        Function Initialize(<MarshalAs(UnmanagedType.LPWStr)> pszFilePath As String, grfMode As UInteger) As Integer
    End Interface
End Namespace
