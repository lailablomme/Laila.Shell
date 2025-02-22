Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.DragDrop
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("00000122-0000-0000-C000-000000000046")>
    Public Interface IDropTarget
        <PreserveSig>
        Function DragEnter(
        ByVal pDataObj As ComTypes.IDataObject,
        ByVal grfKeyState As MK,
        ByVal pt As WIN32POINT,
        ByRef pdwEffect As Integer) As Integer

        <PreserveSig>
        Function DragOver(
        ByVal grfKeyState As MK,
        ByVal pt As WIN32POINT,
        ByRef pdwEffect As Integer) As Integer

        <PreserveSig>
        Function DragLeave() As Integer

        <PreserveSig>
        Function Drop(
        ByVal pDataObj As ComTypes.IDataObject,
        ByVal grfKeyState As MK,
        ByVal pt As WIN32POINT,
        ByRef pdwEffect As Integer) As Integer
    End Interface
End Namespace
