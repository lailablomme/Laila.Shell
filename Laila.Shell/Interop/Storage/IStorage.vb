Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Interop.Storage
    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")>
    Public Interface IStorage
        Function CreateStream(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        grfMode As Integer,
        reserved1 As Integer,
        reserved2 As Integer,
        <Out> ByRef ppstm As IntPtr) As HRESULT

        Sub OpenStream(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        reserved1 As IntPtr,
        grfMode As Integer,
        reserved2 As Integer,
        <Out> ByRef ppstm As IStream)

        Sub CreateStorage(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        grfMode As Integer,
        reserved1 As Integer,
        reserved2 As Integer,
        <Out> ByRef ppstg As IStorage)

        Sub OpenStorage(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        pstgPriority As IStorage,
        grfMode As Integer,
        snbExclude As IntPtr,
        reserved As Integer,
        <Out> ByRef ppstg As IStorage)

        Sub CopyTo(
        ciidExclude As Integer,
        rgiidExclude As Guid(),
        snbExclude As IntPtr,
        pstgDest As IStorage)

        Sub MoveElementTo(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        pstgDest As IStorage,
        pwcsNewName As String,
        grfFlags As Integer)

        Sub Commit(grfCommitFlags As Integer)

        Sub Revert()

        Sub EnumElements(
        reserved1 As Integer,
        reserved2 As IntPtr,
        reserved3 As Integer,
        <Out> ByRef ppenum As IEnumSTATSTG)

        Sub DestroyElement(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String)

        Sub RenameElement(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsOldName As String,
        <MarshalAs(UnmanagedType.LPWStr)> pwcsNewName As String)

        Sub SetElementTimes(
        <MarshalAs(UnmanagedType.LPWStr)> pwcsName As String,
        <[In]> ByRef pctime As FILETIME,
        <[In]> ByRef patime As FILETIME,
        <[In]> ByRef pmtime As FILETIME)

        Sub SetClass(ByRef clsid As Guid)

        Sub SetStateBits(
        grfStateBits As Integer,
        grfMask As Integer)

        Sub Stat(
        <Out> ByRef pstatstg As STATSTG,
        grfStatFlag As Integer)
    End Interface
End Namespace
