Imports System.Runtime.InteropServices

<ComImport, Guid("0000000c-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IStream
    Sub Read(<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> pv() As Byte, cb As UInt32, ByRef pcbRead As UInt32)
    Sub Write(<MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> pv() As Byte, cb As UInt32, ByRef pcbWritten As UInt32)
    Sub Seek(dlibMove As Long, dwOrigin As UInt32, ByRef plibNewPosition As Long)
    Sub SetSize(libNewSize As Long)
    Sub CopyTo(pstm As IStream, cb As Long, ByRef pcbRead As Long, ByRef pcbWritten As Long)
    Sub Commit(grfCommitFlags As UInt32)
    Sub Revert()
    Sub LockRegion(libOffset As Long, cb As Long, dwLockType As UInt32)
    Sub UnlockRegion(libOffset As Long, cb As Long, dwLockType As UInt32)
    Sub Stat(ByRef pstatstg As System.Runtime.InteropServices.ComTypes.STATSTG, grfStatFlag As UInt32)
    Sub Clone(ByRef ppstm As IStream)
End Interface
