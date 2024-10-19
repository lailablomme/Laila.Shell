Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Helpers
    Public Class StreamAdapter
        Implements IStream

        Private _filePath As String
        Private _fileStream As FileStream

        Public Sub New(filePath As String)
            _filePath = filePath
        End Sub

        Public Sub Read(pv() As Byte, cb As Integer, pcbRead As IntPtr) Implements IStream.Read
            If _fileStream Is Nothing Then
                _fileStream = New FileStream(_filePath, FileMode.Open, FileAccess.Read)
                _fileStream.Seek(0, SeekOrigin.End)
            End If

            If _fileStream.Position + cb > _fileStream.Length Then
                cb = _fileStream.Length - _fileStream.Position
            End If

            Dim bytesRead As Integer
            If cb > 0 Then
                bytesRead = _fileStream.Read(pv, 0, cb)
            End If

            If bytesRead = 0 Then
                _fileStream.Dispose()
                _fileStream = Nothing
            End If

            If pcbRead <> IntPtr.Zero Then
                Marshal.WriteInt32(pcbRead, bytesRead)
            End If
        End Sub

        Public Sub Write(pv() As Byte, cb As Integer, pcbWritten As IntPtr) Implements IStream.Write
            Throw New NotSupportedException()
        End Sub

        Public Sub Seek(dlibMove As Long, dwOrigin As Integer, plibNewPosition As IntPtr) Implements IStream.Seek
            If _fileStream Is Nothing Then
                _fileStream = New FileStream(_filePath, FileMode.Open, FileAccess.Read)
                _fileStream.Seek(0, SeekOrigin.End)
            End If

            Select Case dwOrigin
                Case 0 ' STREAM_SEEK_SET
                    _fileStream.Seek(dlibMove, SeekOrigin.Begin)
                Case 1 ' STREAM_SEEK_CUR
                    _fileStream.Seek(dlibMove, SeekOrigin.Current)
                Case 2 ' STREAM_SEEK_END
                    _fileStream.Seek(dlibMove, SeekOrigin.End)
                Case Else
                    Throw New NotSupportedException()
            End Select

            If plibNewPosition <> IntPtr.Zero Then
                Marshal.WriteInt64(plibNewPosition, _fileStream.Position)
            End If

            If _fileStream.Position = _fileStream.Length Then
                _fileStream.Dispose()
                _fileStream = Nothing
            End If
        End Sub

        Public Sub SetSize(libNewSize As Long) Implements IStream.SetSize
            Throw New NotSupportedException()
        End Sub
        Public Sub CopyTo(pstm As IStream, cb As Long, pcbRead As IntPtr, pcbWritten As IntPtr) Implements IStream.CopyTo
            If _fileStream Is Nothing Then
                _fileStream = New FileStream(_filePath, FileMode.Open, FileAccess.Read)
                _fileStream.Seek(0, SeekOrigin.End)
            End If

            If cb = -1 Then cb = Long.MaxValue
            Dim endPosition As Long = Math.Min(_fileStream.Position + cb, _fileStream.Length)
            Dim bytesRead As UInt64 = 0
            Dim bytesWritten As UInt64 = 0
            Dim ptrWritten As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of Integer))
            While _fileStream.Position <> endPosition
                cb = Math.Min(endPosition - _fileStream.Position, 1024 * 1000)
                Dim pv(cb - 1) As Byte
                bytesRead += _fileStream.Read(pv, 0, cb)
                pstm.Write(pv, cb, ptrWritten)
                bytesWritten += Marshal.ReadInt32(ptrWritten)
            End While
            pstm.Commit(0)
            Marshal.FinalReleaseComObject(pstm)

            If Not IntPtr.Zero.Equals(pcbRead) Then
                Marshal.WriteInt64(pcbRead, bytesRead)
            End If
            If Not IntPtr.Zero.Equals(pcbWritten) Then
                Marshal.WriteInt64(pcbWritten, bytesWritten)
            End If

            _fileStream.Dispose()
            _fileStream = Nothing
        End Sub

        Public Sub Commit(grfCommitFlags As Integer) Implements IStream.Commit
            Throw New NotSupportedException()
        End Sub

        Public Sub Revert() Implements IStream.Revert
            Throw New NotSupportedException()
        End Sub

        Public Sub LockRegion(libOffset As Long, cb As Long, dwLockType As Integer) Implements IStream.LockRegion
            Throw New NotSupportedException()
        End Sub

        Public Sub UnlockRegion(libOffset As Long, cb As Long, dwLockType As Integer) Implements IStream.UnlockRegion
            Throw New NotSupportedException()
        End Sub

        Public Sub Stat(ByRef pstatstg As STATSTG, grfStatFlag As Integer) Implements IStream.Stat
            If _fileStream Is Nothing Then
                _fileStream = New FileStream(_filePath, FileMode.Open, FileAccess.Read)
                _fileStream.Seek(0, SeekOrigin.End)
            End If

            ' Provide information about the stream
            pstatstg.cbSize = _fileStream.Length
            pstatstg.type = 2 ' STGTY_STREAM (Stream)
        End Sub

        Public Sub Clone(ByRef ppstm As IStream) Implements IStream.Clone
            Throw New NotSupportedException()
        End Sub
    End Class
End Namespace