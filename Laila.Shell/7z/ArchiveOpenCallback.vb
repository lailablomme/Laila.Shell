' NEEDED FOR 7Z ARCHIVE CREATION -- TEMP DISABLED
'Imports System.Runtime.InteropServices
'Imports Laila.Shell.Interop.SevenZip

'Namespace SevenZip
'    Public Class ArchiveOpenCallback
'        Implements IArchiveOpenCallback

'        Public Function SetTotal(count As IntPtr, bytes As IntPtr) As Integer Implements IArchiveOpenCallback.SetTotal
'            Dim c As UInt32 = Marshal.ReadInt32(count)
'            Dim b As UInt64
'            If Not IntPtr.Zero.Equals(bytes) Then b = Marshal.ReadInt64(bytes)
'            Debug.WriteLine("ArchiveOpenCallback.SetTotal  count=" & c & "   bytes=" & b)
'        End Function

'        Public Function SetCompleted(count As IntPtr, bytes As IntPtr) As Integer Implements IArchiveOpenCallback.SetCompleted
'            Dim c As UInt32 = Marshal.ReadInt32(count)
'            Dim b As UInt64 = Marshal.ReadInt64(bytes)
'            Debug.WriteLine("ArchiveOpenCallback.SetCompleted  count=" & c & "   bytes=" & b)
'        End Function
'    End Class
'End Namespace