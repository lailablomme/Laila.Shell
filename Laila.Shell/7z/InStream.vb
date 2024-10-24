Imports System.IO
Imports System.Runtime.InteropServices

Public Class InStream
    Implements IInStream

    Private fileStream As FileStream

    ' Constructor that opens the file for reading
    Public Sub New(filePath As String)
        fileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read)
    End Sub

    Public Function Read(<Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal data() As Byte, size As UInteger) As Integer Implements IInStream.Read
        'Dim buffer(size - 1) As Byte
        Return fileStream.Read(data, 0, CInt(size))
        'If bytesRead > 0 Then
        '    Marshal.Copy(buffer, 0, data, bytesRead)
        'End If
        '    Return 0 ' S_OK
    End Function

    Public Sub Seek(offset As Long, origin As SeekOrigin, newpos As IntPtr) Implements IInStream.Seek
        fileStream.Seek(offset, origin)
        If Not IntPtr.Zero.Equals(newpos) Then Marshal.WriteInt64(newpos, fileStream.Position)
    End Sub
End Class
