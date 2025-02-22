' File input stream for compression
Imports Laila.Shell.Interop.SevenZip
Imports Laila.Shell.SevenZip
Imports System.IO
Imports System.Runtime.InteropServices

Namespace SevenZip
    Public Class SuquentialInStream
        Implements ISequentialInStream

        Private fileStream As FileStream

        ' Constructor that opens the file for reading
        Public Sub New(filePath As String)
            fileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read)
        End Sub

        ' Implement the Read method to read from the file
        Public Function Read(data As IntPtr, size As UInt32, ByRef processedSize As UInt32) As Integer Implements ISequentialInStream.Read
            Dim buffer(size - 1) As Byte
            Dim bytesRead As Integer = fileStream.Read(buffer, 0, CInt(size))
            If bytesRead > 0 Then
                Marshal.Copy(buffer, 0, data, bytesRead)
            End If
            processedSize = CUInt(bytesRead)
            Return 0 ' S_OK
        End Function

        ' Ensure to close the stream when done
        Public Sub Close()
            fileStream.Close()
        End Sub
    End Class
End Namespace