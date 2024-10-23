' Implement ISequentialOutStream to write to a file
Imports Laila.Shell.SevenZip
Imports System.IO
Imports System.Runtime.InteropServices

Namespace SevenZip
    Public Class FileOutStream
        Implements ISequentialOutStream

        Private fileStream As FileStream

        ' Constructor that accepts the output file path
        Public Sub New(outputFilePath As String)
            ' Open the file stream for writing
            fileStream = New FileStream(outputFilePath, FileMode.Create, FileAccess.Write)
        End Sub

        ' Implement the Write method of ISequentialOutStream
        Public Function Write(data As IntPtr, size As UInt32, processedSize As IntPtr) As Integer Implements ISequentialOutStream.Write
            Try
                ' Create a byte array to hold the data
                Dim buffer(size - 1) As Byte
                ' Copy the data from the unmanaged memory to the managed array
                Marshal.Copy(data, buffer, 0, size)

                ' Write the data to the file
                fileStream.Write(buffer, 0, CInt(size))

                ' Set the number of processed bytes
                Marshal.WriteInt32(processedSize, size)

                Return 0 ' S_OK (success)
            Catch ex As Exception
                Return 1 ' Return an error code on failure
            End Try
        End Function

        ' Make sure to close the stream when done
        Public Sub Close()
            fileStream.Close()
        End Sub
    End Class
End Namespace