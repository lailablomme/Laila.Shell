Namespace Helpers
    Public Class Crc32
        Private Shared ReadOnly Table As UInteger() = GenerateTable()

        Public Shared Function Compute(bytes As Byte()) As UInteger
            Dim crc As UInteger = &HFFFFFFFFUI
            For Each b In bytes
                Dim index As Integer = CInt((crc Xor b) And &HFF)
                crc = (crc >> 8) Xor Table(index)
            Next
            Return Not crc
        End Function

        Private Shared Function GenerateTable() As UInteger()
            Dim table(255) As UInteger
            Const poly As UInteger = &HEDB88320UI
            For i = 0 To 255
                Dim crc As UInteger = CUInt(i)
                For j = 0 To 7
                    If (crc And 1) <> 0 Then
                        crc = (crc >> 1) Xor poly
                    Else
                        crc >>= 1
                    End If
                Next
                table(i) = crc
            Next
            Return table
        End Function
    End Class

End Namespace