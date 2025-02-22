Imports System.Windows.Input
Imports Laila.Shell.Interop

Namespace Helpers
    Public Class KeyboardHelper
        Public Shared Function KeyToChar(key As Key) As Char?
            Dim virtualKey As Integer = KeyInterop.VirtualKeyFromKey(key)
            Dim keyboardState(255) As Byte

            If Not Functions.GetKeyboardState(keyboardState) Then
                Return Nothing
            End If

            Dim scanCode As UInteger = Functions.MapVirtualKey(CUInt(virtualKey), 0)

            Dim buffer As String = " "
            Dim result As Integer = Functions.ToUnicode(CUInt(virtualKey), scanCode, keyboardState, buffer, buffer.Length, 0)

            If result > 0 Then
                Return buffer(0)
            Else
                Return Nothing
            End If
        End Function
    End Class
End Namespace