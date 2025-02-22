Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Interop.DragDrop
    Public Class EnumFORMATETC
        Implements ComTypes.IEnumFORMATETC

        Private formats As FORMATETC()
        Private currentIndex As Integer

        Public Sub New(formats As FORMATETC())
            Me.formats = formats
            Me.currentIndex = 0
        End Sub

        Public Function Skip(celt As Integer) As Integer Implements ComTypes.IEnumFORMATETC.Skip
            currentIndex += celt
            Return If(currentIndex <= formats.Length, 0, 1) ' S_OK or S_FALSE
        End Function

        Public Function Reset() As Integer Implements ComTypes.IEnumFORMATETC.Reset
            Debug.WriteLine("EnumFORMATETC.Reset")
            currentIndex = 0
            Return HRESULT.S_OK
        End Function

        Public Sub Clone(ByRef ppenum As ComTypes.IEnumFORMATETC) Implements ComTypes.IEnumFORMATETC.Clone
            ppenum = New EnumFORMATETC(formats)
        End Sub

        Public Function [Next](celt As Integer, rgelt() As FORMATETC, pceltFetched() As Integer) As Integer Implements ComTypes.IEnumFORMATETC.Next
            Debug.WriteLine("EnumFORMATETC.Next " & celt)
            Dim fetched As Integer = 0
            While currentIndex < formats.Length AndAlso fetched < celt
                rgelt(fetched) = formats(currentIndex)
                currentIndex += 1
                fetched += 1
            End While
            'ReDim Preserve pceltFetched(0)
            If Not pceltFetched Is Nothing Then
                pceltFetched(0) = fetched
            End If
            Return If(fetched = celt, 0, 1) ' S_OK or S_FALSE
        End Function
    End Class
End Namespace
