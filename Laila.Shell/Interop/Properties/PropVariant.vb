Imports System.Runtime.InteropServices

Public Structure PROPVARIANT
    Implements IDisposable

    Public vt As Short

    Public wReserved1 As UInt16

    Public wReserved2 As UInt16

    Public wReserved3 As UInt16

    Public union As PropVariantUnion

    ' Set string value
    Public Sub SetValue(value As String)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Allocate memory for the string and copy the value
        Me.union.ptr = Marshal.StringToBSTR(value)

        ' Set the data type to VT_LPWSTR (wide string)
        Me.vt = VarEnum.VT_BSTR
    End Sub

    ' Set 32-bit integer value (for file size or dates in FILETIME format)
    Public Sub SetValue(value As UInt32)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Set the value
        Me.union.uintVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        Me.vt = VarEnum.VT_UI4
    End Sub

    ' Set 64-bit integer value (for file size or dates in FILETIME format)
    Public Sub SetValue(value As Long)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Set the value
        Me.union.hVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        Me.vt = VarEnum.VT_I8
    End Sub

    ' Set 64-bit integer value (for file size or dates in FILETIME format)
    Public Sub SetValue(value As ULong)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Set the value
        Me.union.uhVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        Me.vt = VarEnum.VT_UI8
    End Sub

    ' Set boolean value
    Public Sub SetValue(value As Boolean)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Set the boolean value
        Me.union.boolVal = If(value, -1, 0) ' TRUE is -1, FALSE is 0

        ' Set the data type to VT_BOOL (boolean)
        Me.vt = VarEnum.VT_BOOL
    End Sub

    ' Set DateTime value (converts to FILETIME)
    Public Sub SetValue(value As DateTime)
        ' Clear previous data
        Functions.PropVariantClear(Me)

        ' Convert DateTime to FILETIME (100-nanosecond intervals since 1601)
        Dim fileTime As Long = value.ToFileTimeUtc()

        ' Set the FILETIME value (which is a 64-bit integer)
        Me.union.hVal = fileTime

        ' Set the data type to VT_FILETIME
        Me.vt = VarEnum.VT_FILETIME
    End Sub

    Public Function GetValue() As Object
        Dim ve As VarEnum = Me.vt
        If ve.HasFlag(VarEnum.VT_VECTOR) Then
            ve = ve Xor VarEnum.VT_VECTOR
            Select Case ve
                Case VarEnum.VT_UNKNOWN
                    Dim count As Integer = Me.union.bstrblobVal.cbSize
                    Dim resultList As New List(Of IntPtr)()
                    Dim pArray As IntPtr = Me.union.bstrblobVal.pData
                    For i As Integer = 0 To count - 1
                        Dim pUnknown As IntPtr = Marshal.ReadIntPtr(pArray, i * IntPtr.Size)
                        resultList.Add(pUnknown)
                    Next
                    Return resultList
                Case VarEnum.VT_LPWSTR
                    Dim count As Integer = Me.union.bstrblobVal.cbSize
                    Dim resultList As New List(Of String)()
                    Dim pData As IntPtr = Me.union.bstrblobVal.pData
                    For i As Integer = 0 To count - 1
                        Dim ptr As IntPtr = Marshal.ReadIntPtr(pData, i * IntPtr.Size)
                        Dim str As String = Marshal.PtrToStringUni(ptr)
                        resultList.Add(str)
                    Next
                    Return resultList
                Case VarEnum.VT_UI1
                    Dim count As Integer = Me.union.bstrblobVal.cbSize
                    Dim byteArray(count - 1) As Byte
                    Marshal.Copy(Me.union.bstrblobVal.pData, byteArray, 0, count)
                    Return byteArray
            End Select
        Else
            Select Case ve
                Case VarEnum.VT_EMPTY
                    Return Nothing
                Case VarEnum.VT_BSTR
                    Return Marshal.PtrToStringBSTR(Me.union.ptr)
                Case VarEnum.VT_LPWSTR
                    Return Marshal.PtrToStringUni(Me.union.ptr)
                Case VarEnum.VT_FILETIME
                    Return DateTime.FromFileTime(Me.union.hVal)
                Case VarEnum.VT_I1
                    Return Chr(Me.union.bVal)
                Case VarEnum.VT_I2
                    Return Me.union.iVal
                Case VarEnum.VT_UI2
                    Return Me.union.uiVal
                Case VarEnum.VT_I4
                    Return Me.union.intVal
                Case VarEnum.VT_UI4
                    Return Me.union.uintVal
                Case VarEnum.VT_UI8
                    Return Me.union.uhVal
                Case Else
                    Debug.WriteLine("Unknown " & ve.ToString())
                    Return Nothing
            End Select
        End If
        Return Nothing
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        Functions.PropVariantClear(Me)
    End Sub
End Structure
