Imports System.Runtime.InteropServices

Public Structure PROPVARIANT
    Implements IDisposable

    Public vt As Short

    Public wReserved1 As UInt16

    Public wReserved2 As UInt16

    Public wReserved3 As UInt16

    Public union As PropVariantUnion

    ' Set string value
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As String)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Allocate memory for the string and copy the value
        propVar.union.ptr = Marshal.StringToBSTR(value)

        ' Set the data type to VT_LPWSTR (wide string)
        propVar.vt = VarEnum.VT_BSTR
    End Sub

    ' Set 32-bit integer value (for file size or dates in FILETIME format)
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As UInt32)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Set the value
        propVar.union.uintVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        propVar.vt = VarEnum.VT_UI4
    End Sub

    ' Set 64-bit integer value (for file size or dates in FILETIME format)
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As Long)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Set the value
        propVar.union.hVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        propVar.vt = VarEnum.VT_I8
    End Sub

    ' Set 64-bit integer value (for file size or dates in FILETIME format)
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As ULong)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Set the value
        propVar.union.uhVal = value

        ' Set the data type to VT_I8 (64-bit integer)
        propVar.vt = VarEnum.VT_UI8
    End Sub

    ' Set boolean value
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As Boolean)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Set the boolean value
        propVar.union.boolVal = If(value, -1, 0) ' TRUE is -1, FALSE is 0

        ' Set the data type to VT_BOOL (boolean)
        propVar.vt = VarEnum.VT_BOOL
    End Sub

    ' Set DateTime value (converts to FILETIME)
    Public Shared Sub SetValue(ByRef propVar As PROPVARIANT, value As DateTime)
        ' Clear previous data
        Functions.PropVariantClear(propVar)

        ' Convert DateTime to FILETIME (100-nanosecond intervals since 1601)
        Dim fileTime As Long = value.ToFileTimeUtc()

        ' Set the FILETIME value (which is a 64-bit integer)
        propVar.union.hVal = fileTime

        ' Set the data type to VT_FILETIME
        propVar.vt = VarEnum.VT_FILETIME
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Functions.PropVariantClear(Me)
    End Sub
End Structure
