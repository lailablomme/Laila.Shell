Imports System.Runtime.InteropServices

Public Class PropertyBagFactory
    <DllImport("ole32.dll", SetLastError:=True)>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef iid As Guid,
        ByRef ppv As IntPtr
    ) As Integer
    End Function

    Public Shared Function CreatePropertyBag() As IPropertyBag
        Dim CLSID_PropertyBag As New Guid("55272A00-42CB-11CE-8135-00AA004BB851")
        Dim IID_IPropertyBag As New Guid("55272A00-42CB-11CE-8135-00AA004BB851")
        Dim pPropertyBag As IntPtr = IntPtr.Zero

        Dim hr As Integer = CoCreateInstance(CLSID_PropertyBag, IntPtr.Zero, 1, IID_IPropertyBag, pPropertyBag)
        If hr <> 0 Then
            Marshal.ThrowExceptionForHR(hr)
        End If

        Return CType(Marshal.GetObjectForIUnknown(pPropertyBag), IPropertyBag)
    End Function
End Class