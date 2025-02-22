Imports System.Runtime.InteropServices

Namespace Interop.Properties
    <ComImport>
    <Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPropertySet
        ' Retrieves a property value from the property set.
        <PreserveSig>
        Function GetValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String, <MarshalAs(UnmanagedType.Struct)> ByRef pPropValue As Object) As Integer

        ' Sets a property value in the property set.
        <PreserveSig>
        Function SetValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String, <MarshalAs(UnmanagedType.Struct)> ByVal pPropValue As Object) As Integer

        ' Removes a property from the property set.
        <PreserveSig>
        Function RemoveValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String) As Integer

        ' Commits all changes to the property set.
        <PreserveSig>
        Function Commit() As Integer

        ' Reverts all changes to the property set.
        <PreserveSig>
        Function Revert() As Integer

        ' Initializes the property set from a given data source.
        <PreserveSig>
        Function Initialize(<MarshalAs(UnmanagedType.Interface)> ByVal pDataSource As Object) As Integer
    End Interface
End Namespace

