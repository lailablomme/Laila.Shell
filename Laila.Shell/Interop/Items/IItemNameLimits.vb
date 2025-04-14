Imports System.Runtime.InteropServices

<ComImport>
<Guid("1DF0D7F1-5EBE-4F8F-B86D-AD26CFD9E1D1")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IItemNameLimits
    ' HRESULT GetValidCharacters(
    '     [out] LPWSTR *ppwszValidChars,
    '     [out] LPWSTR *ppwszInvalidChars
    ' );
    <PreserveSig>
    Function GetValidCharacters(
        <Out, MarshalAs(UnmanagedType.LPWStr)> ByRef pwszValidChars As String,
        <Out, MarshalAs(UnmanagedType.LPWStr)> ByRef pwszInvalidChars As String
    ) As Integer

    ' HRESULT GetMaxLength(
    '     [in, optional] LPCWSTR pszName,
    '     [out] DWORD *pcchMaxName
    ' );
    <PreserveSig>
    Function GetMaxLength(
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String,
        <Out> ByRef pcchMaxName As UInteger
    ) As Integer
End Interface
