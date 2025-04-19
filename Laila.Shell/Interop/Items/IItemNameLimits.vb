Imports System.Runtime.InteropServices

<ComImport>
<Guid("1DF0D7F1-B267-4D28-8B10-12E23202A5C4")>
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
