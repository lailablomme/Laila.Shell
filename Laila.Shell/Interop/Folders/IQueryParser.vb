Imports System.Runtime.InteropServices
Imports System.Text

<ComImport>
<Guid("2ebdee67-3505-43f8-9946-ea44abc8e5b0")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IQueryParser
    Sub Parse(ByVal pszInputString As String, <MarshalAs(UnmanagedType.Interface)> ByVal pCustomProperties As IEnumUnknown, <MarshalAs(UnmanagedType.Interface)> ByRef ppQueryNode As IQuerySolution)

    Sub SetOption(ByVal opt As Integer, ByVal pOption As IntPtr)

    Sub GetOption(ByVal opt As Integer, ByRef pOption As IntPtr)

    Sub SetMultiOption(ByVal option1 As Integer, ByVal option2 As Integer, ByVal pOption As IntPtr)

    Sub GetMultiOption(ByVal option1 As Integer, ByVal option2 As Integer, ByRef pOption As IntPtr)

    Sub SetLocale(ByVal lcid As UInteger)

    Sub SetSyntax(ByVal pSyntax As Guid)

    Sub SetOptionFromANSIString(ByVal pszOptionString As String, ByVal pszOptionValue As String)

    Sub GetOptionAsANSIString(ByVal pszOptionString As String, ByVal pszOptionValue As String, ByVal cchValue As Integer)

    Sub SetOptionFromUnicodeString(ByVal pszOptionString As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal pszOptionValue As String)

    Sub GetOptionAsUnicodeString(ByVal pszOptionString As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal pszOptionValue As String, ByVal cchValue As Integer)

    Sub ParsePropertyValue(ByVal pszProperty As String, ByVal pszValue As String, ByRef pCondition As ICondition, ByRef pIsList As Integer)
End Interface

