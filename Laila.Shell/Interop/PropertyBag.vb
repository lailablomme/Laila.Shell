'Imports System.Runtime.InteropServices

'Public Class PropertyBag
'    Implements IPropertyBag

'    Dim p As Object

'    Public Function Read(<MarshalAs(UnmanagedType.LPWStr)> pszPropName As String, <Out> ByRef pVar As Object, <[In]> pErrorLog As IErrorLog) As Integer Implements IPropertyBag.Read
'        If pszPropName = Functions.STR_ENUM_ITEMS_FLAGS Then
'            pVar = p
'            Return 0
'        Else
'            Return -1
'        End If
'    End Function

'    Public Function Write(<MarshalAs(UnmanagedType.LPWStr)> pszPropName As String, <[In]> ByRef pVar As Object) As Integer Implements IPropertyBag.Write
'        If pszPropName = Functions.STR_ENUM_ITEMS_FLAGS Then
'            p = pVar
'            Return 0
'        Else
'            Return -1
'        End If
'    End Function
'End Class
