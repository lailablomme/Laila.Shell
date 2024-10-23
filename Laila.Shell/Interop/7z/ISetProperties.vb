Imports System.Runtime.InteropServices

<ComImport>
<Guid("23170F69-40C1-278A-0000-000600030000")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface ISetProperties
    ''' <summary>
    ''' Sets properties to the compressing file.
    ''' </summary>
    ''' <param name="names">Property names.</param>
    ''' <param name="values">Property values.</param>
    ''' <param name="numProperties">Number of properties.</param>
    ''' <returns>Operation result.</returns>
    <PreserveSig>
    Function SetProperties(
        <[In], MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPWStr, SizeParamIndex:=2)> names As String(),
        values As IntPtr,
        numProperties As UInteger
    ) As Integer
End Interface