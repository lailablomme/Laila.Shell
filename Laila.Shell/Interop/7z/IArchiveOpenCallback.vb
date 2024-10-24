Imports System.Runtime.InteropServices

<ComImport>
<Guid("23170F69-40C1-278A-0000-000600100000")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IArchiveOpenCallback
    ''' <summary>
    ''' Gets the total size of the compressed file upon decompression.
    ''' </summary>
    ''' <param name="count">Total number of files.</param>
    ''' <param name="bytes">Total compressed bytes.</param>
    ''' <returns>ErrorCode.None for success.</returns>
    <PreserveSig>
    Function SetTotal(count As IntPtr, bytes As IntPtr) As Integer

    ''' <summary>
    ''' Gets the size of the stream ready to be read.
    ''' </summary>
    ''' <param name="count">Number of files.</param>
    ''' <param name="bytes">Completed bytes.</param>
    ''' <returns>ErrorCode.None for success.</returns>
    <PreserveSig>
    Function SetCompleted(count As IntPtr, bytes As IntPtr) As Integer
End Interface