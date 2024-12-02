﻿Imports System.Runtime.InteropServices

<ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IInitializeWithStream
    ''' <summary>
    ''' Initializes the handler with a stream.
    ''' </summary>
    ''' <param name="pstream">The stream to initialize with.</param>
    ''' <param name="grfMode">The access mode (from the STGM enumeration).</param>
    <PreserveSig>
    Function Initialize(pstream As System.Runtime.InteropServices.ComTypes.IStream, grfMode As UInteger) As Integer
End Interface