Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport(), Guid("F79B1D79-F493-4A93-81D2-FC7D4C0D79BC"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileOperationProgressSink
        ' This method is called when an operation is started.
        Sub StartOperation()

        ' This method is called to report progress.
        Sub ReportProgress(ByVal dwComplete As UInteger, ByVal dwTotal As UInteger)

        ' This method is called to indicate that an operation is completed.
        Sub FinishOperation()

        ' This method is called to report an error.
        Sub ReportError(ByVal hResult As Integer, ByVal pszError As String)

        ' This method is called to set the title of the progress dialog.
        Sub SetTitle(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszTitle As String)

        ' This method is called to set the message in the progress dialog.
        Sub SetMessage(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszMessage As String)

        ' This method is called to indicate that the operation should be canceled.
        Sub SetCancelMessage(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszCancelMessage As String)

        ' This method is called to indicate the number of items to be processed.
        Sub SetTotalItems(<MarshalAs(UnmanagedType.U4)> ByVal dwTotalItems As UInteger)
    End Interface
End Namespace
