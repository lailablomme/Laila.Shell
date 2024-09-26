Public Class ContextMenuCB
    Implements IContextMenuCB

    Public Function CallBack(psf As IntPtr, hwndOwner As IntPtr, pdtobj As IntPtr, uMsg As UInteger, wParam As UInteger, lParam As UInteger) As Integer Implements IContextMenuCB.CallBack
        Debug.WriteLine(CType(uMsg, DFM).ToString())
        Select Case uMsg
            Case DFM.DFM_MERGECONTEXTMENU : Return HResult.Ok
            Case DFM.DFM_INVOKECOMMAND : Return HResult.False
            Case DFM.DFM_INVOKECOMMANDEX : Return HResult.False
            Case Else
                Return HResult.NotImplemented
        End Select
    End Function
End Class
