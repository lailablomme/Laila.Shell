Public Class ContextMenuCB
    Implements IContextMenuCB

    Public Function CallBack(psf As IntPtr, hwndOwner As IntPtr, pdtobj As IntPtr, uMsg As UInteger, wParam As UInteger, lParam As UInteger) As Integer Implements IContextMenuCB.CallBack
        Debug.WriteLine(CType(uMsg, DFM).ToString())
        Select Case uMsg
            Case DFM.DFM_MERGECONTEXTMENU : Return HRESULT.Ok
            Case DFM.DFM_INVOKECOMMAND : Return HRESULT.False
            Case DFM.DFM_INVOKECOMMANDEX : Return HRESULT.False
            Case Else
                Return HRESULT.NotImplemented
        End Select
    End Function
End Class
