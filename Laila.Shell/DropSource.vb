'Imports System.Runtime.InteropServices

'Public Class DropSource
'    Implements IDropSource

'    Public Function QueryContinueDrag(<[In]> <MarshalAs(UnmanagedType.Bool)> fEscapePressed As Boolean, <[In]> grfKeyState As Integer) As Integer Implements IDropSource.QueryContinueDrag
'        If fEscapePressed Then
'            Return DragDropResult.DRAGDROP_S_CANCEL
'        End If
'        If (grfKeyState And MK.MK_LBUTTON) = 0 Then
'            Return DragDropResult.DRAGDROP_S_DROP
'        End If
'        Return DragDropResult.S_OK
'    End Function

'    Public Function GiveFeedback(<[In]> dwEffect As Integer) As Integer Implements IDropSource.GiveFeedback
'        Return DragDropResult.S_OK
'    End Function
'End Class
