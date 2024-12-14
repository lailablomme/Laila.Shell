Imports Laila.Shell.Controls
Imports System.Windows
Imports Microsoft
Imports System.Windows.Controls
Imports System.Threading

Namespace Helpers
    Public Class TabControlDropTarget
        Inherits BaseDropTarget

        Private _tabControl As Controls.TabControl
        Private _dragOpenTimer As Timer
        Private _lastOverItem As Controls.TabControl.TabData

        Public Sub New(tabControl As Controls.TabControl)
            _tabControl = tabControl
        End Sub

        Public Overrides Function DragEnter(pDataObj As IDataObject, grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Return dragPoint(grfKeyState, pt, pdwEffect)
        End Function

        Public Overrides Function DragOver(grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            Return dragPoint(grfKeyState, pt, pdwEffect)
        End Function

        Public Overrides Function DragLeave() As Integer
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.Dispose()
            End If
        End Function

        Public Overrides Function Drop(pDataObj As IDataObject, grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer
            If Not _dragOpenTimer Is Nothing Then
                _dragOpenTimer.Dispose()
            End If
        End Function

        Private Function getOverItem(ptWIN32 As WIN32POINT) As Controls.TabControl.TabData
            ' translate point to listview
            Dim pt As Point = UIHelper.WIN32POINTToUIElement(ptWIN32, _tabControl)

            ' find which item we're over
            Dim overObject As IInputElement = _tabControl.InputHitTest(pt)
            Dim overTabItem As TabItem
            If TypeOf overObject Is TabItem Then
                overTabItem = overObject
            Else
                overTabItem = UIHelper.GetParentOfType(Of TabItem)(overObject)
            End If
            If Not overTabItem Is Nothing Then
                Return overTabItem.DataContext
            Else
                Return Nothing
            End If
        End Function

        Private Function dragPoint(grfKeyState As UInteger, ptWIN32 As WIN32POINT, ByRef pdwEffect As UInteger) As Integer
            Dim overItem As Controls.TabControl.TabData = getOverItem(ptWIN32)

            If Not overItem Is Nothing AndAlso (_lastOverItem Is Nothing OrElse Not _lastOverItem.Equals(overItem)) Then
                If Not _dragOpenTimer Is Nothing Then
                    _dragOpenTimer.Dispose()
                End If

                _dragOpenTimer = New Timer(New TimerCallback(
                    Sub()
                        UIHelper.OnUIThread(
                            Sub()
                                _tabControl.SelectedItem = overItem
                            End Sub)
                        _dragOpenTimer.Dispose()
                        _dragOpenTimer = Nothing
                    End Sub), Nothing, 1250, 0)
            End If

            _lastOverItem = overItem

            pdwEffect = DROPEFFECT.DROPEFFECT_NONE
            Return HRESULT.S_OK
        End Function
    End Class
End Namespace