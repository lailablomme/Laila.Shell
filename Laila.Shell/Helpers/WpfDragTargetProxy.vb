Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Controls
Imports System.Windows.Interop
Imports System.Linq
Imports Laila.Shell.ViewModels
Imports System.Windows

Namespace Helpers
    <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959b"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxy")>
    Public Class WpfDragTargetProxy
        Implements IDropTarget

        Private Shared _controls As Dictionary(Of Control, IDropTarget) = New Dictionary(Of Control, IDropTarget)()
        Private Shared _hwnds As Dictionary(Of Control, IntPtr) = New Dictionary(Of Control, IntPtr)()
        Private Shared _activeDropTarget As IDropTarget
        Private Shared _instance As WpfDragTargetProxy = New WpfDragTargetProxy()

        Private _dataObject As ComTypes.IDataObject

        Public Shared Sub RegisterDragDrop(control As Control, dropTarget As IDropTarget)
            If Not _controls.ContainsKey(control) Then
                Dim hwnd As IntPtr = GetHwndFromControl(control)
                If Not _controls.Keys.ToList().Exists(Function(c) GetHwndFromControl(c).Equals(hwnd)) Then
                    Functions.RevokeDragDrop(hwnd)
                    Dim h As HResult = Functions.RegisterDragDrop(hwnd, _instance)
                    Debug.WriteLine("RegisterDragDrop returned " & h.ToString())
                    If Not h = HResult.Ok Then
                        Dim ex As InvalidOperationException = New InvalidOperationException("Error registering drag and drop.")
                        ex.HResult = h
                        Throw ex
                    End If
                End If
                _controls.Add(control, dropTarget)
                _hwnds.Add(control, hwnd)
            Else
                Throw New InvalidOperationException("This control is already registered.")
            End If
        End Sub

        Public Shared Sub RevokeDragDrop(control As Control)
            If _controls.ContainsKey(control) Then
                Dim hwnd As IntPtr = _hwnds(control)
                _controls.Remove(control)
                _hwnds.Remove(control)
                If Not _controls.Keys.ToList().Exists(Function(c) GetHwndFromControl(c).Equals(hwnd)) Then
                    Dim h As HResult = Functions.RevokeDragDrop(hwnd)
                    Debug.WriteLine("RevokeDragDrop returned " & h.ToString())
                    If Not h = HResult.Ok Then
                        Dim ex As InvalidOperationException = New InvalidOperationException("Error revoking drag and drop.")
                        ex.HResult = h
                        Throw ex
                    End If
                End If
            Else
                Throw New InvalidOperationException("This control is not registered.")
            End If
        End Sub

        Public Shared Function GetHwndFromControl(control As Control) As IntPtr
            If _hwnds.ContainsKey(control) Then
                Return _hwnds(control)
            Else
                Return CType(HwndSource.FromVisual(control), HwndSource).Handle
            End If
        End Function

        Public Shared Function GetDropTargetFromWIN32POINT(ptWIN32 As WIN32POINT) As IDropTarget
            Return _controls.FirstOrDefault(Function(kp) Not kp.Key.InputHitTest(UIHelper.WIN32POINTToControl(ptWIN32, kp.Key)) Is Nothing).Value
        End Function

        Public Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter
            _dataObject = pDataObj
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(pt)
            If Not dropTarget Is Nothing Then
                _activeDropTarget = dropTarget
                Return dropTarget.DragEnter(pDataObj, grfKeyState, pt, pdwEffect)
            Else
                _activeDropTarget = Nothing
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HResult.Ok
            End If
        End Function

        Public Function DragOver(grfKeyState As Integer, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(pt)
            If Not dropTarget Is Nothing Then
                If Not _activeDropTarget Is Nothing Then
                    Return _activeDropTarget.DragOver(grfKeyState, pt, pdwEffect)
                Else
                    _activeDropTarget = dropTarget
                    Return _activeDropTarget.DragEnter(_dataObject, grfKeyState, pt, pdwEffect)
                End If
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                If Not _activeDropTarget Is Nothing Then
                    Try
                        Return _activeDropTarget.DragLeave()
                    Finally
                        _activeDropTarget = Nothing
                    End Try
                End If
                Return HResult.Ok
            End If
        End Function

        Public Function DragLeave() As Integer Implements IDropTarget.DragLeave
            If Not _activeDropTarget Is Nothing Then
                Try
                    Return _activeDropTarget.DragLeave()
                Finally
                    _activeDropTarget = Nothing
                End Try
            Else
                Return HResult.Ok
            End If
        End Function

        Public Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(pt)
            If Not dropTarget Is Nothing Then
                Return dropTarget.Drop(pDataObj, grfKeyState, pt, pdwEffect)
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HResult.Ok
            End If
        End Function
    End Class
End Namespace