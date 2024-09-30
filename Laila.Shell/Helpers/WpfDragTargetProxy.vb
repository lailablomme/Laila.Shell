Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Controls
Imports System.Windows.Interop
Imports System.Linq
Imports Laila.Shell.ViewModels
Imports System.Windows
Imports Microsoft

Namespace Helpers
    <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959b"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxy")>
    Public Class WpfDragTargetProxy
        Implements IDropTarget

        Private Shared _controls As Dictionary(Of Control, BaseDropTarget) = New Dictionary(Of Control, BaseDropTarget)()
        Private Shared _hwnds As Dictionary(Of Control, IntPtr) = New Dictionary(Of Control, IntPtr)()
        Private Shared _activeDropTarget As BaseDropTarget
        Private Shared _instance As WpfDragTargetProxy = New WpfDragTargetProxy()

        Private _dataObject As ComTypes.IDataObject
        Private _dropTargetHelper As IDropTargetHelper

        Public Sub New()
            Functions.CoCreateInstance(Guids.CLSID_DragDropHelper, IntPtr.Zero,
                &H1, GetType(IDropTargetHelper).GUID, _dropTargetHelper)
        End Sub

        Public Shared Sub RegisterDragDrop(control As Control, dropTarget As BaseDropTarget)
            If Not _controls.ContainsKey(control) Then
                Dim hwnd As IntPtr = GetHwndFromControl(control)
                If Not _controls.Keys.ToList().Exists(Function(c) GetHwndFromControl(c).Equals(hwnd)) Then
                    Functions.RevokeDragDrop(hwnd)
                    Dim h As HRESULT = Functions.RegisterDragDrop(hwnd, _instance)
                    Debug.WriteLine("RegisterDragDrop returned " & h.ToString())
                    If Not h = HRESULT.Ok Then
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
                    Dim h As HRESULT = Functions.RevokeDragDrop(hwnd)
                    Debug.WriteLine("RevokeDragDrop returned " & h.ToString())
                    If Not h = HRESULT.Ok Then
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

        Public Function DragEnter(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter
            _dataObject = pDataObj
            Dim dropTarget As BaseDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                _activeDropTarget = dropTarget
                Dim h As HRESULT = _activeDropTarget.DragEnterInternal(pDataObj, grfKeyState, ptWIN32, pdwEffect)
                System.Windows.Application.Current.Dispatcher.Invoke(
                    Sub()
                    End Sub)
                Dim hwnd As IntPtr = _hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key)
                _dropTargetHelper.DragEnter(hwnd, _dataObject, ptWIN32, pdwEffect)
                Return h
            Else
                _activeDropTarget = Nothing
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.Ok
            End If
        End Function

        Public Function DragOver(grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                If Not _activeDropTarget Is Nothing Then
                    Dim h As HRESULT = _activeDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                    _dropTargetHelper.DragOver(ptWIN32, pdwEffect)
                    Return h
                Else
                    _activeDropTarget = dropTarget
                    Dim h As HRESULT = _activeDropTarget.DragEnterInternal(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                    _dropTargetHelper.DragEnter(_hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key),
                                                _dataObject, ptWIN32, pdwEffect)
                    Return h
                End If
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                If Not _activeDropTarget Is Nothing Then
                    Try
                        Dim h As HRESULT = _activeDropTarget.DragLeave()
                        _dropTargetHelper.DragLeave()
                        Return h
                    Finally
                        _activeDropTarget = Nothing
                    End Try
                End If
                Return HRESULT.Ok
            End If
        End Function

        Public Function DragLeave() As Integer Implements IDropTarget.DragLeave
            If Not _activeDropTarget Is Nothing Then
                Try
                    Dim h As HRESULT = _activeDropTarget.DragLeave()
                    _dropTargetHelper.DragLeave()
                    Return h
                Finally
                    _activeDropTarget = Nothing
                End Try
            Else
                Return HRESULT.Ok
            End If
        End Function

        Public Function Drop(pDataObj As ComTypes.IDataObject, grfKeyState As Integer, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                _dropTargetHelper.Drop(_dataObject, ptWIN32, pdwEffect)
                Return dropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.Ok
            End If
        End Function
    End Class
End Namespace