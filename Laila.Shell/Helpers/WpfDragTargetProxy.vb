Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms.DataFormats
Imports System.Windows.Interop
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Windows
Imports Microsoft.Xaml.Behaviors

Namespace Helpers
    <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959b"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxy")>
    Public Class WpfDragTargetProxy
        Implements IDropTarget

        Private Shared _controls As Dictionary(Of UIElement, BaseDropTarget) = New Dictionary(Of UIElement, BaseDropTarget)()
        Private Shared _hwnds As Dictionary(Of UIElement, IntPtr) = New Dictionary(Of UIElement, IntPtr)()
        Private Shared _activeDropTarget As BaseDropTarget
        Private Shared _instance As WpfDragTargetProxy = New WpfDragTargetProxy()
        Private Shared _hasDragImage As Boolean

        Private _dataObject As IDataObject_PreserveSig
        Private _dropTargetHelper As IDropTargetHelper

        Public Sub New()
            Functions.CoCreateInstance(Guids.CLSID_DragDropHelper, IntPtr.Zero,
                &H1, GetType(IDropTargetHelper).GUID, _dropTargetHelper)
        End Sub

        Public Shared Sub RegisterDragDrop(element As UIElement, dropTarget As BaseDropTarget)
            If Not _controls.ContainsKey(element) Then
                Dim hwnd As IntPtr = GetHwndFromControl(element)
                If Not _controls.Keys.ToList().Exists(Function(c) GetHwndFromControl(c).Equals(hwnd)) Then
                    Functions.RevokeDragDrop(hwnd)
                    Dim h As HRESULT = Functions.RegisterDragDrop(hwnd, _instance)
                    Debug.WriteLine("RegisterDragDrop returned " & h.ToString())
                    If Not h = HRESULT.S_OK Then
                        Dim ex As InvalidOperationException = New InvalidOperationException("Error registering drag and drop.")
                        ex.HResult = h
                        Throw ex
                    End If
                End If
                _controls.Add(element, dropTarget)
                _hwnds.Add(element, hwnd)
            Else
                Throw New InvalidOperationException("This control is already registered.")
            End If
        End Sub

        Public Shared Sub RevokeDragDrop(element As UIElement)
            If _controls.ContainsKey(element) Then
                Dim hwnd As IntPtr = _hwnds(element)
                _controls.Remove(element)
                _hwnds.Remove(element)
                If Not _controls.Keys.ToList().Exists(Function(c) GetHwndFromControl(c).Equals(hwnd)) Then
                    Dim h As HRESULT = Functions.RevokeDragDrop(hwnd)
                    Debug.WriteLine("RevokeDragDrop returned " & h.ToString())
                    If Not h = HRESULT.S_OK Then
                        Dim ex As InvalidOperationException = New InvalidOperationException("Error revoking drag and drop.")
                        ex.HResult = h
                        Throw ex
                    End If
                End If
            Else
                Throw New InvalidOperationException("This control is not registered.")
            End If
        End Sub

        Public Shared Function GetHwndFromControl(element As UIElement) As IntPtr
            If _hwnds.ContainsKey(element) Then
                Return _hwnds(element)
            Else
                Return CType(HwndSource.FromVisual(element), HwndSource).Handle
            End If
        End Function

        Public Shared Function GetDropTargetFromWIN32POINT(ptWIN32 As WIN32POINT) As IDropTarget
            Dim kvps As List(Of KeyValuePair(Of UIElement, BaseDropTarget)) =
                _controls.Where(Function(kp) Not PresentationSource.FromVisual(kp.Key) Is Nothing _
                                             AndAlso Not kp.Key.InputHitTest(UIHelper.WIN32POINTToUIElement(ptWIN32, kp.Key)) Is Nothing).ToList()
            Return kvps.FirstOrDefault(Function(kvp) _
                UIHelper.FindVisualChildren(Of Control)(kvp.Key) _
                     .FirstOrDefault(Function(child) _
                         kvps.Select(Function(kvp2) kvp2.Key).Contains(child)) Is Nothing).Value
        End Function

        Private Shared _type As DROPIMAGETYPE = -5
        Public Shared Sub SetDropDescription(dataObject As IDataObject_PreserveSig, type As DROPIMAGETYPE, message As String, insert As String)
            If Not _hasDragImage Then Return

            Dim dropDescription As DROPDESCRIPTION = New DROPDESCRIPTION() With {
                .type = type,
                .szMessage = message,
                .szInsert = insert
            }
            Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of DROPDESCRIPTION))
            Marshal.StructureToPtr(dropDescription, ptr, False)
            Dim format2 As FORMATETC = New FORMATETC() With {
                .cfFormat = Functions.RegisterClipboardFormat("DropDescription"),
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .ptd = IntPtr.Zero,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            Dim medium As STGMEDIUM = New STGMEDIUM() With {
                .pUnkForRelease = Nothing,
                .tymed = TYMED.TYMED_HGLOBAL,
                .unionmember = ptr
            }
            dataObject.SetData(format2, medium, True)
            Debug.WriteLine("Drop description set to " & type.ToString()) '& "   h=" & h.ToString())
        End Sub

        Public Function DragEnter(pDataObj As IDataObject_PreserveSig, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter
            Debug.WriteLine("WpfDragTargetProxy.DragEnter")
            _dataObject = pDataObj

            _hasDragImage = Clipboard.GetHasGlobalData(_dataObject, "DragImageBits")

            If Drag._isDragging Then
            End If

            Dim dropTarget As BaseDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                _activeDropTarget = dropTarget
                Dim hwnd As IntPtr = _hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key)
                Dim h As HRESULT = _activeDropTarget.DragEnter(pDataObj, grfKeyState, ptWIN32, pdwEffect)
                Debug.WriteLine("_dropTargetHelper.DragEnter")
                If _hasDragImage Then _dropTargetHelper.DragEnter(hwnd, _dataObject, ptWIN32, pdwEffect)
                Debug.WriteLine("DragEnter out")
                Return h
            Else
                _activeDropTarget = Nothing
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.S_OK
            End If
        End Function

        Public Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver
            Debug.WriteLine("WpfDragTargetProxy.DragOver")
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                If Not _activeDropTarget Is Nothing AndAlso dropTarget.Equals(_activeDropTarget) Then
                    Dim h As HRESULT = _activeDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                    Debug.WriteLine("_dropTargetHelper.DragOver")
                    If _hasDragImage Then _dropTargetHelper.DragOver(ptWIN32, pdwEffect)
                    Return h
                Else
                    If Not _activeDropTarget Is Nothing Then
                        Debug.WriteLine("_dropTargetHelper.DragLeave")
                        If _hasDragImage Then _dropTargetHelper.DragLeave()
                        _activeDropTarget.DragLeave()
                    End If

                    _activeDropTarget = dropTarget
                    Dim h As HRESULT = _activeDropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                    Debug.WriteLine("_dropTargetHelper.DragEnter2")
                    If _hasDragImage Then _dropTargetHelper.DragEnter(_hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key),
                                                                  _dataObject, ptWIN32, pdwEffect)
                    Return h
                End If
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                If Not _activeDropTarget Is Nothing Then
                    Try
                        Debug.WriteLine("_dropTargetHelper.DragLeave2")
                        If _hasDragImage Then _dropTargetHelper.DragLeave()
                        Dim h As HRESULT = _activeDropTarget.DragLeave()
                        Return h
                    Finally
                        _activeDropTarget = Nothing
                    End Try
                End If
                Return HRESULT.S_OK
            End If
        End Function

        Public Function DragLeave() As Integer Implements IDropTarget.DragLeave
            Debug.WriteLine("WpfDragTargetProxy.DragLeave")
            If Not _activeDropTarget Is Nothing Then
                Try
                    Debug.WriteLine("_dropTargetHelper.DragLeave3")
                    If _hasDragImage Then _dropTargetHelper.DragLeave()
                    Dim h As HRESULT = _activeDropTarget.DragLeave()
                    Return h
                Finally
                    _activeDropTarget = Nothing
                End Try
            Else
                Return HRESULT.S_OK
            End If
        End Function

        Public Function Drop(pDataObj As IDataObject_PreserveSig, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
            Debug.WriteLine("WpfDragTargetProxy.Drop")
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                Debug.WriteLine("_dropTargetHelper.Drop")
                If _hasDragImage Then _dropTargetHelper.Drop(_dataObject, ptWIN32, pdwEffect)
                Return dropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.S_OK
            End If
        End Function
    End Class
End Namespace