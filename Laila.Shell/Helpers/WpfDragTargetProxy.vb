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
        Public Shared _isDropDescriptionSet As Boolean = False

        Private _dataObject As IDataObject
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
            Dim kvps As List(Of KeyValuePair(Of Control, BaseDropTarget)) =
                _controls.Where(Function(kp) Not PresentationSource.FromVisual(kp.Key) Is Nothing _
                                             AndAlso Not kp.Key.InputHitTest(UIHelper.WIN32POINTToControl(ptWIN32, kp.Key)) Is Nothing).ToList()
            Return kvps.FirstOrDefault(Function(kvp) _
                UIHelper.FindVisualChildren(Of Control)(kvp.Key) _
                     .FirstOrDefault(Function(child) _
                         kvps.Select(Function(kvp2) kvp2.Key).Contains(child)) Is Nothing).Value
        End Function

        Public Shared Sub SetDropDescription(dataObject As IDataObject, type As DROPIMAGETYPE, message As String, insert As String)
            Dim dropDescription As DROPDESCRIPTION
            Dim format As FORMATETC = New FORMATETC() With {
                .cfFormat = Functions.RegisterClipboardFormat("DropDescription"),
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .ptd = IntPtr.Zero,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            Dim doSet As Boolean = True
            If Drag._isDragging AndAlso dataObject.QueryGetData(format) = 0 Then
                Dim m As STGMEDIUM = New STGMEDIUM()
                dataObject.GetData(format, m)
                If m.tymed <> TYMED.TYMED_NULL Then
                    dropDescription = Marshal.PtrToStructure(Of DROPDESCRIPTION)(m.unionmember)
                    'Debug.WriteLine("Drop description type = " & dropDescription.type.ToString())
                    dropDescription.type = type
                    dropDescription.szMessage = message
                    dropDescription.szInsert = insert
                    Marshal.StructureToPtr(dropDescription, m.unionmember, False)
                    'Debug.WriteLine("Drop description overwritten to " & type.ToString())
                    doSet = False
                    'Else
                    'Debug.WriteLine("Drop description is NULL         " & type.ToString())
                End If
                'Else
                'Debug.WriteLine("Drop description does not exist         " & type.ToString())
                'If type = DROPIMAGETYPE.DROPIMAGE_INVALID Then doSet = False
            End If
            If doSet Then
                dropDescription = New DROPDESCRIPTION() With {
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
                    .pUnkForRelease = IntPtr.Zero,
                    .tymed = TYMED.TYMED_HGLOBAL,
                    .unionmember = ptr
                }
                Dim h As HRESULT = dataObject.SetData(format2, medium, True)
                Debug.WriteLine("Drop description set to " & type.ToString() & "   h=" & h.ToString())
            End If
            _isDropDescriptionSet = True
        End Sub

        Public Function DragEnter(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter
            'Debug.WriteLine("WpfDragTargetProxy.DragEnter")
            'Drag.InitializeDragImage()
            _dataObject = pDataObj

            If Drag._isDragging Then
            End If

            Dim dropTarget As BaseDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                _activeDropTarget = dropTarget
                Dim hwnd As IntPtr = _hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key)
                _isDropDescriptionSet = False
                Dim h As HRESULT = _activeDropTarget.DragEnter(pDataObj, grfKeyState, ptWIN32, pdwEffect)
                If Drag.GetHasGlobalData(_dataObject, "DropDescription") AndAlso Not _isDropDescriptionSet AndAlso Drag._isDragging Then
                    SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_INVALID, Nothing, Nothing)
                End If
                'Debug.WriteLine("_dropTargetHelper.DragEnter")
                _dropTargetHelper.DragEnter(hwnd, _dataObject, ptWIN32, pdwEffect)
                Return h
            Else
                _activeDropTarget = Nothing
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.Ok
            End If
        End Function

        Public Function DragOver(grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver
            'Debug.WriteLine("WpfDragTargetProxy.DragOver")
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                If Not _activeDropTarget Is Nothing AndAlso dropTarget.Equals(_activeDropTarget) Then
                    _isDropDescriptionSet = False
                    Dim h As HRESULT = _activeDropTarget.DragOver(grfKeyState, ptWIN32, pdwEffect)
                    If Drag.GetHasGlobalData(_dataObject, "DropDescription") AndAlso Not _isDropDescriptionSet AndAlso Drag._isDragging Then
                        SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_INVALID, Nothing, Nothing)
                    End If
                    'Debug.WriteLine("_dropTargetHelper.DragOver")
                    _dropTargetHelper.DragOver(ptWIN32, pdwEffect)
                    Return h
                Else
                    If Not _activeDropTarget Is Nothing Then
                        'Debug.WriteLine("_dropTargetHelper.DragLeave")
                        _dropTargetHelper.DragLeave()
                        _activeDropTarget.DragLeave()
                    End If

                    _activeDropTarget = dropTarget
                    _isDropDescriptionSet = False
                    Dim h As HRESULT = _activeDropTarget.DragEnter(_dataObject, grfKeyState, ptWIN32, pdwEffect)
                    If Drag.GetHasGlobalData(_dataObject, "DropDescription") AndAlso Not _isDropDescriptionSet AndAlso Drag._isDragging Then
                        SetDropDescription(_dataObject, DROPIMAGETYPE.DROPIMAGE_INVALID, Nothing, Nothing)
                    End If
                    'Debug.WriteLine("_dropTargetHelper.DragEnter")
                    _dropTargetHelper.DragEnter(_hwnds(_controls.FirstOrDefault(Function(kv) kv.Value.Equals(_activeDropTarget)).Key),
                                                _dataObject, ptWIN32, pdwEffect)
                    Return h
                End If
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                If Not _activeDropTarget Is Nothing Then
                    Try
                        'Debug.WriteLine("_dropTargetHelper.DragLeave")
                        _dropTargetHelper.DragLeave()
                        Dim h As HRESULT = _activeDropTarget.DragLeave()
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
                    'Debug.WriteLine("_dropTargetHelper.DragLeave")
                    _dropTargetHelper.DragLeave()
                    Dim h As HRESULT = _activeDropTarget.DragLeave()
                    Return h
                Finally
                    _activeDropTarget = Nothing
                End Try
            Else
                Return HRESULT.Ok
            End If
        End Function

        Public Function Drop(pDataObj As IDataObject, grfKeyState As MK, ptWIN32 As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
            Dim dropTarget As IDropTarget = GetDropTargetFromWIN32POINT(ptWIN32)
            If Not dropTarget Is Nothing Then
                'Debug.WriteLine("_dropTargetHelper.Drop")
                _dropTargetHelper.Drop(_dataObject, ptWIN32, pdwEffect)
                Return dropTarget.Drop(pDataObj, grfKeyState, ptWIN32, pdwEffect)
            Else
                pdwEffect = DROPEFFECT.DROPEFFECT_NONE
                Return HRESULT.Ok
            End If
        End Function
    End Class
End Namespace