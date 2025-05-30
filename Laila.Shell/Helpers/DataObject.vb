﻿Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop

Namespace Helpers
    <ComVisible(True), Guid("a3b5f678-9cde-4f12-85a6-3f7b2e9d1c45")>
    Public Class DataObject
        Implements IDataObject_PreserveSig, IDisposable

        Private _data As Dictionary(Of FORMATETC, Tuple(Of STGMEDIUM, Boolean)) = New Dictionary(Of FORMATETC, Tuple(Of STGMEDIUM, Boolean))()
        Private disposedValue As Boolean

        Public Function GetData(ByRef format As ComTypes.FORMATETC, ByRef medium As ComTypes.STGMEDIUM) As Integer Implements IDataObject_PreserveSig.GetData
            Select Case format.cfFormat
                Case ClipboardFormat.CF_HDROP
                    Debug.WriteLine("GetData CF_HDROP")
                Case Functions.RegisterClipboardFormat("DragImageBits")
                    Debug.WriteLine("GetData DragImageBits")
                Case Functions.RegisterClipboardFormat("ComputedImage")
                    Debug.WriteLine("GetData ComputedImage")
                Case Functions.RegisterClipboardFormat("IsComputingImage")
                    Debug.WriteLine("GetData IsComputingImage")
                Case Functions.RegisterClipboardFormat("DragSourceHelperFlags")
                    Debug.WriteLine("GetData DragSourceHelperFlags")
                Case Functions.RegisterClipboardFormat("DragContext")
                    Debug.WriteLine("GetData DragContext")
                Case Functions.RegisterClipboardFormat("DragWindow")
                    Debug.WriteLine("GetData DragWindow")
                Case Functions.RegisterClipboardFormat("DisableDragText")
                    Debug.WriteLine("GetData DisableDragText")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("GetData DropDescription")
                Case Functions.RegisterClipboardFormat("IsShowingLayered")
                    Debug.WriteLine("GetData IsShowingLayered")
                Case Functions.RegisterClipboardFormat("IsShowingText")
                    Debug.WriteLine("GetData IsShowingText")
                Case Functions.RegisterClipboardFormat("UsingDefaultDragImage")
                    Debug.WriteLine("GetData UsingDefaultDragImage")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptorW")
                    Debug.WriteLine("GetData FileGroupDescriptorW")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptor")
                    Debug.WriteLine("GetData FileGroupDescriptor")
                Case Functions.RegisterClipboardFormat("FileContents")
                    Debug.WriteLine("GetData FileContents")
                Case Functions.RegisterClipboardFormat("Shell IDList Array")
                    Debug.WriteLine("GetData Shell IDList Array")
                Case Functions.RegisterClipboardFormat("Shell Object Offsets")
                    Debug.WriteLine("GetData Shell Object Offsets")
                Case Functions.RegisterClipboardFormat("Net Resource")
                    Debug.WriteLine("GetData Net Resource")
                Case Functions.RegisterClipboardFormat("FileName")
                    Debug.WriteLine("GetData FileName")
                Case Functions.RegisterClipboardFormat("FileNameW")
                    Debug.WriteLine("GetData FileNameW")
                Case Functions.RegisterClipboardFormat("PrinterFriendlyName")
                    Debug.WriteLine("GetData PrinterFriendlyName")
                Case Functions.RegisterClipboardFormat("FileNameMap")
                    Debug.WriteLine("GetData FileNameMap")
                Case Functions.RegisterClipboardFormat("FileNameMapW")
                    Debug.WriteLine("GetData FileNameMapW")
                Case Functions.RegisterClipboardFormat("UniformResourceLocator")
                    Debug.WriteLine("GetData UniformResourceLocator")
                Case Functions.RegisterClipboardFormat("UniformResourceLocatorW")
                    Debug.WriteLine("GetData UniformResourceLocatorW")
                Case Functions.RegisterClipboardFormat("Preferred DropEffect")
                    Debug.WriteLine("GetData Preferred DropEffect")
                Case Functions.RegisterClipboardFormat("Performed DropEffect")
                    Debug.WriteLine("GetData Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Paste Succeeded")
                    Debug.WriteLine("GetData Paste Succeeded")
                Case Functions.RegisterClipboardFormat("InShellDragLoop")
                    Debug.WriteLine("GetData InShellDragLoop")
                Case Functions.RegisterClipboardFormat("MountedVolume")
                    Debug.WriteLine("GetData MountedVolume")
                Case Functions.RegisterClipboardFormat("PersistedDataObject")
                    Debug.WriteLine("GetData PersistedDataObject")
                Case Functions.RegisterClipboardFormat("TargetCLSID")
                    Debug.WriteLine("GetData TargetCLSID")
                Case Functions.RegisterClipboardFormat("Logical Performed DropEffect")
                    Debug.WriteLine("GetData Logical Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Autoplay Enumerated IDList Array")
                    Debug.WriteLine("GetData Autoplay Enumerated IDList Array")
                Case Functions.RegisterClipboardFormat("UntrustedDragDrop")
                    Debug.WriteLine("GetData UntrustedDragDrop")
                Case Functions.RegisterClipboardFormat("File Attributes Array")
                    Debug.WriteLine("GetData File Attributes Array")
                Case Functions.RegisterClipboardFormat("InvokeCommand DropParam")
                    Debug.WriteLine("GetData InvokeCommand DropParam")
                Case Functions.RegisterClipboardFormat("DropHandlerCLSID")
                    Debug.WriteLine("GetData DropHandlerCLSID")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("GetData DropDescription")
                Case Else
                    Debug.WriteLine("GetData " & format.cfFormat)
            End Select
            Debug.WriteLine("GetData " & format.tymed.ToString() & "," & format.lindex)

            Dim fmt2 As FORMATETC = New FORMATETC() With {
            .cfFormat = format.cfFormat,
            .dwAspect = format.dwAspect,
            .lindex = format.lindex,
            .tymed = format.tymed
        }
            If _data.Keys.ToList().Exists(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                                          AndAlso f.lindex = fmt2.lindex AndAlso (f.tymed And fmt2.tymed)) Then
                Dim m As Tuple(Of STGMEDIUM, Boolean) = _data.ToList().FirstOrDefault(Function(p) p.Key.cfFormat = fmt2.cfFormat AndAlso p.Key.dwAspect = fmt2.dwAspect _
                                          AndAlso p.Key.lindex = fmt2.lindex AndAlso (p.Key.tymed And fmt2.tymed)).Value
                Functions.CopyStgMedium(m.Item1, medium)
                Debug.WriteLine("GetData S_OK")
                Return HRESULT.S_OK
                'If m.Item2 Then
                '    Functions.ReleaseStgMedium(m.Item1)
                'End If
                '_data.Remove(_data.Keys.ToList().FirstOrDefault(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                '                              AndAlso f.lindex = fmt2.lindex AndAlso f.tymed = fmt2.tymed))
                '_data.Add(fmt2, New Tuple(Of STGMEDIUM, Boolean)(medium, m.Item2))
            Else
                Debug.WriteLine("GetData S_FALSE")
                medium.tymed = TYMED.TYMED_NULL
                Return HRESULT.S_FALSE
            End If
        End Function

        Public Function GetDataHere(ByRef format As ComTypes.FORMATETC, ByRef medium As ComTypes.STGMEDIUM) As Integer Implements IDataObject_PreserveSig.GetDataHere
            Debug.WriteLine("GetDataHere")

            Dim fmt2 As FORMATETC = New FORMATETC() With {
            .cfFormat = format.cfFormat,
            .dwAspect = format.dwAspect,
            .lindex = format.lindex,
            .tymed = format.tymed
        }
            If _data.Keys.ToList().Exists(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                                          AndAlso f.lindex = fmt2.lindex AndAlso (f.tymed And fmt2.tymed)) Then
                Dim m As STGMEDIUM = _data.ToList().FirstOrDefault(Function(p) p.Key.cfFormat = fmt2.cfFormat AndAlso p.Key.dwAspect = fmt2.dwAspect _
                                          AndAlso p.Key.lindex = fmt2.lindex AndAlso (p.Key.tymed And fmt2.tymed)).Value.Item1
                medium.unionmember = m.unionmember
                medium.pUnkForRelease = m.pUnkForRelease
                medium.tymed = m.tymed
                Return HRESULT.S_OK
            Else
                Return HRESULT.S_FALSE
            End If
        End Function

        Public Function QueryGetData(ByRef format As ComTypes.FORMATETC) As Integer Implements IDataObject_PreserveSig.QueryGetData
            Select Case format.cfFormat
                Case ClipboardFormat.CF_HDROP
                    Debug.WriteLine("QueryGetData CF_HDROP")
                Case Functions.RegisterClipboardFormat("DragImageBits")
                    Debug.WriteLine("QueryGetData DragImageBits")
                Case Functions.RegisterClipboardFormat("ComputedImage")
                    Debug.WriteLine("QueryGetData ComputedImage")
                Case Functions.RegisterClipboardFormat("IsComputingImage")
                    Debug.WriteLine("QueryGetData IsComputingImage")
                Case Functions.RegisterClipboardFormat("DragSourceHelperFlags")
                    Debug.WriteLine("QueryGetData DragSourceHelperFlags")
                Case Functions.RegisterClipboardFormat("DragContext")
                    Debug.WriteLine("QueryGetData DragContext")
                Case Functions.RegisterClipboardFormat("DragWindow")
                    Debug.WriteLine("QueryGetData DragWindow")
                Case Functions.RegisterClipboardFormat("DisableDragText")
                    Debug.WriteLine("QueryGetData DisableDragText")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("QueryGetData DropDescription")
                Case Functions.RegisterClipboardFormat("IsShowingLayered")
                    Debug.WriteLine("QueryGetData IsShowingLayered")
                Case Functions.RegisterClipboardFormat("IsShowingText")
                    Debug.WriteLine("QueryGetData IsShowingText")
                Case Functions.RegisterClipboardFormat("UsingDefaultDragImage")
                    Debug.WriteLine("QueryGetData UsingDefaultDragImage")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptorW")
                    Debug.WriteLine("QueryGetData FileGroupDescriptorW")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptor")
                    Debug.WriteLine("QueryGetData FileGroupDescriptor")
                Case Functions.RegisterClipboardFormat("FileContents")
                    Debug.WriteLine("QueryGetData FileContents")
                Case Functions.RegisterClipboardFormat("Shell IDList Array")
                    Debug.WriteLine("QueryGetData Shell IDList Array")
                Case Functions.RegisterClipboardFormat("Shell Object Offsets")
                    Debug.WriteLine("QueryGetData Shell Object Offsets")
                Case Functions.RegisterClipboardFormat("Net Resource")
                    Debug.WriteLine("QueryGetData Net Resource")
                Case Functions.RegisterClipboardFormat("FileName")
                    Debug.WriteLine("QueryGetData FileName")
                Case Functions.RegisterClipboardFormat("FileNameW")
                    Debug.WriteLine("QueryGetData FileNameW")
                Case Functions.RegisterClipboardFormat("PrinterFriendlyName")
                    Debug.WriteLine("QueryGetData PrinterFriendlyName")
                Case Functions.RegisterClipboardFormat("FileNameMap")
                    Debug.WriteLine("QueryGetData FileNameMap")
                Case Functions.RegisterClipboardFormat("FileNameMapW")
                    Debug.WriteLine("QueryGetData FileNameMapW")
                Case Functions.RegisterClipboardFormat("UniformResourceLocator")
                    Debug.WriteLine("QueryGetData UniformResourceLocator")
                Case Functions.RegisterClipboardFormat("UniformResourceLocatorW")
                    Debug.WriteLine("QueryGetData UniformResourceLocatorW")
                Case Functions.RegisterClipboardFormat("Preferred DropEffect")
                    Debug.WriteLine("QueryGetData Preferred DropEffect")
                Case Functions.RegisterClipboardFormat("Performed DropEffect")
                    Debug.WriteLine("QueryGetData Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Paste Succeeded")
                    Debug.WriteLine("QueryGetData Paste Succeeded")
                Case Functions.RegisterClipboardFormat("InShellDragLoop")
                    Debug.WriteLine("QueryGetData InShellDragLoop")
                Case Functions.RegisterClipboardFormat("MountedVolume")
                    Debug.WriteLine("QueryGetData MountedVolume")
                Case Functions.RegisterClipboardFormat("PersistedDataObject")
                    Debug.WriteLine("QueryGetData PersistedDataObject")
                Case Functions.RegisterClipboardFormat("TargetCLSID")
                    Debug.WriteLine("QueryGetData TargetCLSID")
                Case Functions.RegisterClipboardFormat("Logical Performed DropEffect")
                    Debug.WriteLine("QueryGetData Logical Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Autoplay Enumerated IDList Array")
                    Debug.WriteLine("QueryGetData Autoplay Enumerated IDList Array")
                Case Functions.RegisterClipboardFormat("UntrustedDragDrop")
                    Debug.WriteLine("QueryGetData UntrustedDragDrop")
                Case Functions.RegisterClipboardFormat("File Attributes Array")
                    Debug.WriteLine("QueryGetData File Attributes Array")
                Case Functions.RegisterClipboardFormat("InvokeCommand DropParam")
                    Debug.WriteLine("QueryGetData InvokeCommand DropParam")
                Case Functions.RegisterClipboardFormat("DropHandlerCLSID")
                    Debug.WriteLine("QueryGetData DropHandlerCLSID")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("QueryGetData DropDescription")
                Case Else
                    Debug.WriteLine("QueryGetData " & format.cfFormat)
            End Select

            Debug.WriteLine("QueryGetData " & format.tymed.ToString() & "," & format.lindex)

            Dim fmt2 As FORMATETC = New FORMATETC() With {
            .cfFormat = format.cfFormat,
            .dwAspect = format.dwAspect,
            .lindex = format.lindex,
            .tymed = format.tymed
        }
            If _data.Keys.ToList().Exists(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                                            AndAlso (f.tymed And fmt2.tymed)) Then
                Debug.WriteLine("QueryGetData S_OK")
                Return HRESULT.S_OK
            Else
                Debug.WriteLine("QueryGetData S_FALSE")
                Return HRESULT.S_FALSE
            End If
        End Function
        Public Function GetCanonicalFormatEtc(ByRef formatIn As ComTypes.FORMATETC, ByRef formatOut As ComTypes.FORMATETC) As Integer Implements IDataObject_PreserveSig.GetCanonicalFormatEtc
            Debug.WriteLine("GetCanonicalFormatEtc")
            formatOut = formatIn
            Return HRESULT.S_OK
        End Function

        Public Function SetData(ByRef format As ComTypes.FORMATETC, ByRef medium As ComTypes.STGMEDIUM, ByVal release As Boolean) As Integer Implements IDataObject_PreserveSig.SetData
            Select Case format.cfFormat
                Case ClipboardFormat.CF_HDROP
                    Debug.WriteLine("SetData CF_HDROP")
                Case Functions.RegisterClipboardFormat("DragImageBits")
                    Debug.WriteLine("SetData DragImageBits")
                Case Functions.RegisterClipboardFormat("ComputedImage")
                    Debug.WriteLine("SetData ComputedImage")
                Case Functions.RegisterClipboardFormat("IsComputingImage")
                    Debug.WriteLine("SetData IsComputingImage")
                Case Functions.RegisterClipboardFormat("DragSourceHelperFlags")
                    Debug.WriteLine("SetData DragSourceHelperFlags")
                Case Functions.RegisterClipboardFormat("DragContext")
                    Debug.WriteLine("SetData DragContext")
                Case Functions.RegisterClipboardFormat("DragWindow")
                    Debug.WriteLine("SetData DragWindow")
                Case Functions.RegisterClipboardFormat("DisableDragText")
                    Debug.WriteLine("SetData DisableDragText")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("SetData DropDescription")
                Case Functions.RegisterClipboardFormat("IsShowingLayered")
                    Debug.WriteLine("SetData IsShowingLayered")
                Case Functions.RegisterClipboardFormat("IsShowingText")
                    Debug.WriteLine("SetData IsShowingText")
                Case Functions.RegisterClipboardFormat("UsingDefaultDragImage")
                    Debug.WriteLine("SetData UsingDefaultDragImage")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptorW")
                    Debug.WriteLine("SetData FileGroupDescriptorW")
                Case Functions.RegisterClipboardFormat("FileGroupDescriptor")
                    Debug.WriteLine("SetData FileGroupDescriptor")
                Case Functions.RegisterClipboardFormat("FileContents")
                    Debug.WriteLine("SetData FileContents")
                Case Functions.RegisterClipboardFormat("Shell IDList Array")
                    Debug.WriteLine("SetData Shell IDList Array")
                Case Functions.RegisterClipboardFormat("Shell Object Offsets")
                    Debug.WriteLine("SetData Shell Object Offsets")
                Case Functions.RegisterClipboardFormat("Net Resource")
                    Debug.WriteLine("SetData Net Resource")
                Case Functions.RegisterClipboardFormat("FileName")
                    Debug.WriteLine("SetData FileName")
                Case Functions.RegisterClipboardFormat("FileNameW")
                    Debug.WriteLine("SetData FileNameW")
                Case Functions.RegisterClipboardFormat("PrinterFriendlyName")
                    Debug.WriteLine("SetData PrinterFriendlyName")
                Case Functions.RegisterClipboardFormat("FileNameMap")
                    Debug.WriteLine("SetData FileNameMap")
                Case Functions.RegisterClipboardFormat("FileNameMapW")
                    Debug.WriteLine("SetData FileNameMapW")
                Case Functions.RegisterClipboardFormat("UniformResourceLocator")
                    Debug.WriteLine("SetData UniformResourceLocator")
                Case Functions.RegisterClipboardFormat("UniformResourceLocatorW")
                    Debug.WriteLine("SetData UniformResourceLocatorW")
                Case Functions.RegisterClipboardFormat("Preferred DropEffect")
                    Debug.WriteLine("SetData Preferred DropEffect")
                Case Functions.RegisterClipboardFormat("Performed DropEffect")
                    Debug.WriteLine("SetData Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Paste Succeeded")
                    Debug.WriteLine("SetData Paste Succeeded")
                Case Functions.RegisterClipboardFormat("InShellDragLoop")
                    Debug.WriteLine("SetData InShellDragLoop")
                Case Functions.RegisterClipboardFormat("MountedVolume")
                    Debug.WriteLine("SetData MountedVolume")
                Case Functions.RegisterClipboardFormat("PersistedDataObject")
                    Debug.WriteLine("SetData PersistedDataObject")
                Case Functions.RegisterClipboardFormat("TargetCLSID")
                    Debug.WriteLine("SetData TargetCLSID")
                Case Functions.RegisterClipboardFormat("Logical Performed DropEffect")
                    Debug.WriteLine("SetData Logical Performed DropEffect")
                Case Functions.RegisterClipboardFormat("Autoplay Enumerated IDList Array")
                    Debug.WriteLine("SetData Autoplay Enumerated IDList Array")
                Case Functions.RegisterClipboardFormat("UntrustedDragDrop")
                    Debug.WriteLine("SetData UntrustedDragDrop")
                Case Functions.RegisterClipboardFormat("File Attributes Array")
                    Debug.WriteLine("SetData File Attributes Array")
                Case Functions.RegisterClipboardFormat("InvokeCommand DropParam")
                    Debug.WriteLine("SetData InvokeCommand DropParam")
                Case Functions.RegisterClipboardFormat("DropHandlerCLSID")
                    Debug.WriteLine("SetData DropHandlerCLSID")
                Case Functions.RegisterClipboardFormat("DropDescription")
                    Debug.WriteLine("SetData DropDescription")
                Case Else
                    Debug.WriteLine("SetData " & format.cfFormat)
            End Select

            If Not (medium.tymed = TYMED.TYMED_HGLOBAL AndAlso medium.unionmember.Equals(IntPtr.Zero)) Then
                Dim fmt2 As FORMATETC = New FORMATETC() With {
                          .cfFormat = format.cfFormat,
                          .dwAspect = format.dwAspect,
                          .lindex = format.lindex,
                          .tymed = format.tymed,
                          .ptd = format.ptd
                      }
                If _data.Keys.ToList().Exists(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                                          AndAlso f.lindex = fmt2.lindex AndAlso (f.tymed And fmt2.tymed)) Then
                    Dim m As Tuple(Of STGMEDIUM, Boolean) = _data.ToList().FirstOrDefault(Function(p) p.Key.cfFormat = fmt2.cfFormat AndAlso p.Key.dwAspect = fmt2.dwAspect _
                                          AndAlso p.Key.lindex = fmt2.lindex AndAlso (p.Key.tymed And fmt2.tymed)).Value
                    If m.Item2 Then
                        Functions.ReleaseStgMedium(m.Item1)
                    End If
                    _data.Remove(_data.Keys.ToList().FirstOrDefault(Function(f) f.cfFormat = fmt2.cfFormat AndAlso f.dwAspect = fmt2.dwAspect _
                                          AndAlso f.lindex = fmt2.lindex AndAlso (f.tymed And fmt2.tymed)))
                End If
                _data.Add(fmt2, New Tuple(Of STGMEDIUM, Boolean)(medium, release))
                Return HRESULT.S_OK
            Else
                Return HRESULT.E_INVALIDARG
            End If
        End Function

        Public Function EnumFormatEtc(ByVal direction As ComTypes.DATADIR, ByRef ienumformatetc As ComTypes.IEnumFORMATETC) As Integer Implements IDataObject_PreserveSig.EnumFormatEtc
            Debug.WriteLine("EnumFormatEtc")
            ienumformatetc = New EnumFORMATETC(_data.Keys.ToArray())
            Return HRESULT.S_OK
        End Function

        Public Function DAdvise(ByRef pFormatetc As ComTypes.FORMATETC, advf As ADVF, adviseSink As ComTypes.IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject_PreserveSig.DAdvise
            Debug.WriteLine("DAdvise")
            Return &H80004001 ' E_NOTIMPL
        End Function

        Public Function DUnadvise(ByVal connection As Integer) As Integer Implements IDataObject_PreserveSig.DUnadvise
            Debug.WriteLine("DUnadvise")
            Return &H80004001 ' E_NOTIMPL
        End Function

        Public Function EnumDAdvise(ByRef enumAdvise As ComTypes.IEnumSTATDATA) As Integer Implements IDataObject_PreserveSig.EnumDAdvise
            Debug.WriteLine("EnumDAdvise")
            Return &H80004001 ' E_NOTIMPL
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True

                For Each item In _data
                    Functions.ReleaseStgMedium(item.Value.Item1)
                Next
                _data.Clear()
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace