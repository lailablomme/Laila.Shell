Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComVisible(True), Guid("e5fdcd58-61e1-4c99-9e7b-94882e93fd58"), ProgId("Laila.Shell.DragDataObject")>
Public Class DragDataObject
    Implements Laila.Shell.IDataObject
    Implements IDisposable

    Private _data As List(Of Tuple(Of FORMATETC, STGMEDIUM, Boolean)) =
        New List(Of Tuple(Of FORMATETC, STGMEDIUM, Boolean))
    Private _toRelease As List(Of STGMEDIUM) = New List(Of STGMEDIUM)()
    Private disposedValue As Boolean
    Const DV_E_FORMATETC As Integer = &H80040064

    Public Sub New()
    End Sub

    Public Function GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer Implements IDataObject.GetData
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

        For Each d In _data
            If compare(d.Item1, format) Then
                Debug.WriteLine("Found")
                Return Functions.CopyStgMedium(d.Item2, medium)
            End If
        Next

        Debug.WriteLine("Not found")
        Return DV_E_FORMATETC
    End Function

    Private Function compare(format As FORMATETC, format2 As FORMATETC) As Boolean
        Return format2.cfFormat = format.cfFormat AndAlso format2.lindex = format.lindex 'AndAlso format2.tymed = format.tymed 'AndAlso format2.dwAspect = format.dwAspect
    End Function

    Public Function GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer Implements IDataObject.GetDataHere
        Debug.WriteLine("GetDataHere")

        For Each d In _data
            If compare(d.Item1, format) Then
                Debug.WriteLine("Found")
                medium.pUnkForRelease = d.Item2.pUnkForRelease
                medium.unionmember = d.Item2.unionmember
                medium.tymed = d.Item2.tymed
                Return 0
            End If
        Next

        Debug.WriteLine("Not found")
        Return DV_E_FORMATETC
    End Function

    Public Function QueryGetData(ByRef format As FORMATETC) As Integer Implements IDataObject.QueryGetData
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

        For Each d In _data
            If compare(d.Item1, format) Then
                Debug.WriteLine("Found")
                Return 0
            End If
        Next

        Debug.WriteLine("Not Found")
        Return &H80040064 ' DV_E_FORMATETC
    End Function

    Public Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer Implements IDataObject.GetCanonicalFormatEtc
        formatOut = formatIn
        Return &H80040130 ' DATA_S_SAMEFORMATETC
    End Function

    Public Function SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) As Integer Implements IDataObject.SetData
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

        For i = 0 To _data.Count - 1
            Dim d As Tuple(Of FORMATETC, STGMEDIUM, Boolean) = _data(i)
            If compare(d.Item1, format) Then
                If d.Item3 Then
                    _toRelease.Add(d.Item2)
                End If
                _data.Remove(d)
                Exit For
            End If
        Next

        _data.Add(New Tuple(Of FORMATETC, STGMEDIUM, Boolean)(format, medium, release))

        Return 0
    End Function

    Public Function EnumFormatEtc(ByVal direction As DATADIR, ByRef ppenumFormatEtc As ComTypes.IEnumFORMATETC) As Integer Implements IDataObject.EnumFormatEtc
        Debug.WriteLine("EnumFormatEtc")
        Dim formats As List(Of FORMATETC) = New List(Of FORMATETC)()
        For Each f In _data.Select(Function(v) v.Item1)
            If Not formats.Exists(Function(fmt) compare(fmt, f)) Then
                formats.Add(f)
            End If
        Next

        ppenumFormatEtc = New EnumFORMATETC(formats.ToArray())
        Return 0
    End Function

    Public Function DAdvise(ByRef pFormatetc As FORMATETC, advf As Integer, adviseSink As IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject.DAdvise
        Debug.WriteLine("DAdvise")
        Return &H80004001 ' E_NOTIMPL
    End Function

    Public Function DUnadvise(ByVal connection As Integer) As Integer Implements IDataObject.DUnadvise
        Debug.WriteLine("DUnadvise")
        Return &H80004001 ' E_NOTIMPL
    End Function

    Public Function EnumDAdvise(ByRef enumAdvise As IEnumSTATDATA) As Integer Implements IDataObject.EnumDAdvise
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
            Try
                For Each m In _toRelease.Union(_data.Select(Function(d) d.Item2))
                    Functions.ReleaseStgMedium(m)
                Next
            Catch ex As Exception
            End Try
            disposedValue = True
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    Protected Overrides Sub Finalize()
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=False)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Public Class EnumFORMATETC
    Implements ComTypes.IEnumFORMATETC

    Private formats As FORMATETC()
    Private currentIndex As Integer

    Public Sub New(formats As FORMATETC())
        Me.formats = formats
        Me.currentIndex = 0
    End Sub

    Public Function Skip(celt As Integer) As Integer Implements ComTypes.IEnumFORMATETC.Skip
        currentIndex += celt
        Return If(currentIndex <= formats.Length, 0, 1) ' S_OK or S_FALSE
    End Function

    Public Function Reset() As Integer Implements ComTypes.IEnumFORMATETC.Reset
        Debug.WriteLine("EnumFORMATETC.Reset")
        currentIndex = 0
    End Function

    Public Sub Clone(ByRef ppenum As ComTypes.IEnumFORMATETC) Implements ComTypes.IEnumFORMATETC.Clone
        ppenum = New EnumFORMATETC(formats)
    End Sub

    Public Function [Next](celt As Integer, rgelt() As FORMATETC, pceltFetched() As Integer) As Integer Implements ComTypes.IEnumFORMATETC.Next
        Debug.WriteLine("EnumFORMATETC.Next " & celt)
        If celt = 64 Then
            Dim i = 9
        End If
        Dim fetched As Integer = 0
        While currentIndex < formats.Length AndAlso fetched < celt
            rgelt(fetched) = formats(currentIndex)
            currentIndex += 1
            fetched += 1
        End While
        'ReDim Preserve pceltFetched(0)
        If Not pceltFetched Is Nothing Then
            pceltFetched(0) = fetched
        End If
        Return If(fetched <> celt, 1, 0) ' S_OK or S_FALSE
    End Function
End Class