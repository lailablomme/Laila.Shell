Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComVisible(True), Guid("e5fdcd58-61e1-4c99-9e7b-94882e93fd58"), ProgId("Laila.Shell.DataObject")>
Public Class DragDataObject
    Implements ComTypes.IDataObject
    Implements IDisposable

    Private _data As List(Of Tuple(Of FORMATETC, STGMEDIUM, Boolean)) =
        New List(Of Tuple(Of FORMATETC, STGMEDIUM, Boolean))
    Private _toRelease As List(Of STGMEDIUM) = New List(Of STGMEDIUM)()
    Private disposedValue As Boolean

    Public Sub New()
    End Sub

    Public Sub GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetData
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
            Case Else
                Debug.WriteLine("GetData " & format.cfFormat)
        End Select

        Dim item As Tuple(Of FORMATETC, STGMEDIUM, Boolean)
        For Each d In _data
            If compare(d.Item1, format) Then
                item = d
                Exit For
            End If
        Next
        If Not item Is Nothing Then
            Functions.CopyStgMedium(item.Item2, medium)
            Debug.WriteLine("Found")
        Else
            Debug.WriteLine("Not found")
        End If
    End Sub

    Private Function compare(format As FORMATETC, format2 As FORMATETC) As Boolean
        Return format2.cfFormat = format.cfFormat AndAlso format2.dwAspect = format.dwAspect AndAlso format2.tymed = format.tymed
    End Function

    Public Sub GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetDataHere
    End Sub

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
            Case Else
                Debug.WriteLine("QueryGetData " & format.cfFormat)
        End Select

        Dim item As Tuple(Of FORMATETC, STGMEDIUM, Boolean)
        For Each d In _data
            If compare(d.Item1, format) Then
                item = d
                Exit For
            End If
        Next
        If Not item Is Nothing Then
            Return 0 ' S_OK
        Else
            Return &H80040064 ' DV_E_FORMATETC
        End If
    End Function

    Public Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer Implements IDataObject.GetCanonicalFormatEtc
        formatOut = formatIn
        Return &H80040130 ' DATA_S_SAMEFORMATETC
    End Function

    Public Sub SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) Implements IDataObject.SetData
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
            Case Else
                Debug.WriteLine("SetData " & format.cfFormat)
        End Select

        '_internal.SetData(format, medium, True)

        Dim item As Tuple(Of FORMATETC, STGMEDIUM, Boolean)
        For Each d In _data
            If compare(d.Item1, format) Then
                item = d
                Exit For
            End If
        Next
        If Not item Is Nothing Then
            Dim m As STGMEDIUM = item.Item2
            If item.Item3 Then
                _toRelease.Add(m)
            End If
            _data.Remove(item)
        End If
        _data.Add(New Tuple(Of FORMATETC, STGMEDIUM, Boolean)(format, medium, release))
    End Sub

    Public Function EnumFormatEtc(ByVal direction As DATADIR) As IEnumFORMATETC Implements IDataObject.EnumFormatEtc
        Debug.WriteLine("EnumFormatEtc")
        Return New EnumFORMATETC(_data.Select(Function(v) v.Item1).ToArray())
    End Function

    Public Function DAdvise(ByRef pFormatetc As FORMATETC, advf As ADVF, adviseSink As IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject.DAdvise
        Debug.WriteLine("DAdvise")
        Return &H80004001 ' E_NOTIMPL
    End Function

    Public Sub DUnadvise(ByVal connection As Integer) Implements IDataObject.DUnadvise
        Debug.WriteLine("DUnadvise")
    End Sub

    Public Function EnumDAdvise(ByRef enumAdvise As IEnumSTATDATA) As Integer Implements IDataObject.EnumDAdvise
        Debug.WriteLine("EnumDAdvise")
        Return Nothing
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
    Implements IEnumFORMATETC

    Private formats As FORMATETC()
    Private currentIndex As Integer

    Public Sub New(formats As FORMATETC())
        Me.formats = formats
        Me.currentIndex = 0
    End Sub

    Public Function [Next](celt As Integer, rgelt() As FORMATETC, pceltFetched() As Integer) As Integer Implements IEnumFORMATETC.Next
        Dim fetched As Integer = 0
        While currentIndex < formats.Length AndAlso fetched < celt
            rgelt(fetched) = formats(currentIndex)
            currentIndex += 1
            fetched += 1
        End While
        ReDim pceltFetched(0)
        pceltFetched(0) = fetched
        Return If(fetched = celt, 0, 1) ' S_OK or S_FALSE
    End Function

    Public Function Skip(celt As Integer) As Integer Implements IEnumFORMATETC.Skip
        currentIndex += celt
        Return If(currentIndex <= formats.Length, 0, 1) ' S_OK or S_FALSE
    End Function

    Public Function Reset() As Integer Implements IEnumFORMATETC.Reset
        currentIndex = 0
    End Function

    Public Sub Clone(ByRef ppenum As IEnumFORMATETC) Implements IEnumFORMATETC.Clone
        ppenum = New EnumFORMATETC(formats)
    End Sub
End Class