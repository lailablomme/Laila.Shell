Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport, Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IDataObject
    <PreserveSig>
    Function GetData(ByRef formatetcIn As FORMATETC, <Out> ByRef medium As STGMEDIUM) As Integer

    <PreserveSig>
    Function GetDataHere(ByRef formatetc As FORMATETC, <[In], Out> ByRef medium As STGMEDIUM) As Integer

    <PreserveSig>
    Function QueryGetData(ByRef formatetc As FORMATETC) As Integer

    <PreserveSig>
    Function GetCanonicalFormatEtc(ByRef formatetcIn As FORMATETC, <Out> ByRef formatetcOut As FORMATETC) As Integer

    <PreserveSig>
    Function SetData(ByRef formatetc As FORMATETC, ByRef medium As STGMEDIUM, fRelease As Boolean) As Integer

    <PreserveSig>
    Function EnumFormatEtc(ByVal direction As DATADIR, ByRef ppenumFormatEtc As ComTypes.IEnumFORMATETC) As Integer

    <PreserveSig>
    Function DAdvise(ByRef formatetc As FORMATETC, advf As Integer, <[In], MarshalAs(UnmanagedType.Interface)> pAdvSink As IAdviseSink, <Out> ByRef pdwConnection As Integer) As Integer

    <PreserveSig>
    Function DUnadvise(dwConnection As Integer) As Integer

    <PreserveSig>
    Function EnumDAdvise(<Out> ByRef ppenumAdvise As IEnumSTATDATA) As Integer
End Interface

<ComImport, Guid("00000103-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IEnumFORMATETC
    <PreserveSig>
    Function [Next](celt As Integer, <Out> rgelt() As FORMATETC, <Out> ByRef pceltFetched() As Integer) As Integer

    <PreserveSig>
    Function Skip(celt As Integer) As Integer

    <PreserveSig>
    Function Reset() As Integer

    <PreserveSig>
    Function Clone(<Out> ByRef ppenum As IEnumFORMATETC) As Integer
End Interface

<ComImport, Guid("0000010F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IAdviseSink
    Sub OnDataChange(<[In]> ByRef format As FORMATETC, <[In]> ByRef stgmedium As STGMEDIUM)
    Sub OnViewChange(dwAspect As Integer, lindex As Integer)
    Sub OnRename(<MarshalAs(UnmanagedType.Interface)> pmk As IMoniker)
    Sub OnSave()
    Sub OnClose()
End Interface

<ComImport, Guid("00000105-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IEnumSTATDATA
    <PreserveSig>
    Function [Next](celt As Integer, <Out> rgelt() As STATDATA, <Out> ByRef pceltFetched As Integer) As Integer

    <PreserveSig>
    Function Skip(celt As Integer) As Integer

    <PreserveSig>
    Function Reset() As Integer

    <PreserveSig>
    Function Clone(<Out> ByRef ppenum As IEnumSTATDATA) As Integer
End Interface

<StructLayout(LayoutKind.Sequential)>
Public Structure STATDATA
    Public formatetc As FORMATETC
    Public advf As Integer
    <MarshalAs(UnmanagedType.Interface)>
    Public pAdvSink As IAdviseSink
    Public dwConnection As Integer
End Structure
