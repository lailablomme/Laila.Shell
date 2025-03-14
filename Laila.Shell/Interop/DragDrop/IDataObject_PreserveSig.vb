Imports System
Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.DragDrop
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization

Namespace Interop.DragDrop
    <ComImport, Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDataObject_PreserveSig
        <PreserveSig>
        Function GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer

        <PreserveSig>
        Function GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer
        <PreserveSig>
        Function QueryGetData(ByRef format As FORMATETC) As Integer
        <PreserveSig>
        Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer
        <PreserveSig>
        Function SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, fRelease As Boolean) As Integer
        <PreserveSig>
        Function EnumFormatEtc(dwDirection As ComTypes.DATADIR, <Out> ByRef ienumFormatEtc As IEnumFORMATETC) As Integer
        <PreserveSig>
        Function DAdvise(ByRef format As FORMATETC, advf As ADVF, pAdvSink As IAdviseSink, ByRef connection As Integer) As Integer

        <PreserveSig>
        Function DUnadvise(connection As Integer) As Integer

        <PreserveSig>
        Function EnumDAdvise(<Out> ByRef enumAdvise As IEnumSTATDATA) As Integer
    End Interface
End Namespace
