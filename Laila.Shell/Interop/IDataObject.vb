'' Structure representing FORMATETC
'Imports System.Runtime.InteropServices
'Imports System.Runtime.InteropServices.ComTypes

'<ComImport(), Guid("00000103-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'Public Interface IEnumFORMATETC
'    <PreserveSig>
'    Function [Next](celt As Integer, rgelt() As FORMATETC, ByRef pceltFetched As Integer) As Integer

'    <PreserveSig>
'    Function Skip(ByVal celt As Integer) As Integer

'    <PreserveSig>
'    Function Reset() As Integer

'    <PreserveSig>
'    Function Clone(<Out> ByRef ppenum As IEnumFORMATETC) As Integer
'End Interface
'<StructLayout(LayoutKind.Sequential)>
'Public Structure FORMATETC
'    Public cfFormat As Integer
'    Public ptd As IntPtr
'    Public dwAspect As Integer
'    Public lindex As Integer
'    Public tymed As Integer
'End Structure

'' Structure representing STGMEDIUM
'<StructLayout(LayoutKind.Sequential)>
'Public Structure STGMEDIUM
'    Public tymed As Integer
'    Public hGlobal As IntPtr
'    Public pUnkForRelease As IntPtr
'End Structure
'<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000010E-0000-0000-C000-000000000046")>
'Public Interface IDataObject
'    <PreserveSig>
'    Function GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer

'    <PreserveSig>
'    Function GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer

'    <PreserveSig>
'    Function QueryGetData(ByRef format As FORMATETC) As Integer

'    <PreserveSig>
'    Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer

'    <PreserveSig>
'    Function SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) As Integer

'    <PreserveSig>
'    Function EnumFormatEtc(ByVal direction As DATADIR, <Out()> ByRef enumFormatEtc2 As IEnumFORMATETC) As Integer

'    <PreserveSig>
'    Function DAdvise(ByRef format As FORMATETC, ByVal advf As ADVF, ByVal adviseSink As IAdviseSink, ByRef connection As Integer) As Integer

'    <PreserveSig>
'    Function DUnadvise(ByVal connection As Integer) As Integer

'    <PreserveSig>
'    Function EnumDAdvise(<Out()> ByRef enumAdvise As IEnumSTATDATA) As Integer
'End Interface