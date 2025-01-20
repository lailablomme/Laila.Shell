Imports System.Runtime.InteropServices
Imports Microsoft.Xaml.Behaviors.Core

<ComImport>
<Guid("A0FFBC28-5482-4366-BE27-3E81E78E06C2")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface ISearchFolderItemFactory
    <PreserveSig>
    Function SetDisplayName(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszDisplayName As String) As HRESULT
    <PreserveSig>
    Function SetFolderTypeID(ByRef ftid As Guid) As HRESULT
    <PreserveSig>
    Function SetFolderLogicalViewMode(ByVal flvt As Integer) As HRESULT
    <PreserveSig>
    Function SetIconSize(ByVal iIconSize As Integer) As HRESULT
    <PreserveSig>
    Function SetVisibleColumns(ByVal cVisibleColumns As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgKey As PROPERTYKEY()) As HRESULT
    <PreserveSig>
    Function SetSortColumns(ByVal cSortColumns As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgSortColumns As SortColumn()) As HRESULT
    <PreserveSig>
    Function SetGroupColumn(ByRef keyGroup As PROPERTYKEY) As HRESULT
    <PreserveSig>
    Function SetStacks(ByVal cStackKeys As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgStackKeys As PROPERTYKEY()) As HRESULT
    <PreserveSig>
    Function SetScope(<MarshalAs(UnmanagedType.Interface)> ByVal psiaScope As IShellItemArray) As HRESULT
    <PreserveSig>
    Function SetCondition(<MarshalAs(UnmanagedType.Interface)> ByVal pCondition As ICondition) As HRESULT
    <PreserveSig>
    Function GetShellItem(ByVal riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As HRESULT
    <PreserveSig>
    Function GetIDList(ByRef ppidl As IntPtr) As HRESULT
End Interface

