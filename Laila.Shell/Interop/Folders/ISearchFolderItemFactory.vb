Imports System.Runtime.InteropServices
Imports Microsoft.Xaml.Behaviors.Core

<ComImport>
<Guid("A0FFBC28-5482-4366-BE27-3E81E78E06C2")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface ISearchFolderItemFactory
    Sub SetDisplayName(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszDisplayName As String)

    Sub SetFolderTypeID(ByRef ftid As Guid)

    Sub SetFolderLogicalViewType(ByVal flvt As Integer)

    Sub SetIconSize(ByVal iIconSize As Integer)

    Sub SetVisibleColumns(ByVal cVisibleColumns As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgKey As PROPERTYKEY())

    Sub SetSortColumns(ByVal cSortColumns As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgSortColumns As SORTCOLUMN())

    Sub SetGroupColumn(ByRef keyGroup As PROPERTYKEY)

    Sub SetStacks(ByVal cStackKeys As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgStackKeys As PROPERTYKEY())

    Sub SetScope(<MarshalAs(UnmanagedType.Interface)> ByVal psiaScope As IShellItemArray)

    Sub SetCondition(<MarshalAs(UnmanagedType.Interface)> ByVal pCondition As ICondition)

    Sub GetShellItem(ByVal riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object)

    Sub GetIDList(ByVal flags As Integer, <MarshalAs(UnmanagedType.Interface)> ByRef ppidl As IntPtr)
End Interface

