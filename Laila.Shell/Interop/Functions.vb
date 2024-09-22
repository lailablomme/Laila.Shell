Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Shapes
Imports System.Runtime.InteropServices.ComTypes

Public Class Functions
    Public Const STGM_READWRITE As Integer = 2
    Public Const STR_ENUM_ITEMS_FLAGS As String = "EnumItemsFlags"

    '<DllImport("propsys.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    'Public Shared Function PSCreateMemoryPropertyStore(
    '    ByRef riid As Guid,
    '    ByRef ppv As IntPtr
    ') As Integer
    'End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateDataObject(pidlFolder As IntPtr, cidl As UInteger, apidl As IntPtr(), pdtInner As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IDataObject) As Integer
    End Function
    <DllImport("ole32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function CoCreateBindCtx(dwFlags As UInt32, reserved As UInt32, pMalloc As IntPtr, ByRef ppbc As IBindCtx) As Integer
    End Function

    <DllImport("gdi32.dll", SetLastError:=True)>
    Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function
    <DllImport("ole32.dll", SetLastError:=True)>
    Public Shared Function PropVariantClear(ByRef pvar As PROPVARIANT) As Integer
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHGetDesktopFolder(ByRef ppshf As IShellFolder) As Integer
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHGetKnownFolderItem(<[In]> rfid As Guid, <[In]> flags As UInteger, <[In]> hToken As IntPtr, <[In]> riid As Guid, <Out> ByRef ppv As IShellItem) As HResult
    End Function

    <DllImport("shlwapi.dll")>
    Public Shared Function StrRetToBuf(ByVal pstr As IntPtr, ByVal pidl As IntPtr, ByVal pszBuf As StringBuilder, <MarshalAs(UnmanagedType.U4)> ByVal cchBuf As Integer) As Integer
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function PathParseIconLocationA(<[In], Out, MarshalAs(UnmanagedType.LPStr)> ByRef pszIconFile As String) As Integer
    End Function

    <DllImport("Shell32.dll", EntryPoint:="ExtractIconExW", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function ExtractIconEx(sFile As String, iIndex As Integer, ByRef piLargeVersion As IntPtr, ByRef piSmallVersion As IntPtr, amountIcons As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetIconInfo(ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHCreateItemWithParent(pidlParent As IntPtr, psfParent As IShellFolder, pidl As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CreateBindCtx(reserved As UInt32, ByRef ppbc As IntPtr) As Integer
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreatePropertyBagOnMemory(
    grfMode As Integer,
    ByRef riid As Guid,
    <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyBag) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateShellItemArrayFromShellItem(
        <MarshalAs(UnmanagedType.Interface)> psi As IShellItem,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IShellItemArray
    ) As Integer

    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreateItemFromParsingName(pszPath As String, pbc As IntPtr, riid As Guid, ByRef ppv As IntPtr) As HResult
    End Function

    <DllImport("shell32.dll")>
    Public Shared Sub SHParseDisplayName(<MarshalAs(UnmanagedType.LPWStr)> name As String, bindingContext As IntPtr, <Out()> ByRef pidl As IntPtr, sfgaoIn As UInt32, <Out()> ByRef sfgaoOut As UInt32)
    End Sub

    <DllImport("shell32.dll")>
    Public Shared Function SHCreateItemFromIDList(<[In]> pidl As IntPtr, <[In]> riid As Guid, <Out> ByRef ppv As IntPtr) As HResult
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function ILRemoveLastID(<[In]> <Out> ByRef pidl As IntPtr) As Boolean
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSGetPropertyDescription(propKey As PROPERTYKEY, riid As Guid, ByRef ppv As IntPtr) As Long
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSGetImageReferenceForValue(<[In]> ByRef propKey As PROPERTYKEY, <[In]> ByRef propvar As PROPVARIANT, <Out> ByRef ppszImageRes As IntPtr) As Long
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSCreateMemoryPropertyStore(<[In]> riid As Guid, <Out> ByRef ppv As IntPtr) As HResult
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSFormatForDisplay(<[In]> ByRef key As PROPERTYKEY, <[In]> ByRef propvar As PROPVARIANT, pdff As PropertyDescriptionFormatOptions, <[Out]> ppszDisplay() As Char) As Integer

    End Function

    <DllImport("propsys.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function PSGetPropertyDescription(
            ByRef propkey As PROPERTYKEY,
            ByRef riid As Guid,
            <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyDescription
        ) As HResult
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function SHLoadPropertyStoreFromBlob(pBlob As PROPVARIANT, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppPropStore As IPropertyStore) As HResult
    End Function

    <DllImport("shlwapi.dll", SetLastError:=True)>
    Public Shared Function SHCreateMemStream(pvBuf As IntPtr, ByVal cbSize As Long) As IntPtr
    End Function

    <DllImport("shell32.dll", SetLastError:=True)>
    Public Shared Function SHGetPropertyStoreForFile(
        ByVal pszFilePath As String,
        ByVal grfFlags As UInt32,
        ByRef ppPropStore As IntPtr
    ) As HResult
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertyStoreForFile(
            <MarshalAs(UnmanagedType.LPWStr)> pszFilePath As String,
             grfFlags As Integer,
            <MarshalAs(UnmanagedType.Interface)> ByRef ppPropStore As IPropertyStore
        ) As HResult
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertySystem(
  <[In]> riid As Guid,
  <[Out]> ByRef ppv As IntPtr
) As HResult
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertyDescriptionByName(
    <[In], MarshalAs(UnmanagedType.LPWStr)> pszCanonicalName As String,
  <[In]> riid As Guid,
  <[Out], MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyDescription
) As HResult

    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMessage(ByRef lpMsg As MSG, ByVal hWnd As IntPtr, ByVal wMsgFilterMin As UInteger, ByVal wMsgFilterMax As UInteger) As Boolean
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function SHGetIDListFromObject(ByVal punk As IntPtr, ByRef ppidl As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function ILFindLastID(pidl As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function CreatePopupMenu() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuItemCount(hMenu As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function Shell_GetImageLists(ByRef phimlLarge As IntPtr, ByRef phimlSmall As IntPtr) As Integer
    End Function

    <DllImport("comctl32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ImageList_GetIcon(himl As IntPtr, i As Integer, flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuString(hMenu As IntPtr, uIDItem As UInteger, lpString As StringBuilder, nMaxCount As Integer, uFlag As UInteger) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuItemInfo(hMenu As IntPtr, uItem As UInteger, fByPosition As Boolean, ByRef lpmii As MENUITEMINFO) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyMenu(ByVal hMenu As IntPtr) As Boolean
    End Function

    Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As IntPtr,
           ByVal uFlags As Integer, ByVal x As Integer, ByVal y As Integer,
           ByVal hWnd As IntPtr, ByVal lptpm As IntPtr) As Integer
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Sub mouse_event(ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal dwData As Integer, ByVal dwExtraInfo As IntPtr)
    End Sub
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function EndMenu() As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetMenu(ByVal hWnd As IntPtr, ByVal hMenu As IntPtr) As Boolean
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetPathFromIDList(ByVal pidl As IntPtr, ByVal pszPath As System.Text.StringBuilder) As Boolean
    End Function

    <DllImport("shell32", EntryPoint:="#2", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotifyRegister(hwnd As IntPtr, fSources As SHCNRF, fEvents As SHCNE, wMsg As WM, cEntries As Integer, <MarshalAs(UnmanagedType.LPArray)> pfsne As SHChangeNotifyEntry()) As Integer
    End Function

    <DllImport("shell32", EntryPoint:="#4", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotifyDeregister(hNotify As Integer) As Boolean
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotification_Lock(ByVal hChange As IntPtr, ByVal dwProcId As UInt32, ByRef pppidl As IntPtr, ByRef plEvent As SHCNE) As IntPtr
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotification_Unlock(ByVal hLock As IntPtr) As Int32
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Sub SHChangeNotify(ByVal wEventId As Integer, ByVal uFlags As Integer, ByVal dwItem1 As IntPtr, ByVal dwItem2 As IntPtr)
    End Sub
End Class
