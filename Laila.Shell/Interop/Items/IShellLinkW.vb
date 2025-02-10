Imports System.Runtime.InteropServices
Imports System.Text

<ComImport(), Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IShellLinkW
    <PreserveSig()> Function GetPath(pszFile As StringBuilder, cch As Integer, pfd As IntPtr, fFlags As Integer) As Integer
    <PreserveSig()> Function GetIDList(ByRef ppidl As IntPtr) As Integer
    <PreserveSig()> Function SetIDList(pidl As IntPtr) As Integer
    <PreserveSig()> Function GetDescription(pszName As StringBuilder, cch As Integer) As Integer
    <PreserveSig()> Function SetDescription(pszName As String) As Integer
    <PreserveSig()> Function GetWorkingDirectory(pszDir As StringBuilder, cch As Integer) As Integer
    <PreserveSig()> Function SetWorkingDirectory(pszDir As String) As Integer
    <PreserveSig()> Function GetArguments(pszArgs As StringBuilder, cch As Integer) As Integer
    <PreserveSig()> Function SetArguments(pszArgs As String) As Integer
    <PreserveSig()> Function GetHotkey(ByRef pwHotkey As Short) As Integer
    <PreserveSig()> Function SetHotkey(wHotkey As Short) As Integer
    <PreserveSig()> Function GetShowCmd(ByRef piShowCmd As Integer) As Integer
    <PreserveSig()> Function SetShowCmd(iShowCmd As Integer) As Integer
    <PreserveSig()> Function GetIconLocation(pszIconPath As StringBuilder, cch As Integer, ByRef piIcon As Integer) As Integer
    <PreserveSig()> Function SetIconLocation(pszIconPath As String, iIcon As Integer) As Integer
    <PreserveSig()> Function SetRelativePath(pszPathRel As String, dwReserved As Integer) As Integer
    <PreserveSig()> Function Resolve(hWnd As IntPtr, fFlags As Integer) As Integer
    <PreserveSig()> Function SetPath(pszFile As String) As Integer
End Interface