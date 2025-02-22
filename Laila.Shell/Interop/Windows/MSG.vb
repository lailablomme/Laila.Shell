Imports System.Runtime.InteropServices

Namespace Interop.Windows
    <StructLayout(LayoutKind.Sequential)>
    Public Structure MSG
        Public hwnd As IntPtr
        Public message As UInteger
        Public wParam As IntPtr
        Public lParam As IntPtr
        Public time As UInteger
        Public pt As WIN32POINT
    End Structure
End Namespace
