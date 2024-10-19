Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure SYSTEMTIME
    Public wYear As UShort
    Public wMonth As UShort
    Public wDayOfWeek As UShort
    Public wDay As UShort
    Public wHour As UShort
    Public wMinute As UShort
    Public wSecond As UShort
    Public wMilliseconds As UShort
End Structure