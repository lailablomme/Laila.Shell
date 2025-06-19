Imports System.Runtime.InteropServices

Namespace Helpers
    Public Class OSVersionHelper

        <StructLayout(LayoutKind.Sequential)>
        Private Structure RTL_OSVERSIONINFOEX
            Public dwOSVersionInfoSize As UInteger
            Public dwMajorVersion As UInteger
            Public dwMinorVersion As UInteger
            Public dwBuildNumber As UInteger
            Public dwPlatformId As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
            Public szCSDVersion As String
        End Structure

        <DllImport("ntdll.dll", SetLastError:=True)>
        Private Shared Function RtlGetVersion(ByRef lpVersionInformation As RTL_OSVERSIONINFOEX) As Integer
        End Function

        Private Shared Function GetOSVersion() As (Major As UInteger, Minor As UInteger, Build As UInteger)
            Dim osInfo As New RTL_OSVERSIONINFOEX()
            osInfo.dwOSVersionInfoSize = CUInt(Marshal.SizeOf(osInfo))

            If RtlGetVersion(osInfo) = 0 Then
                Return (osInfo.dwMajorVersion, osInfo.dwMinorVersion, osInfo.dwBuildNumber)
            End If

            Return (0, 0, 0) ' Unknown or error
        End Function

        Private Shared _isWindows11_21H2OrGreater As Boolean? = Nothing
        Public Shared Function IsWindows11_21H2OrGreater() As Boolean
            If Not _isWindows11_21H2OrGreater.HasValue Then
                Dim ver = GetOSVersion()
                _isWindows11_21H2OrGreater = (ver.Major = 10 AndAlso ver.Build >= 22000)
            End If
            Return _isWindows11_21H2OrGreater
        End Function

        Private Shared _isWindows10_1709OrGreater As Boolean? = Nothing
        Public Shared Function IsWindows10_1709OrGreater() As Boolean
            If Not _isWindows10_1709OrGreater.HasValue Then
                Dim ver = GetOSVersion()
                _isWindows10_1709OrGreater = (ver.Major = 10 AndAlso ver.Build >= 16299)
            End If
            Return _isWindows10_1709OrGreater
        End Function

        Private Shared _isWindows81OrLower As Boolean? = Nothing
        Public Shared Function IsWindows81OrLower() As Boolean
            If Not _isWindows81OrLower.HasValue Then
                Dim ver = GetOSVersion()
                ' Windows 8.1 = 6.3
                _isWindows81OrLower = (ver.Major < 6) OrElse (ver.Major = 6 AndAlso ver.Minor <= 3)
            End If
            Return _isWindows81OrLower
        End Function

        Private Shared _isWindows7OrLower As Boolean? = Nothing
        Public Shared Function IsWindows7OrLower() As Boolean
            If Not _isWindows7OrLower.HasValue Then
                Dim ver = GetOSVersion()
                ' Windows 7 = 6.1
                _isWindows7OrLower = (ver.Major < 6) OrElse (ver.Major = 6 AndAlso ver.Minor <= 1)
            End If
            Return _isWindows7OrLower
        End Function
    End Class
End Namespace