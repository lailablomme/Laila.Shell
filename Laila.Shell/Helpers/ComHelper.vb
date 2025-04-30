Imports System.Runtime.InteropServices
Imports System
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.COM

Namespace Helpers
    Public Class ComHelper
        Implements IDisposable

        Private Const DllGetClassObjectName As String = "DllGetClassObject"

        <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
        Private Shared Function LoadLibrary(lpFileName As String) As IntPtr
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetProcAddress(hModule As IntPtr, lpProcName As String) As IntPtr
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function FreeLibrary(hModule As IntPtr) As Boolean
        End Function

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function DllGetClassObjectDelegate(
            ByRef rclsid As Guid,
            ByRef riid As Guid,
            ByRef ppv As IntPtr
        ) As Integer

        <ComImport()>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        <Guid("00000001-0000-0000-C000-000000000046")>
        Private Interface IClassFactory
            <PreserveSig>
            Function CreateInstance(
                <[In]> pUnkOuter As IntPtr,
                <[In]> ByRef riid As Guid,
                <Out> ByRef ppvObject As IntPtr
            ) As Integer

            <PreserveSig>
            Function LockServer(fLock As Boolean) As Integer
        End Interface

        Private Shared ReadOnly IID_IClassFactory As Guid = New Guid("00000001-0000-0000-C000-000000000046")

        Private _objects As List(Of Object) = New List(Of Object)()
        Private _modules As List(Of IntPtr) = New List(Of IntPtr)()
        Private disposedValue As Boolean

        Public Function MakeComObject(dllPath As String, clsid As Guid, interfaceId As Guid) As Object
            Dim ptr As IntPtr = IntPtr.Zero
            Functions.CoCreateInstance(clsid, IntPtr.Zero, ClassContext.LocalServer, interfaceId, ptr)
            Return Marshal.GetObjectForIUnknown(ptr)

            Dim hModule As IntPtr = LoadLibrary(dllPath)
            _modules.Add(hModule)
            If hModule = IntPtr.Zero Then
                Throw New Exception($"Failed to LoadLibrary: {dllPath} (Win32 error {Marshal.GetLastWin32Error()})")
            End If

            Dim procAddress As IntPtr = GetProcAddress(hModule, DllGetClassObjectName)
            If procAddress = IntPtr.Zero Then
                Throw New Exception("DllGetClassObject not found in DLL.")
            End If

            Dim dllGetClassObject = CType(Marshal.GetDelegateForFunctionPointer(procAddress, GetType(DllGetClassObjectDelegate)), DllGetClassObjectDelegate)

            Dim classFactoryPtr As IntPtr, classFactory As IClassFactory = Nothing
            Dim hr = dllGetClassObject(clsid, IID_IClassFactory, classFactoryPtr)
            If hr <> 0 OrElse classFactoryPtr = IntPtr.Zero Then
                Throw New COMException($"DllGetClassObject failed with HRESULT 0x{hr:X8}", hr)
            End If

            Try
                classFactory = CType(Marshal.GetUniqueObjectForIUnknown(classFactoryPtr), IClassFactory)

                Dim objectPtr As IntPtr
                hr = Functions.CoCreateInstance(clsid, IntPtr.Zero, ClassContext.LocalServer, interfaceId, objectPtr)
                '  hr = classFactory.CreateInstance(IntPtr.Zero, interfaceId, objectPtr)
                If hr <> 0 OrElse objectPtr = IntPtr.Zero Then
                    Throw New COMException($"CreateInstance failed with HRESULT 0x{hr:X8}", hr)
                End If

                ' Now safely get a managed RCW for the final COM object
                Dim comObject = Marshal.GetObjectForIUnknown(objectPtr)
                _objects.Add(comObject)
                Return comObject
            Finally
                Marshal.Release(classFactoryPtr)
                If Not classFactory Is Nothing Then Marshal.ReleaseComObject(classFactory)
            End Try
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                For Each obj As Object In _objects
                    Marshal.ReleaseComObject(obj)
                Next
                For Each hModule As IntPtr In _modules
                    FreeLibrary(hModule)
                Next

                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
        ' Protected Overrides Sub Finalize()
        '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        '     Dispose(disposing:=False)
        '     MyBase.Finalize()
        ' End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace