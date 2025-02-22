Namespace Interop.COM
    <Flags>
    Public Enum ClassContext As UInteger
        ''' <summary>
        ''' The code that manages objects runs in the same process as the caller.
        ''' </summary>
        InProcServer = 1

        ''' <summary>
        ''' The object is managed by a DLL server that runs in a new process.
        ''' </summary>
        InProcHandler = 2

        ''' <summary>
        ''' The object is managed by a DLL server running in the same process as the caller.
        ''' </summary>
        LocalServer = 4

        ''' <summary>
        ''' The object is managed by a remote server on another computer.
        ''' </summary>
        RemoteServer = 16

        ''' <summary>
        ''' The object is managed by a handler DLL and runs in a new process.
        ''' </summary>
        InProcHandler16 = 32

        ''' <summary>
        ''' The object is managed by a 16-bit executable.
        ''' </summary>
        RemoteServer16 = 64

        ''' <summary>
        ''' The object is run in the same process as the caller.
        ''' </summary>
        InProcServer16 = 128

        ''' <summary>
        ''' Enable Activate At Storage.
        ''' </summary>
        ActivateAtStorage = 512

        ''' <summary>
        ''' Reserved.
        ''' </summary>
        EnableCodeDownload = 1024

        ''' <summary>
        ''' Reserved.
        ''' </summary>
        NoCustomMarshal = 2048

        ''' <summary>
        ''' Specify that a new object is created.
        ''' </summary>
        EnableAsync = 4096

        ''' <summary>
        ''' Reserved.
        ''' </summary>
        Reserved = 8192

        ''' <summary>
        ''' Run in a 32-bit server.
        ''' </summary>
        InProcHandler32 = 16384

        ''' <summary>
        ''' Run in a 32-bit server.
        ''' </summary>
        InProcServer32 = 32768

        ''' <summary>
        ''' The object is managed by an out-of-process server that supports multithreading apartments.
        ''' </summary>
        MultiThreaded = 131072

        ''' <summary>
        ''' Run in a 64-bit server.
        ''' </summary>
        InProcServer64 = 65536

        ''' <summary>
        ''' Run in a 64-bit handler.
        ''' </summary>
        InProcHandler64 = 262144
    End Enum
End Namespace
