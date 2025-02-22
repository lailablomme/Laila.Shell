Namespace Interop.Storage
    <Flags>
    Public Enum STGM
        STGM_CREATE = &H1000
        STGM_CONVERT = &H2000
        STGM_FAILIFTHERE = &H0
        STGM_SHARE_EXCLUSIVE = &H10
        STGM_READ = &H0
        STGM_WRITE = &H1
        STGM_READWRITE = &H2
        STGM_SHARE_DENY_WRITE = &H20
        STGM_SHARE_DENY_READ = &H30
        STGM_SHARE_DENY_NONE = &H40
    End Enum
End Namespace
