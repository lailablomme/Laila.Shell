Namespace Interop.Items
    <Flags>
    Public Enum AssocF As UInteger
        None = 0
        InitNoRemapCLSID = &H1
        InitByExeName = &H2
        OpenByExeName = &H2
        InitDefaultToStar = &H4
        InitDefaultToFolder = &H8
        NoUserSettings = &H10
        NoTruncate = &H20
        Verify = &H40
        RemapRunDll = &H80
        NoFixUps = &H100
        IgnoreBaseClass = &H200
    End Enum
End Namespace
