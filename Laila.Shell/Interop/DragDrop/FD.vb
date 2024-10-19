<Flags>
Public Enum FD As UInteger
    FD_ATTRIBUTES = &H4        ' Indicates that the dwFileAttributes member is valid
    FD_FILESIZE = &H40         ' Indicates that the nFileSizeLow and nFileSizeHigh members are valid
    FD_WRITESTIME = &H8        ' Indicates that the ftLastWriteTime member is valid
    FD_CREATETIME = &H2        ' Indicates that the ftCreationTime member is valid
    FD_ACCESSTIME = &H10       ' Indicates that the ftLastAccessTime member is valid
    FD_UNICODE = &H80000000UI  ' Indicates that the cFileName is Unicode
    FD_PROGRESSUI = &H4000     ' Flag not typically used in file descriptors
End Enum