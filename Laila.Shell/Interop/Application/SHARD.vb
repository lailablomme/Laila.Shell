Namespace Interop.Application
    Public Enum SHARD
        SHARD_PIDL = &H1        ' The pv parameter points to an ITEMIDLIST structure
        SHARD_PATHA = &H2       ' The pv parameter points to a null-terminated ANSI string with a file path
        SHARD_PATHW = &H3       ' The pv parameter points to a null-terminated Unicode string with a file path
        SHARD_APPIDINFO = &H4   ' The pv parameter points to an APPIDINFO structure (used for AppUserModelID)
        SHARD_APPIDINFOIDLIST = &H5 ' The pv parameter points to an APPIDINFOIDLIST structure
        SHARD_LINK = &H6        ' The pv parameter points to an IShellLink object
        SHARD_APPIDINFOLINK = &H7 ' The pv parameter points to an APPIDINFOLINK structure
        SHARD_SHELLITEM = &H8   ' The pv parameter points to an IShellItem object
    End Enum
End Namespace
