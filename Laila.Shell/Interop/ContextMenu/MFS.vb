﻿Namespace Interop.ContextMenu
    <Flags>
    Public Enum MFS As Integer
        MFS_GRAYED = &H3
        MFS_DISABLED = MFS_GRAYED
        MFS_CHECKED = &H8
        MFS_HILITE = &H80
        MFS_ENABLED = &H0
        MFS_UNCHECKED = &H0
        MFS_UNHILITE = &H0
        MFS_DEFAULT = &H1000
    End Enum
End Namespace
