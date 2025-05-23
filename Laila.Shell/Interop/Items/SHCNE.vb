﻿Namespace Interop.Items
    <Flags()>
    Public Enum SHCNE As Int32
        RENAMEITEM = &H1
        CREATE = &H2
        DELETE = &H4
        MKDIR = &H8
        RMDIR = &H10
        MEDIAINSERTED = &H20
        MEDIAREMOVED = &H40
        DRIVEREMOVED = &H80
        DRIVEADD = &H100
        NETSHARE = &H200
        NETUNSHARE = &H400
        ATTRIBUTES = &H800
        UPDATEDIR = &H1000
        UPDATEITEM = &H2000
        SERVERDISCONNECT = &H4000
        UPDATEIMAGE = &H8000
        DRIVEADDGUI = &H10000
        RENAMEFOLDER = &H20000
        FREESPACE = &H40000
        EXTENDED_EVENT = &H4000000
        ASSOCCHANGED = &H8000000
        DISKEVENTS = &H2381F
        GLOBALEVENTS = &HC0581E0
        ALLEVENTS = &H7FFFFFFF
        INTERRUPT = &H80000000
    End Enum
End Namespace
