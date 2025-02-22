﻿Namespace Interop.Windows
    <Flags()>
    Public Enum WS As UInteger
        ' Window styles
        WS_OVERLAPPED = &H0
        'WS_POPUP = &H80000000
        WS_CHILD = &H40000000
        WS_MINIMIZE = &H20000000
        WS_VISIBLE = &H10000000
        WS_DISABLED = &H8000000
        WS_CLIPSIBLINGS = &H4000000
        WS_CLIPCHILDREN = &H2000000
        WS_MAXIMIZE = &H1000000
        WS_BORDER = &H800000
        WS_DLGFRAME = &H400000
        WS_VSCROLL = &H200000
        WS_HSCROLL = &H100000
        WS_SYSMENU = &H80000
        WS_THICKFRAME = &H40000
        WS_GROUP = &H20000
        WS_TABSTOP = &H10000
        WS_MINIMIZEBOX = &H20000
        WS_MAXIMIZEBOX = &H10000

        ' Extended window styles
        WS_EX_TOOLWINDOW = &H80
        WS_EX_APPWINDOW = &H40000
        WS_EX_CLIENTEDGE = &H200
        WS_EX_CONTEXTHELP = &H400
        WS_EX_DLGMODALFRAME = &H1
        WS_EX_LAYERED = &H80000
        WS_EX_NOINHERITLAYOUT = &H100000
        WS_EX_NOREDIRECTIONBITMAP = &H8000000
        ' WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
        ' WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST

        ' Special styles for certain windows
        WS_EX_TOPMOST = &H8
        WS_EX_TRANSPARENT = &H20
        WS_EX_RIGHT = &H1000
        WS_EX_RTLREADING = &H2000
        WS_EX_LEFTSCROLLBAR = &H4000
        WS_EX_COMPOSITED = &H2000000
        WS_EX_NOACTIVATE = &H8000000
    End Enum
End Namespace
