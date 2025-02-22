Namespace Interop.System
    <Flags>
    Public Enum SSF As UInt32
        SSF_SHOWALLOBJECTS = &H1             ' Show all objects
        SSF_SHOWEXTENSIONS = &H2            ' Show file name extensions
        SSF_HIDDENFILEEXTENSIONS = &H4      ' Hide extensions for known file types
        SSF_SHOWSYSFILES = &H20             ' Show system files
        SSF_DOUBLECLICKINWEBVIEW = &H80     ' Double-click in web view
        SSF_SHOWSUPERHIDDEN = &H40000         ' Show super-hidden files
        SSF_SEPPROCESS = &H400              ' Separate process for folders
        SSF_AUTOCHECKSELECT = &H800000       ' Auto-check select
        SSF_ICONSONLY = &H1000000            ' Icons only, no thumbnails
        SSF_SHOWTYPEOVERLAY = &H2000000       ' Show type overlay
        SSF_SHOWSTATUSBAR = &H4000000         ' Show status bar
        SSF_SHOWCOMPCOLOR = &H8
        SSF_SHOWINFOTIP = &H2000
    End Enum
End Namespace
