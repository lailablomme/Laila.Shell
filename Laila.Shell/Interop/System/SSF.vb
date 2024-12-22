<Flags>
Public Enum SSF As UInt32
    SSF_SHOWALLOBJECTS = &H1             ' Show all objects
    SSF_SHOWEXTENSIONS = &H2            ' Show file name extensions
    SSF_HIDDENFILEEXTENSIONS = &H4      ' Hide extensions for known file types
    SSF_SHOWSYSFILES = &H20             ' Show system files
    SSF_DOUBLECLICKINWEBVIEW = &H80     ' Double-click in web view
    SSF_SHOWSUPERHIDDEN = &H40000         ' Show super-hidden files
    SSF_SEPPROCESS = &H400              ' Separate process for folders
    SSF_NONETCRAWLING = &H800           ' No network crawling
    SSF_STARTPANELON = &H4000           ' Start panel on
    SSF_SHOWSTARTPAGE = &H8000          ' Show start page
    SSF_AUTOCHECKSELECT = &H800000       ' Auto-check select
    SSF_ICONSONLY = &H20000             ' Icons only, no thumbnails
    SSF_SHOWTYPEOVERLAY = &H40000       ' Show type overlay
    SSF_SHOWSTATUSBAR = &H80000         ' Show status bar
End Enum