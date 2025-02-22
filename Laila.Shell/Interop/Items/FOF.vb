Namespace Interop.Items
    <Flags>
    Public Enum FOF As UInteger
        ' Preserve undo information (move to Recycle Bin instead of permanent delete)
        FOF_ALLOWUNDO = &H40

        ' Do not display a progress dialog box
        FOF_SILENT = &H4

        ' Respond with Yes to All for any dialog box that is displayed
        FOF_NOCONFIRMATION = &H10

        ' If an error occurs, skip the problem file or folder and continue with the operation
        FOF_NOERRORUI = &H400

        ' Do not copy connected files as a group; only copy the specified files
        FOF_NOCONFIRMMKDIR = &H200

        ' Perform the operation on files only if a wildcard file name (*.*) is specified
        FOF_FILESONLY = &H80

        ' Perform the operation on a folder's contents, not the folder itself
        FOF_NORECURSION = &H1000

        ' Do not move connected files as a group; only move the specified files
        FOF_RENAMEONCOLLISION = &H8

        ' Display a progress dialog box but do not show the names of the files
        FOF_SIMPLEPROGRESS = &H100

        ' Do not allow the progress dialog box to be canceled
        FOF_NO_UI = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOERRORUI

        ' Treat reparse points (e.g., symlinks) as files instead of folders
        FOF_NORECURSEREPARSE = &H8000

        ' Only work with items in the selection if they belong to the same parent folder
        FOF_NO_CONNECTED_ELEMENTS = &H2000

        ' Do not send any of the items to the Recycle Bin
        FOF_WANTNUKEWARNING = &H4000

        ' Enable operations on hard links or files with streams
        FOFX_PRESERVEFILEEXTENSIONS = &H80000

        ' Notify the calling application of any user interface updates
        FOFX_SHOWELEVATIONPROMPT = &H400000

        ' Allow the operation to include unregistered file types
        FOFX_PRESERVEMETADATA = &H80000

        ' Enable extended error reporting
        FOFX_ERRORUIFILTER = &H1000000

        ' Suppress operations on items that are duplicates of other items
        FOFX_DUPLICATEFILES = &H2000000

        ' Do not use any undo operations
        FOFX_NOUNDO = &H4000000

        ' Do not remove the copy source after the operation
        FOFX_NOERRORUI = FOF_NOERRORUI

        ' Disable the final notification for success or failure
        FOFX_DISABLE_UNDO = &H8000

        ' Request confirmation for any irreversible operation
        FOFX_REQUIRES_USER_CONSENT = &H10000

        ' Always prompt for confirmation of the operation
        FOFX_WANTNUKEWARNING = FOF_WANTNUKEWARNING
    End Enum
End Namespace
