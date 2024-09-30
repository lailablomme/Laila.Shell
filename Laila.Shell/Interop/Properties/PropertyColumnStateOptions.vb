<Flags>
Public Enum PropertyColumnStateOptions
    '/// <summary>
    '   /// Default value
    '   /// </summary>
    None = &H0

    '/// <summary>
    '/// The value Is displayed as a string.
    '/// </summary>
    StringType = &H1

    '/// <summary>
    '/// The value Is displayed as an integer.
    '/// </summary>
    IntegerType = &H2

    '/// <summary>
    '/// The value Is displayed as a date/time.
    '/// </summary>
    DateType = &H3

    '/// <summary>
    '/// A mask for display type values StringType, IntegerType, And DateType.
    '/// </summary>
    TypeMask = &HF

    '/// <summary>
    '/// The column should be on by default in Details view.
    '/// </summary>
    OnByDefault = &H10

    '/// <summary>
    '/// Will be slow to compute. Perform on a background thread.
    '/// </summary>
    Slow = &H20

    '/// <summary>
    '/// Provided by a handler, Not the folder.
    '/// </summary>
    Extended = &H40

    '/// <summary>
    '/// Not displayed in the context menu, but Is listed in the More... dialog.
    '/// </summary>
    SecondaryUI = &H80

    '/// <summary>
    '/// Not displayed in the user interface (UI).
    '/// </summary>
    Hidden = &H100

    '/// <summary>
    '/// VarCmp produces same result as IShellFolder:CompareIDs.
    '/// </summary>
    PreferVariantCompare = &H200

    '/// <summary>
    '/// PSFormatForDisplay produces same result as IShellFolder:CompareIDs.
    '/// </summary>
    PreferFormatForDisplay = &H400

    '/// <summary>
    '/// Do Not sort folders separately.
    '/// </summary>
    NoSortByFolders = &H800

    '/// <summary>
    '/// Only displayed in the UI.
    '/// </summary>
    ViewOnly = &H10000

    '/// <summary>
    '/// Marks columns with values that should be read in a batch.
    '/// </summary>
    BatchRead = &H20000

    '/// <summary>
    '/// Grouping Is disabled for this column.
    '/// </summary>
    NoGroupBy = &H40000

    '/// <summary>
    '/// Can't resize the column.
    '/// </summary>
    FixedWidth = &H1000

    '/// <summary>
    '/// The width Is the same in all dots per inch (dpi)s.
    '/// </summary>
    NoDpiScale = &H2000

    '/// <summary>
    '/// Fixed width And height ratio.
    '/// </summary>
    FixedRatio = &H4000

    '/// <summary>
    '/// Filters out New display flags.
    '/// </summary>
    DisplayMask = &HF000
End Enum
