Namespace Interop.Properties
    <Flags>
    Public Enum PropertyViewOptions
        '/// <summary>
        '   /// The property Is shown by default.
        '   /// </summary>
        None = &H0

        '/// <summary>
        '/// The property Is centered.
        '/// </summary>
        CenterAlign = &H1

        '/// <summary>
        '/// The property Is right aligned.
        '/// </summary>
        RightAlign = &H2

        '/// <summary>
        '/// The property Is shown as the beginning of the next collection of properties in the view.
        '/// </summary>
        BeginNewGroup = &H4

        '/// <summary>
        '/// The remainder of the view area Is filled with the content of this property.
        '/// </summary>
        FillArea = &H8

        '/// <summary>
        '/// The property Is reverse sorted if it Is a property in a list of sorted properties.
        '/// </summary>
        SortDescending = &H10

        '/// <summary>
        '/// The property Is only shown if it Is present.
        '/// </summary>
        ShowOnlyIfPresent = &H20

        '/// <summary>
        '/// The property Is shown by default in a view (where applicable).
        '/// </summary>
        ShowByDefault = &H40

        '/// <summary>
        '/// The property Is shown by default in primary column selection user interface (UI).
        '/// </summary>
        ShowInPrimaryList = &H80

        '/// <summary>
        '/// The property Is shown by default in secondary column selection UI.
        '/// </summary>
        ShowInSecondaryList = &H100

        '/// <summary>
        '/// The label Is hidden if the view Is normally inclined to show the label.
        '/// </summary>
        HideLabel = &H200

        '/// <summary>
        '/// The property Is Not displayed as a column in the UI.
        '/// </summary>
        Hidden = &H800

        '/// <summary>
        '/// The property Is wrapped to the next row.
        '/// </summary>
        CanWrap = &H1000

        '/// <summary>
        '/// A mask used to retrieve all flags.
        '/// </summary>
        MaskAll = &H3FF
    End Enum
End Namespace
