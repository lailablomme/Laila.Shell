<Flags>
Public Enum PropertyTypeOptions
    '/// <summary>
    '  /// The property uses the default values for all attributes.
    '  /// </summary>
    None = &H0

    '/// <summary>
    '/// The property can have multiple values.   
    '/// </summary>
    '/// <remarks>
    '/// These values are stored as a VT_VECTOR in the PROPVARIANT structure.
    '/// This value Is set by the multipleValues attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    MultipleValues = &H1

    '/// <summary>
    '/// This property cannot be written to. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the isInnate attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IsInnate = &H2

    '/// <summary>
    '/// The property Is a group heading. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the isGroup attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IsGroup = &H4

    '/// <summary>
    '/// The user can group by this property. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the canGroupBy attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    CanGroupBy = &H8

    '/// <summary>
    '/// The user can stack by this property. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the canStackBy attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    CanStackBy = &H10

    '/// <summary>
    '/// This property contains a hierarchy. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the isTreeProperty attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IsTreeProperty = &H20

    '/// <summary>
    '/// Include this property in any full text query that Is performed. 
    '/// </summary>
    '/// <remarks>
    '/// This value Is set by the includeInFullTextQuery attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IncludeInFullTextQuery = &H40

    '/// <summary>
    '/// This property Is meant to be viewed by the user.  
    '/// </summary>
    '/// <remarks>
    '/// This influences whether the property shows up in the "Choose Columns" dialog, for example.
    '/// This value Is set by the isViewable attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IsViewable = &H80

    '/// <summary>
    '/// This property Is included in the list of properties that can be queried.   
    '/// </summary>
    '/// <remarks>
    '/// A queryable property must also be viewable.
    '/// This influences whether the property shows up in the query builder UI.
    '/// This value Is set by the isQueryable attribute of the typeInfo element in the property's .propdesc file.
    '/// </remarks>
    IsQueryable = &H100

    '/// <summary>
    '/// Used with an innate property (that Is, a value calculated from other property values) to indicate that it can be deleted.  
    '/// </summary>
    '/// <remarks>
    '/// Windows Vista with Service Pack 1 (SP1) And later.
    '/// This value Is used by the Remove Properties user interface (UI) to determine whether to display a check box next to an property that allows that property to be selected for removal.
    '/// Note that a property that Is Not innate can always be purged regardless of the presence Or absence of this flag.
    '/// </remarks>
    CanBePurged = &H200

    '/// <summary>
    '/// This property Is owned by the system.
    '/// </summary>
    IsSystemProperty = &H80000000

    '/// <summary>
    '/// A mask used to retrieve all flags.
    '/// </summary>
    MaskAll = &H800001FF
End Enum
