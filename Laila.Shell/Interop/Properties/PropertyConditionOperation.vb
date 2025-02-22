Namespace Interop.Properties
    Public Enum PropertyConditionOperation
        '/// <summary>
        '  /// The implicit comparison between the value of the property And the value of the constant.
        '  /// </summary>
        Implicit

        '/// <summary>
        '/// The value of the property And the value of the constant must be equal.
        '/// </summary>
        Equal

        '/// <summary>
        '/// The value of the property And the value of the constant must Not be equal.
        '/// </summary>
        NotEqual

        '/// <summary>
        '/// The value of the property must be less than the value of the constant.
        '/// </summary>
        LessThan

        '/// <summary>
        '/// The value of the property must be greater than the value of the constant.
        '/// </summary>
        GreaterThan

        '/// <summary>
        '/// The value of the property must be less than Or equal to the value of the constant.
        '/// </summary>
        LessThanOrEqual

        '/// <summary>
        '/// The value of the property must be greater than Or equal to the value of the constant.
        '/// </summary>
        GreaterThanOrEqual

        '/// <summary>
        '/// The value of the property must begin with the value of the constant.
        '/// </summary>
        ValueStartsWith

        '/// <summary>
        '/// The value of the property must end with the value of the constant.
        '/// </summary>
        ValueEndsWith

        '/// <summary>
        '/// The value of the property must contain the value of the constant.
        '/// </summary>
        ValueContains

        '/// <summary>
        '/// The value of the property must Not contain the value of the constant.
        '/// </summary>
        ValueNotContains

        '/// <summary>
        '/// The value of the property must match the value of the constant, where '?' matches any single character and '*' matches any sequence of characters.
        '/// </summary>
        DOSWildCards

        '/// <summary>
        '/// The value of the property must contain a word that Is the value of the constant.
        '/// </summary>
        WordEqual

        '/// <summary>
        '/// The value of the property must contain a word that begins with the value of the constant.
        '/// </summary>
        WordStartsWith

        '/// <summary>
        '/// The application Is free to interpret this in any suitable way.
        '/// </summary>
        ApplicationSpecific
    End Enum
End Namespace

