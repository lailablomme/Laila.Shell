Namespace Interop.Properties
    Public Enum PropEnumType
        '/// <summary>
        '  /// Use DisplayText And either RangeMinValue Or RangeSetValue.
        '  /// </summary>
        DiscreteValue = 0

        '/// <summary>
        '/// Use DisplayText And either RangeMinValue Or RangeSetValue
        '/// </summary>
        RangedValue = 1

        '/// <summary>
        '/// Use DisplayText
        '/// </summary>
        DefaultValue = 2

        '/// <summary>
        '/// Use Value Or RangeMinValue
        '/// </summary>
        EndRange = 3
    End Enum
End Namespace
