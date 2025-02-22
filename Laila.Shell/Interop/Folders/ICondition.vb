Imports System.Runtime.InteropServices

Namespace Interop.Folders
    <ComImport>
    <Guid("0fc988d4-c935-4b97-a973-46282ea175c8")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ICondition
        <PreserveSig>
        Function GetConditionType(ByRef pNodeType As ConditionType) As Integer

        <PreserveSig>
        Function GetSubConditions(<MarshalAs(UnmanagedType.Interface)> ByRef ppEnumUnknown As IEnumUnknown) As Integer

        <PreserveSig>
        Function GetComparisonInfo(<MarshalAs(UnmanagedType.LPWStr)> ByRef ppszPropertyName As String, <MarshalAs(UnmanagedType.Interface)> ByRef pOperation As ConditionOperation, <MarshalAs(UnmanagedType.Struct)> ByRef ppropvar As Object) As Integer

        <PreserveSig>
        Function GetValueType(<MarshalAs(UnmanagedType.Struct)> ByRef pValueType As Object) As Integer

        <PreserveSig>
        Function GetValueNormalization(ByRef pNormalization As ConditionNormalization) As Integer

        <PreserveSig>
        Function GetInputTerms(<MarshalAs(UnmanagedType.Interface)> ByRef ppTerms As IEnumUnknown) As Integer

        <PreserveSig>
        Function Clone(<MarshalAs(UnmanagedType.Interface)> ByRef ppCondition As ICondition) As Integer
    End Interface

    Public Enum ConditionType
        CTAndCondition = 0
        CTOrCondition = 1
        CTNotCondition = 2
        CTLeafCondition = 3
    End Enum

    Public Enum ConditionOperation
        COPEqual = 0
        COPNotEqual = 1
        COPLessThan = 2
        COPGreaterThan = 3
        COPLessEqual = 4
        COPGreaterEqual = 5
        COPValueStartsWith = 6
        COPValueEndsWith = 7
        COPValueContains = 8
        COPValueNotContains = 9
        COPDOSWildcards = 10
        COPWordEqual = 11
        COPWordStartsWith = 12
    End Enum

    Public Enum ConditionNormalization
        CNFullNormalization = 0
        CNMinimalNormalization = 1
    End Enum
End Namespace
