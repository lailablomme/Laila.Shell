Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.COM
Imports Laila.Shell.Interop.System

Namespace Interop.Folders
    <ComImport>
    <Guid("D6EBC66B-8921-4193-AFDD-A1789FB7FF57")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IQuerySolution
        ' Creates a condition node that is a logical conjunction (AND) or disjunction (OR)
        Sub MakeAndOr(
        <[In]> ByVal rgpConditions() As ICondition,
        <[In]> ByVal cConditions As Integer,
        <[In]> ByVal fAnd As Boolean,
        <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppCondition As ICondition)

        ' Creates a leaf condition node that represents a comparison of property value and constant value
        Sub MakeLeaf(
        <[In]> ByVal pszPropertyName As String,
        <[In]> ByVal nOp As Integer, ' Comparison operation type (e.g., equal, less than, etc.)
        <[In]> ByVal pszValue As String,
        <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppCondition As ICondition)

        ' Creates a condition node that is a logical negation (NOT) of another condition (subnode)
        Sub MakeNot(
        <[In], MarshalAs(UnmanagedType.Interface)> ByVal pCondition As ICondition,
        <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppCondition As ICondition)

        ' Resolves conditions to simpler or more absolute forms
        Function Resolve(<MarshalAs(UnmanagedType.Interface)> ByVal pCondition As ICondition, ByVal dwResolveFlags As Integer, ByRef stResolveTime As SYSTEMTIME, <MarshalAs(UnmanagedType.Interface)> ByRef ppResolvedCondition As ICondition) As Integer

        ' Gets the root of the tree of the query.
        <PreserveSig>
        Function GetQuery(<MarshalAs(UnmanagedType.Interface)> ByRef ppQueryNode As ICondition, ByRef pMainType As QueryNodeType) As Integer

        ' Retrieves lexical data with correct parameters and types
        <PreserveSig>
        Function GetLexicalData(<Out> ByRef ppszInputString As IntPtr,
                            <Out> ByRef ppTokens As IntPtr,
                            <Out> ByRef plcid As Integer,
                            <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppWordBreaker As IUnknown) As Integer

        ' Retrieves parsing errors
        Sub GetErrors(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppParseError As IEnumUnknown)
    End Interface

    Public Enum QueryNodeType
        AndCondition = 0
        OrCondition = 1
        NotCondition = 2
        LeafCondition = 3
    End Enum
End Namespace




