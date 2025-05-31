Namespace Controls.Parts
    Public Interface ISuggestionProvider
        Inherits ISuggestionProviderSyncOrAsync
        Function GetSuggestions(ByVal filter As String) As IEnumerable
    End Interface
End Namespace
