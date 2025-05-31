Namespace Controls.Parts
    Public Interface ISuggestionProviderAsync
        Inherits ISuggestionProviderSyncOrAsync

        Function GetSuggestions(ByVal filter As String) As Task(Of IEnumerable)
    End Interface
End Namespace