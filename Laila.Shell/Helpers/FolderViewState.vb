Imports Laila.Shell.Behaviors.GridViewExtBehavior
Imports LiteDB
Imports System.IO
Imports System.Xml.Serialization

Namespace Helpers
    Public Class FolderViewState
        Inherits Behaviors.GridViewExtBehavior.GridViewStateData

        Public Property View As String

        Public Sub New()
            Me.View = "Details"
            Me.SortPropertyName = "ItemNameDisplaySortValue"
        End Sub

        Public Shared Function FromViewName(viewName As String) As FolderViewState
            Dim dbFileName As String = GetStateDBFileName()
            Dim viewId As String = Convert.ToBase64String(Text.Encoding.UTF8.GetBytes(viewName))

            If Not File.Exists(dbFileName) Then
                Return New FolderViewState()
            End If

            Using mem As MemoryStream = New MemoryStream()
                Using db = New LiteDatabase(dbFileName)
                    Dim info As LiteFileInfo(Of String) = db.FileStorage.FindById(viewId)
                    If Not info Is Nothing Then
                        info.CopyTo(mem)
                        mem.Seek(0, SeekOrigin.Begin)
                        Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
                        Return s.Deserialize(mem)
                    Else
                        Return New FolderViewState()
                    End If
                End Using
            End Using
        End Function

        Public Sub Persist(viewName As String)
            Dim dbFileName As String = GetStateDBFileName()
            Dim viewId As String = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(viewName))

            If Not Directory.Exists(Path.GetDirectoryName(dbFileName)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(dbFileName))
            End If

            Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
            Using mem As MemoryStream = New MemoryStream()
                s.Serialize(mem, Me)
                mem.Seek(0, SeekOrigin.Begin)

                Using db = New LiteDatabase(dbFileName)
                    db.FileStorage.Upload(viewId, "state.xml", mem)
                End Using
            End Using
        End Sub

        Public Shared Function GetStateDBFileName() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return IO.Path.Combine(path, "FolderViewState.db")
        End Function
    End Class
End Namespace