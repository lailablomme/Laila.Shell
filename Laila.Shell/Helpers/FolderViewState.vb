﻿Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml.Serialization
Imports LiteDB

Namespace Helpers
    Public Class FolderViewState
        Inherits Behaviors.GridViewExtBehavior.GridViewStateData

        Public Property ViewName As String
        Public Property View As String

        Public Sub New()
            Me.View = "Details"
            Me.SortPropertyName = "ItemNameDisplaySortValue"
        End Sub

        Public Shared Function FromViewName(viewName As String) As FolderViewState
            Dim result As FolderViewState

            Dim dbPath As String = GetStateDBPath()
            Dim viewId As String = getMD5Hash(viewName)

            If Not File.Exists(Path.Combine(dbPath, viewId)) Then
                result = New FolderViewState()
            Else
                Using stream As FileStream = New FileStream(Path.Combine(dbPath, viewId), FileMode.Open, FileAccess.Read)
                    Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
                    result = s.Deserialize(stream)
                End Using
            End If

            result.ViewName = viewName
            Return result
        End Function

        Public Sub Persist()
            Dim dbPath As String = GetStateDBPath()
            Dim viewId As String = getMD5Hash(ViewName)

            Using stream As FileStream = New FileStream(Path.Combine(dbPath, viewId), FileMode.Create, FileAccess.Write)
                Dim s As XmlSerializer = New XmlSerializer(GetType(FolderViewState))
                s.Serialize(stream, Me)
            End Using
        End Sub

        Private Shared Function getMD5Hash(input As String) As String
            Using md5 As MD5 = MD5.Create()
                ' Convert the input string to a byte array and compute the hash
                Dim inputBytes As Byte() = Encoding.UTF8.GetBytes(input)
                Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)

                ' Convert the byte array to a hexadecimal string
                Dim sb As New StringBuilder()
                For Each b As Byte In hashBytes
                    sb.Append(b.ToString("x2"))
                Next

                Return sb.ToString()
            End Using
        End Function

        Public Shared Function GetStateDBPath() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell", "FolderViewState")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return path
        End Function
    End Class
End Namespace