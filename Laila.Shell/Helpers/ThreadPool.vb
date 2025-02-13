﻿Imports System.Collections.Concurrent
Imports System.Threading

Namespace Helpers
    Public Class ThreadPool
        Implements IDisposable

        Public ReadOnly Property TaskQueues As List(Of BlockingCollection(Of Action))

        Private _threadsLock As Object = New Object()
        Private _threads As List(Of Thread) = New List(Of Thread)()
        Private _isThreadFree As Boolean()
        Private _nextThreadIdLock As Object = New Object()
        Private _nextThreadId As Integer = 0
        Private _disposeTokensSource As CancellationTokenSource = New CancellationTokenSource()
        Private _disposeToken As CancellationToken = _disposeTokensSource.Token
        Private disposedValue As Boolean

        Public Sub New(size As Integer)
            SyncLock _threadsLock
                ' initialize
                Me.TaskQueues = New List(Of BlockingCollection(Of Action))
                ReDim _isThreadFree(size - 1)
                For i = 0 To size - 1
                    _isThreadFree(i) = True
                Next

                ' create threads
                For i = 0 To size - 1
                    Me.TaskQueues.Add(New BlockingCollection(Of Action))
                    Dim thread As Thread = New Thread(
                        Sub(obj As Object)
                            Dim threadId As Integer = Convert.ToInt32(obj)
                            Try
                                ' Process tasks from the queue
                                Functions.OleInitialize(IntPtr.Zero)
                                For Each task In Me.TaskQueues(threadId).GetConsumingEnumerable(_disposeToken)
                                    _isThreadFree(threadId) = False
                                    task.Invoke()
                                    _isThreadFree(threadId) = True
                                Next
                            Catch ex As OperationCanceledException
                                Debug.WriteLine("ThreadPool TaskQueue thread was canceled.")
                            End Try
                            Functions.OleUninitialize()
                            SyncLock _threadsLock
                                _threads.Remove(thread)
                            End SyncLock
                        End Sub)
                    thread.IsBackground = False
                    thread.SetApartmentState(ApartmentState.STA)
                    thread.Start(i)
                    _threads.Add(thread)
                Next
            End SyncLock

            Shell.AddToThreadPoolCache(Me)
        End Sub

        Public Function GetNextFreeThreadId() As Integer
            If _isThreadFree.Count > 1 Then
                Dim isFirst As Boolean = True
                SyncLock _nextThreadIdLock
                    Do
                        _nextThreadId += 1
                        If _nextThreadId >= _isThreadFree.Length Then
                            _nextThreadId = 0
                            If Not isFirst Then Thread.Sleep(10)
                            isFirst = False
                        End If
                    Loop Until _isThreadFree(_nextThreadId)
                    Return _nextThreadId
                End SyncLock
            Else
                Return 0
            End If
        End Function

        Public Sub Add(action As Action, Optional threadId As Integer? = Nothing)
            If Not threadId.HasValue Then
                threadId = GetNextFreeThreadId()
            End If
            Me.TaskQueues(threadId).Add(action)
        End Sub

        Public Sub Run(action As Action, Optional maxRetries As Integer = 1, Optional threadId As Integer? = Nothing)
            Me.Run(
                Function() As Boolean
                    action()
                    Return True
                End Function, maxRetries, threadId)
        End Sub

        Public Function Run(Of TResult)(func As Func(Of TResult), Optional maxRetries As Integer = 1, Optional threadId As Integer? = Nothing) As TResult
            If Not threadId.HasValue Then
                threadId = GetNextFreeThreadId()
            End If
            Dim tcs As TaskCompletionSource(Of TResult)
            Dim numTries = 1
            While (tcs Is Nothing OrElse Not (tcs.Task.IsCompleted OrElse tcs.Task.IsCanceled)) _
                AndAlso numTries <= 3 AndAlso Not _disposeToken.IsCancellationRequested
                Try
                    tcs = New TaskCompletionSource(Of TResult)()
                    Me.TaskQueues(threadId).Add(
                        Sub()
                            Try
                                tcs.SetResult(func())
                            Catch ex As Exception
                                Debug.WriteLine("ThreadPool.Run TaskQueue Exception: " & ex.Message)
                                tcs.SetException(ex)
                            End Try
                        End Sub)
                    tcs.Task.Wait(_disposeToken)
                    If numTries > 1 AndAlso tcs.Task.IsCompleted Then
                        Debug.WriteLine("ThreadPool.Run succeeded after " & numTries & " tries")
                    ElseIf tcs.Task.IsFaulted Then
                        numTries += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine("ThreadPool.Run While Exception: " & ex.Message)
                    numTries += 1
                End Try
            End While
            If Not tcs Is Nothing AndAlso tcs.Task.IsCompleted _
                AndAlso Not _disposeToken.IsCancellationRequested Then
                Return tcs.Task.Result
            Else
                Return Nothing
            End If
        End Function

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects)
                    _disposeTokensSource.Cancel()

                    ' wait for threads to finish so we don't have memory issues on shutdown
                    Dim count As Integer = -1
                    Do
                        SyncLock _threadsLock
                            count = _threads.Count
                        End SyncLock
                        Thread.Sleep(1)
                    Loop While count <> 0

                    Shell.RemoveFromThreadPoolCache(Me)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace