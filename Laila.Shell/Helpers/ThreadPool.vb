Imports System.Collections.Concurrent
Imports System.Threading
Imports Laila.Shell.Interop

Namespace Helpers
    Public Class ThreadPool
        Implements IDisposable

        Public Property TaskQueues As List(Of BlockingCollection(Of Action))

        Private _size As Integer
        Private _threadsLock As Object = New Object()
        Private _threads As List(Of Thread) = New List(Of Thread)()
        Private _isThreadFree As Boolean()
        Private _isThreadLocked As Boolean()
        Private _nextThreadId As Integer = 0
        Private _disposeTokensSource As CancellationTokenSource = New CancellationTokenSource()
        Private _disposeToken As CancellationToken = _disposeTokensSource.Token
        Private disposedValue As Boolean
        Private _rand As New Random()

        Public Sub New(size As Integer)
            Me.TaskQueues = New List(Of BlockingCollection(Of Action))
            Me.Redimension(size)
            Shell.AddToThreadPoolCache(Me)
        End Sub

        Public Sub Redimension(size As Integer)
            SyncLock _threadsLock
                If size < _size Then Throw New ArgumentException("ThreadPool.Redimension: size must be greater than the current size.")
                If size = _size Then Return

                ' initialize
                ReDim Preserve _isThreadFree(size - 1)
                ReDim Preserve _isThreadLocked(size - 1)
                For i = _size To size - 1
                    _isThreadFree(i) = True
                Next

                ' create threads
                For i = _size To size - 1
                    Me.TaskQueues.Add(New BlockingCollection(Of Action))
                    Dim thread As Thread = Nothing
                    thread = New Thread(
                        Sub(obj As Object)
                            Dim threadId As Integer = Convert.ToInt32(obj)
                            Try
                                ' Process tasks from the queue
                                Functions.OleInitialize(IntPtr.Zero)
                                For Each task In Me.TaskQueues(threadId).GetConsumingEnumerable(_disposeToken)
                                    _isThreadFree(threadId) = False
                                    task.Invoke()
                                    If Not _disposeToken.IsCancellationRequested Then _isThreadFree(threadId) = True
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

                ' remember new size
                _size = size
            End SyncLock
        End Sub

        Public ReadOnly Property Size As Integer
            Get
                Return _size
            End Get
        End Property

        Public Function GetNextFreeThreadId() As Integer
            If _isThreadFree.Count > 1 Then
                SyncLock _threadsLock
                    Do
                        _nextThreadId += 1
                        If _nextThreadId >= _isThreadFree.Length Then
                            _nextThreadId = 0
                            If _isThreadFree.All(Function(f) f = False) Then
                                Thread.Sleep(250)
                                If System.Windows.Application.Current.Dispatcher.CheckAccess Then
                                    UIHelper.OnUIThread(
                                        Sub()
                                        End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                                End If
                            End If
                            End If
                    Loop Until _isThreadFree(_nextThreadId) AndAlso Not _isThreadLocked(_nextThreadId)
                    Return _nextThreadId
                End SyncLock
            Else
                Return 0
            End If
        End Function

        Public Sub LockThread(threadId As Integer)
            _isThreadLocked(threadId) = True
            Do While Me.TaskQueues(threadId).Count > 0
                Thread.Sleep(1)
            Loop
        End Sub

        Public Sub UnlockThread(threadId As Integer)
            _isThreadLocked(threadId) = False
        End Sub

        Public Sub Add(action As Action, Optional threadId As Integer? = Nothing)
            If Not threadId.HasValue OrElse threadId.Value = -1 Then
                threadId = _rand.Next(0, _size)
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
            Dim tcs As TaskCompletionSource(Of TResult) = Nothing
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

        Public Sub Cancel()
            _disposeTokensSource.Cancel()
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects)
                    Me.Cancel()

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