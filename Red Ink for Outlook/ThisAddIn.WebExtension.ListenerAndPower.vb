' =============================================================================
' File: ThisAddIn.WebExtension.ListenerAndPower.vb
' Part of: Red Ink for Outlook
' Purpose: Hosts a local HttpListener for the web extension, manages suspend/resume
'          transitions, listener lifecycle, watchdog monitoring, job cancellation,
'          and recovery of stalled operations.
'
' Copyright: David Rosenthal, david.rosenthal@vischer.com
' License: May only be used with an appropriate license (see redink.ai)
'
' Architecture:
' - HttpListener loop: restart-safe via a generation counter (listenerGeneration).
' - Power handling: hidden message-only window (PowerNotificationWindow) receives
'   WM_POWERBROADCAST suspend/resume messages and triggers async handlers.
' - Synchronization: SemaphoreSlim (suspendResumeGate) serializes suspend/resume sequences.
' - Cancellation: Active jobs and in-flight LLM operations are cancelled before suspend.
' - Watchdog: Timer periodically evaluates listener and request/job activity, restarting
'   listener if it is dead or faulted.
' - Recovery: Performs stuck job cancellation, listener reset (if inactive), and COM/OLE
'   message filter re-registration.
' - Resume cooldown: Prevents premature watchdog actions immediately after system resume.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Net
Imports System.Threading.Tasks

Partial Public Class ThisAddIn

    Private httpListener As HttpListener
    Private isShuttingDown As Boolean = False
    Private listenerTask As System.Threading.Tasks.Task

    Private llmOperationCts As System.Threading.CancellationTokenSource

    Private activeRequests As Integer = 0
    Private ModelTimeout As Integer = 300

    ' Threading & listener state (class level)
    Private llmSyncContext As System.Threading.SynchronizationContext
    Private llmThread As System.Threading.Thread
    Private cts As System.Threading.CancellationTokenSource

    Private wasListenerRunningBeforeSleep As System.Boolean = False
    Private wasLlmThreadAliveBeforeSleep As System.Boolean = False
    Private restartingAfterResume As System.Int32 = 0  ' 0/1 via Interlocked

    ' Guard to mute watchdog and concurrent restarts during power transitions
    Private powerChanging As System.Int32 = 0

    ' Generation protection (pre/post sleep)
    Private listenerGeneration As System.Int64 = 0

    ' Progress watchdog
    Private lastListenerProgressUtc As System.DateTime = System.DateTime.UtcNow
    Private watchdog As System.Threading.Timer

    ' Power notifications via hidden window
    Private powerWindow As PowerNotificationWindow

    ' Resume cooldown boundary
    Private resumeCooldownUntilUtc As System.DateTime = System.DateTime.MinValue
    Private inMainMenu As System.Int32 = 0

    ''' <summary>
    ''' Returns true if resume cooldown interval has not elapsed yet.
    ''' </summary>
    Private Function IsInResumeCooldown() As System.Boolean
        Return System.DateTime.UtcNow < resumeCooldownUntilUtc
    End Function

    ''' <summary>
    ''' Message-only window receiving WM_POWERBROADCAST for suspend/resume events.
    ''' </summary>
    Private NotInheritable Class PowerNotificationWindow
        Inherits System.Windows.Forms.NativeWindow
        Implements System.IDisposable

        Private Const WM_POWERBROADCAST As System.Int32 = &H218
        Private Const PBT_APMSUSPEND As System.Int32 = &H4
        Private Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
        Private Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

        Private ReadOnly owner As ThisAddIn

        Public Sub New(ByVal owner As ThisAddIn)
            Me.owner = owner
            Dim cp As New System.Windows.Forms.CreateParams()
            cp.Caption = "InkyPowerWnd"
            cp.X = 0 : cp.Y = 0 : cp.Height = 0 : cp.Width = 0
            cp.Style = 0 : cp.ExStyle = 0
            ' Important: message-only window (HWND_MESSAGE)
            cp.Parent = New System.IntPtr(-3) ' HWND_MESSAGE
            Me.CreateHandle(cp)
        End Sub

        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            Const WM_POWERBROADCAST As System.Int32 = &H218
            Const PBT_APMQUERYSUSPEND As System.Int32 = &H0
            Const PBT_APMSUSPEND As System.Int32 = &H4
            Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
            Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

            If m.Msg = WM_POWERBROADCAST Then
                Dim wp As System.Int32 = m.WParam.ToInt32()
                Select Case wp
                    Case PBT_APMQUERYSUSPEND
                        ' Immediately acknowledge and do nothing.
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMSUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerSuspendAsync() : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMEAUTOMATIC
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerResumeAsync(userPresent:=False) : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMESUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerResumeAsync(userPresent:=True) : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return
                End Select
            End If

            MyBase.WndProc(m)
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If Me.Handle <> System.IntPtr.Zero Then
                Me.DestroyHandle()
            End If
        End Sub
    End Class

    ' Ensure only one suspend/resume sequence runs at a time
    Private suspendResumeGate As New System.Threading.SemaphoreSlim(1, 1)

    ''' <summary>
    ''' Asynchronously handles power suspend: cancels jobs, records state, stops listener, signals LLM thread exit.
    ''' </summary>
    Friend Sub HandlePowerSuspendAsync()
        System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task
                If Not Await TryEnterGateAsync().ConfigureAwait(False) Then Return
                System.Threading.Interlocked.Exchange(powerChanging, 1)
                Try
                    ' Cancel all jobs FIRST and wait for them briefly
                    Dim cancelTasks As New List(Of System.Threading.Tasks.Task)()
                    For Each kv In jobMap
                        Try
                            kv.Value.Cts.Cancel()
                            ' Add a wait for the task to complete
                            If kv.Value.Tcs IsNot Nothing AndAlso Not kv.Value.Tcs.Task.IsCompleted Then
                                cancelTasks.Add(kv.Value.Tcs.Task.ContinueWith(Sub(t)
                                                                                   ' Swallow exceptions
                                                                               End Sub, TaskContinuationOptions.ExecuteSynchronously))
                            End If
                        Catch
                        End Try
                    Next

                    ' Wait briefly for jobs to terminate
                    If cancelTasks.Count > 0 Then
                        Try
                            Await System.Threading.Tasks.Task.WhenAny(
                                System.Threading.Tasks.Task.WhenAll(cancelTasks),
                                System.Threading.Tasks.Task.Delay(2000)
                            ).ConfigureAwait(False)
                        Catch
                            ' Ignore exceptions from cancelled tasks
                        End Try
                    End If

                    ' Clear the job map after cancellation
                    jobMap.Clear()
                    System.Threading.Interlocked.Exchange(activeJobs, 0)

                    ' Mute watchdog during suspend
                    Try : StopListenerWatchdog() : Catch : End Try

                    ' Remember state
                    Try
                        wasListenerRunningBeforeSleep =
                            (httpListener IsNot Nothing AndAlso httpListener.IsListening)
                    Catch
                        wasListenerRunningBeforeSleep = False
                    End Try
                    Try
                        wasLlmThreadAliveBeforeSleep =
                            (llmThread IsNot Nothing AndAlso llmThread.IsAlive)
                    Catch
                        wasLlmThreadAliveBeforeSleep = False
                    End Try

                    ' Proactively cancel any in-flight LLM op (prevents wake-up on dead STA)
                    Try
                        If llmOperationCts IsNot Nothing Then
                            If Not llmOperationCts.IsCancellationRequested Then llmOperationCts.Cancel()
                            llmOperationCts.Dispose()
                        End If
                    Catch
                    Finally
                        llmOperationCts = Nothing
                    End Try

                    ' Force any stale listener loop to exit quickly
                    System.Threading.Interlocked.Increment(listenerGeneration)

                    ' Stop listener without stopping UI STA thread
                    Try
                        Dim t As System.Threading.Tasks.Task = ShutdownHttpListener(stopUiThread:=False)
                        Await System.Threading.Tasks.Task.WhenAny(
                            t,
                            System.Threading.Tasks.Task.Delay(1000)
                        ).ConfigureAwait(False)
                    Catch
                    End Try

                    ' LLM STA: request exit without waiting (no Join while suspending)
                    Try
                        If wasLlmThreadAliveBeforeSleep Then
                            StopLlmUiThreadNonBlocking()
                        End If
                    Catch
                    End Try

                Finally
                    suspendResumeGate.Release()
                End Try
            End Function)
    End Sub

    ''' <summary>
    ''' Asynchronously handles power resume: conditional listener restart and cooldown initialization.
    ''' </summary>
    Friend Sub HandlePowerResumeAsync(userPresent As Boolean)
        System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task
                ' Do not wait for gate - if acquisition fails immediately, skip
                If Not suspendResumeGate.Wait(0) Then Return

                Try
                    ' Minimal delay before restart
                    Await System.Threading.Tasks.Task.Delay(500).ConfigureAwait(False)

                    isShuttingDown = False

                    ' Clean up old listener reference
                    httpListener = Nothing

                    ' Restart only if user was present pre-suspend
                    If userPresent AndAlso wasListenerRunningBeforeSleep Then
                        Try
                            StartupHttpListener()
                        Catch
                        End Try
                    End If

                    ' Set resume cooldown
                    resumeCooldownUntilUtc = System.DateTime.UtcNow.AddSeconds(30)

                    ' UI not updated here

                Finally
                    System.Threading.Interlocked.Exchange(powerChanging, 0)
                    suspendResumeGate.Release()
                End Try
            End Function)
    End Sub

    ''' <summary>
    ''' Signals UI STA thread to exit without blocking caller. Clears related references.
    ''' </summary>
    Private Sub StopLlmUiThreadNonBlocking()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(Sub() System.Windows.Forms.Application.ExitThread(), Nothing)
            End If
        Catch
        End Try
        ' No Join here.
        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub

    ''' <summary>
    ''' Attempts to enter the suspend/resume gate with a short timeout.
    ''' </summary>
    Private Async Function TryEnterGateAsync() As System.Threading.Tasks.Task(Of System.Boolean)
        Try
            Return Await suspendResumeGate.WaitAsync(100).ConfigureAwait(False)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Initializes listener generation, cancellation token source, logs, and starts listener loop.
    ''' </summary>
    Private Sub StartupHttpListener()
        ' Make sure the loop can run again
        isShuttingDown = False

        Dim gen As System.Int64 = System.Threading.Interlocked.Increment(listenerGeneration)
        cts = New System.Threading.CancellationTokenSource()

        ' Log generation and UTC timestamp
        System.Diagnostics.Debug.WriteLine(
            "HttpListener START gen=" &
            gen.ToString(System.Globalization.CultureInfo.InvariantCulture) &
            " at " &
            System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))

        lastListenerProgressUtc = System.DateTime.UtcNow
        listenerTask = StartHttpListener(cts.Token, gen)
    End Sub

    ''' <summary>
    ''' Shuts down listener loop, aborts outstanding contexts, disposes resources and optionally stops UI STA.
    ''' </summary>
    Private Async Function ShutdownHttpListener(
        Optional ByVal stopUiThread As System.Boolean = True
    ) As System.Threading.Tasks.Task
        isShuttingDown = True

        ' Cancel current loop
        Try
            If cts IsNot Nothing Then cts.Cancel()
        Catch
        End Try

        ' Force-break any pending GetContextAsync and clean up thoroughly
        Try
            If httpListener IsNot Nothing Then
                Try
                    If httpListener.IsListening Then httpListener.Stop()
                Catch
                End Try
                Try
                    httpListener.Abort() ' harsher than Close; reliably breaks GetContextAsync
                Catch
                End Try
                Try
                    If httpListener.Prefixes IsNot Nothing Then httpListener.Prefixes.Clear()
                Catch
                End Try
                Try
                    httpListener.Close()
                Catch
                End Try
            End If
        Catch
        Finally
            httpListener = Nothing
        End Try

        ' Await the running listener task to completion
        Try
            If listenerTask IsNot Nothing Then
                Await listenerTask.ConfigureAwait(False)
            End If
        Catch
        Finally
            listenerTask = Nothing
        End Try

        ' Dispose CTS after we've awaited its dependents
        Try
            If cts IsNot Nothing Then cts.Dispose()
        Catch
        Finally
            cts = Nothing
        End Try

        System.Diagnostics.Debug.WriteLine(
            "HttpListener STOP at " &
            System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))

        ' Stop UI STA only if requested
        If stopUiThread Then
            StopLlmUiThread()
        End If
    End Function

    ''' <summary>
    ''' Listener loop: accepts requests, dispatches per-request tasks, logs metrics, and self-recovers after repeated failures.
    ''' </summary>
    Private Async Function StartHttpListener(
        ByVal token As System.Threading.CancellationToken,
        ByVal gen As System.Int64
    ) As System.Threading.Tasks.Task

        Const prefix As System.String = "http://127.0.0.1:12333/"
        Dim consecutiveFailures As System.Int32 = 0
        Dim lastMetrics As System.DateTime = System.DateTime.UtcNow

        While (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested)
            ' Bail out if a newer generation has started
            If gen <> listenerGeneration Then Return

            Dim needShortDelay As System.Boolean = False

            Try
                ' Listener initialization
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.IgnoreWriteExceptions = True
                    With httpListener.TimeoutManager
                        ' Permissive timeouts for potentially slow handlers
                        .IdleConnection = System.TimeSpan.FromMinutes(10)
                        .HeaderWait = System.TimeSpan.FromSeconds(30)
                        .EntityBody = System.TimeSpan.FromMinutes(10)
                        .DrainEntityBody = System.TimeSpan.FromSeconds(30)
                        .MinSendBytesPerSecond = CType(0UL, System.UInt64)
                    End With
                    If Not httpListener.Prefixes.Contains(prefix) Then
                        httpListener.Prefixes.Add(prefix)
                    End If
                    httpListener.Start()
                    System.Diagnostics.Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    Try : httpListener.Close() : Catch : End Try
                    httpListener = Nothing
                    Continue While
                End If

                Dim ctx As System.Net.HttpListenerContext =
                    Await httpListener.GetContextAsync().ConfigureAwait(False)

                ' Progress heartbeat for watchdog
                lastListenerProgressUtc = System.DateTime.UtcNow

                Dim ctxLocal As System.Net.HttpListenerContext = ctx
                System.Threading.Tasks.Task.Run(
                    Async Function()
                        Dim resLocal As System.Net.HttpListenerResponse = Nothing
                        Try
                            Await HandleHttpRequest(ctxLocal).ConfigureAwait(False)
                        Catch
                            Try
                                resLocal = ctxLocal.Response
                                resLocal.StatusCode = 500
                                resLocal.KeepAlive = False
                                resLocal.Headers("Connection") = "close"
                                resLocal.SendChunked = False
                                Dim bufErr() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Internal server error.")
                                resLocal.ContentType = "text/plain; charset=utf-8"
                                resLocal.ContentLength64 = bufErr.LongLength
                                Using os As System.IO.Stream = resLocal.OutputStream
                                    os.Write(bufErr, 0, bufErr.Length)
                                End Using
                            Catch
                            Finally
                                Try
                                    If resLocal IsNot Nothing Then resLocal.Close()
                                Catch
                                End Try
                            End Try
                        Finally
                            ' Mark progress at the end of a handled request too
                            lastListenerProgressUtc = System.DateTime.UtcNow
                        End Try
                    End Function)

                ' Metrics block
                Dim now As System.DateTime = System.DateTime.UtcNow
                If (now - lastMetrics).TotalSeconds >= 10.0 Then
                    Dim gdi As System.UInt32 = GetGdiCount()
                    Dim usr As System.UInt32 = GetUserCount()
                    System.Diagnostics.Debug.WriteLine(
                        System.String.Format(
                            System.Globalization.CultureInfo.InvariantCulture,
                            "RES {0:HH:mm:ss}: GDI={1}  USER={2}",
                            now, gdi, usr))
                    If gdi >= 8000UI OrElse usr >= 8000UI Then
                        System.Diagnostics.Debug.WriteLine("WARN: High handle count – check for GDI/USER leaks.")
                    End If
                    lastMetrics = now
                End If

                consecutiveFailures = 0

            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
                needShortDelay = True

            Catch ex As System.Exception
                consecutiveFailures += 1
                System.Diagnostics.Debug.WriteLine(System.String.Concat("Listener error: ", ex.Message))
            End Try

            If needShortDelay AndAlso (Not token.IsCancellationRequested) Then
                Try
                    Await System.Threading.Tasks.Task.Delay(50, token).ConfigureAwait(False)
                Catch
                End Try
            End If

            If consecutiveFailures >= 10 AndAlso (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested) Then
                System.Diagnostics.Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then
                        Try : httpListener.Abort() : Catch : End Try
                        Try : httpListener.Close() : Catch : End Try
                    End If
                Catch
                Finally
                    httpListener = Nothing
                End Try
                consecutiveFailures = 0
                Try
                    Await System.Threading.Tasks.Task.Delay(5000, token).ConfigureAwait(False)
                Catch
                End Try
            End If
        End While
    End Function

    ''' <summary>
    ''' Starts hidden power notification window if not already active.
    ''' </summary>
    Private Sub StartPowerWatch()
        If powerWindow Is Nothing Then
            powerWindow = New PowerNotificationWindow(Me)
        End If
    End Sub

    ''' <summary>
    ''' Disposes hidden power notification window if active.
    ''' </summary>
    Private Sub StopPowerWatch()
        If powerWindow IsNot Nothing Then
            powerWindow.Dispose()
            powerWindow = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Handles system power mode changes (legacy event path). Delegates to suspend/resume logic.
    ''' </summary>
    Private Sub OnPowerModeChanged(ByVal sender As System.Object,
                                   ByVal e As Microsoft.Win32.PowerModeChangedEventArgs)
        If e Is Nothing Then Return

        Select Case e.Mode
            Case Microsoft.Win32.PowerModes.Suspend
                ' Graceful listener stop in the background
                System.Threading.ThreadPool.QueueUserWorkItem(
                    Sub(state As Object)
                        Try : ShutdownHttpListener().GetAwaiter().GetResult() : Catch : End Try
                    End Sub)

            Case Microsoft.Win32.PowerModes.Resume
                ' Avoid re-entrancy; delegate to unified resume path (userPresent:=True)
                If System.Threading.Interlocked.Exchange(restartingAfterResume, 1) = 1 Then Return
                System.Threading.ThreadPool.QueueUserWorkItem(
                    Sub(state As Object)
                        Try
                            HandlePowerResumeAsync(userPresent:=True)
                        Finally
                            System.Threading.Interlocked.Exchange(restartingAfterResume, 0)
                        End Try
                    End Sub)
        End Select
    End Sub

    ''' <summary>
    ''' Starts watchdog timer that monitors listener health and restarts if dead or faulted.
    ''' </summary>
    Private Sub StartListenerWatchdog()
        If watchdog IsNot Nothing Then Return

        watchdog = New System.Threading.Timer(
            Sub(stateObj As System.Object)
                Try
                    ' Skip during power transitions
                    If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then Return

                    ' Skip if in cooldown after resume
                    If IsInResumeCooldown() Then Return

                    ' Check active work
                    Dim inFlight As Integer = Threading.Interlocked.CompareExchange(activeRequests, 0, 0)
                    Dim jobsInFlight As Integer = Threading.Interlocked.CompareExchange(activeJobs, 0, 0)
                    If inFlight > 0 OrElse jobsInFlight > 0 Then
                        lastListenerProgressUtc = System.DateTime.UtcNow
                        Return
                    End If

                    ' Assess listener health without penalizing idle
                    Dim listenerIsDead As Boolean = False
                    Try
                        listenerIsDead = (httpListener Is Nothing) OrElse (Not httpListener.IsListening)
                    Catch
                        listenerIsDead = True
                    End Try

                    ' Simple timeout check (idle is OK; only use if listener isnot listening/faulted)
                    Dim age As Double = (System.DateTime.UtcNow - lastListenerProgressUtc).TotalSeconds

                    ' If the listener is gone or loop has faulted, restart cleanly
                    Dim loopFaulted As Boolean = False
                    Try
                        loopFaulted = (listenerTask IsNot Nothing AndAlso listenerTask.IsCompleted AndAlso listenerTask.IsFaulted)
                    Catch
                        loopFaulted = True
                    End Try

                    If (listenerIsDead OrElse loopFaulted) AndAlso Not isShuttingDown Then
                        System.Threading.Tasks.Task.Run(
                            Async Function()
                                Try
                                    ' Fully stop old instance to free the port
                                    Await ShutdownHttpListener(stopUiThread:=False)
                                Catch
                                End Try
                                Try
                                    StartupHttpListener()
                                Catch
                                End Try
                            End Function)
                        Return
                    End If

                    ' No restart solely due to idle

                Catch
                End Try
            End Sub,
            state:=Nothing,
            dueTime:=System.TimeSpan.FromSeconds(30),
            period:=System.TimeSpan.FromSeconds(10))
    End Sub

    ''' <summary>
    ''' Stops and disposes the listener watchdog timer.
    ''' </summary>
    Private Sub StopListenerWatchdog()
        Try
            If watchdog IsNot Nothing Then
                watchdog.Dispose()
            End If
        Catch
        Finally
            watchdog = Nothing
        End Try
    End Sub

    '-------------------------------------------------- Not used currently --------------------------------------------------

    Private lastSuccessfulOperationUtc As System.DateTime = System.DateTime.UtcNow
    Private Const StuckDetectionMinutes As Integer = 15  ' If no successful op for 15 min, consider stuck

    ''' <summary>
    ''' Performs recovery: cancels long-running jobs, resets listener if inactive, re-registers OLE message filter.
    ''' </summary>
    Private Sub PerformRecovery()
        System.Diagnostics.Debug.WriteLine("Performing recovery actions...")

        ' 1. Cancel all long-running jobs
        Dim now = System.DateTime.UtcNow
        Dim stuckJobs = jobMap.Values.Where(Function(j)
                                                Return j IsNot Nothing AndAlso
                   Not j.Tcs.Task.IsCompleted AndAlso
                   (now - j.CreatedUtc).TotalMinutes > 10
                                            End Function).ToList()

        For Each job In stuckJobs
            Try
                System.Diagnostics.Debug.WriteLine($"Cancelling stuck job: {job.Id}")
                If Not job.Cts.IsCancellationRequested Then
                    job.Cts.Cancel()
                End If
                ' Force completion after a grace period
                System.Threading.Tasks.Task.Run(Async Function()
                                                    Await System.Threading.Tasks.Task.Delay(5000)
                                                    If Not job.Tcs.Task.IsCompleted Then
                                                        job.Tcs.TrySetResult("Operation timed out during recovery.")
                                                    End If
                                                End Function)
            Catch
            End Try
        Next

        ' 2. Reset the HTTP listener if needed
        Dim listenerAge = (System.DateTime.UtcNow - lastListenerProgressUtc).TotalMinutes
        If listenerAge > 10 Then
            System.Diagnostics.Debug.WriteLine("Resetting HTTP listener due to inactivity")
            System.Threading.Tasks.Task.Run(Async Function()
                                                Try
                                                    Await ShutdownHttpListener(stopUiThread:=False)
                                                    Await System.Threading.Tasks.Task.Delay(2000)
                                                    StartupHttpListener()
                                                Catch
                                                End Try
                                            End Function)
        End If

        ' 3. Re-register OLE message filter (COM related)
        Try
            SwitchToUi(Sub()
                           OleMessageFilter.Revoke()
                           OleMessageFilter.Register()
                           EnableOleFilterFor(30000)
                       End Sub).Wait(5000) ' Do not wait indefinitely
        Catch
        End Try

        ' Mark recovery as successful
        lastSuccessfulOperationUtc = System.DateTime.UtcNow
    End Sub

    ''' <summary>
    ''' Cleans up orphaned or expired jobs based on completion state and TTL.
    ''' </summary>
    Private Sub CleanupOrphanedJobs()
        Dim now = System.DateTime.UtcNow
        Dim toRemove As New List(Of String)()

        For Each kv In jobMap
            Dim job = kv.Value
            If job Is Nothing Then
                toRemove.Add(kv.Key)
                Continue For
            End If

            ' Remove completed jobs older than 5 minutes
            If job.Tcs.Task.IsCompleted AndAlso (now - job.CreatedUtc).TotalMinutes > 5 Then
                toRemove.Add(kv.Key)
                ' Force-remove jobs older than TTL regardless of state
            ElseIf (now - job.CreatedUtc).TotalMinutes > JobTtlMinutes Then
                toRemove.Add(kv.Key)
                Try
                    If Not job.Cts.IsCancellationRequested Then job.Cts.Cancel()
                    If Not job.Tcs.Task.IsCompleted Then
                        job.Tcs.TrySetResult("Job expired.")
                    End If
                Catch
                End Try
            End If
        Next

        For Each key In toRemove
            Dim removed As LlmJob = Nothing
            If jobMap.TryRemove(key, removed) Then
                Try
                    removed?.Cts?.Dispose()
                Catch
                End Try
                System.Threading.Interlocked.Decrement(activeJobs)
            End If
        Next
    End Sub

End Class