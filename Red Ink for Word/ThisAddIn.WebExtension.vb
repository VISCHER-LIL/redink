' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Windows.Forms

Partial Public Class ThisAddIn


    Private httpListener As System.Net.HttpListener
    Private listenerTask As System.Threading.Tasks.Task   ' replaces the raw Thread
    Private isShuttingDown As Boolean = False

    '───────────────────────────────────────────────────────────────────────────
    ' Run a Sub on the UI thread and *wait* for it to finish.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ' Run a Func(Of T) on the UI thread and wait for its return value.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function


    Private Sub StartupHttpListener()
        ' fire-and-forget – no raw Thread needed
        listenerTask = StartHttpListener()      ' captures the returned Task
    End Sub

    Private Sub ShutdownHttpListener()
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub


    Private Async Function StartHttpListener() As System.Threading.Tasks.Task
        Const prefix As String = "http://127.0.0.1:12334/"   ' ← Word gets its own port
        Dim consecutiveFailures As Integer = 0

        While Not isShuttingDown
            Try
                ' ensure listener exists and is running
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    httpListener.Close()
                    httpListener = Nothing
                    Continue While                      ' next loop restarts it
                End If

                ' wait for one incoming request
                Dim ctx As System.Net.HttpListenerContext =
                Await httpListener.GetContextAsync().ConfigureAwait(False)

                ' handle the request (fire-and-forget)
                Call HandleHttpRequest(ctx) _
                .ContinueWith(
                    Sub(t)
                        If t.IsFaulted AndAlso t.Exception IsNot Nothing Then
                            Debug.WriteLine("HandleHttpRequest error: " &
                                            t.Exception.GetBaseException().Message)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)

                consecutiveFailures = 0                       ' success
            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
            Catch ex As System.Exception
                consecutiveFailures += 1
                Debug.WriteLine("Listener error: " & ex.Message)
            End Try

            ' recycle after too many consecutive errors
            If consecutiveFailures >= 10 AndAlso Not isShuttingDown Then
                Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then httpListener.Close()
                Catch
                End Try
                httpListener = Nothing
                consecutiveFailures = 0
                Await System.Threading.Tasks.Task.Delay(5000).ConfigureAwait(False)
            End If
        End While
    End Function


    Private Async Function HandleHttpRequest(
        ctx As System.Net.HttpListenerContext) _
        As System.Threading.Tasks.Task

        Dim req = ctx.Request
        Dim res = ctx.Response

        '─── CORS pre-flight────────────────────────────────────────────────────
        If req.HttpMethod = "OPTIONS" Then
            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
            res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
            res.StatusCode = 204 : res.Close() : Return
        End If

        '─── Read body (if any)─────────────────────────────────────────────────
        Dim body As String = ""
        If req.HasEntityBody Then
            Using rdr As New IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8)
                body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
            End Using
        End If

        '─── Dispatch to our add-in logic───────────────────────────────────────
        Dim responseText As String =
        Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)

        '─── Send response──────────────────────────────────────────────────────
        Dim buf = System.Text.Encoding.UTF8.GetBytes(responseText)
        res.ContentLength64 = buf.Length
        res.ContentType = "text/plain; charset=utf-8"
        res.AddHeader("Access-Control-Allow-Origin", "*")
        Using os = res.OutputStream
            Await os.WriteAsync(buf, 0, buf.Length).ConfigureAwait(False)
        End Using
        res.Close()
    End Function


    ' ---------------------------------------------------------------------------
    ' MAIN REQUEST DISPATCH (Word – only "redink_sendtoword")
    ' ---------------------------------------------------------------------------
    Private Async Function ProcessRequestInAddIn(
        body As String,
        rawUrl As String) _
        As System.Threading.Tasks.Task(Of String)

        ' guard clause – empty body
        If String.IsNullOrWhiteSpace(body) Then Return ""

        Dim j = Newtonsoft.Json.Linq.JObject.Parse(body)
        Dim cmd = j("Command")?.ToString()
        Dim textBody = j("Text")?.ToString()
        Dim sourceUrl = j("URL")?.ToString()

        Select Case cmd
        '───────────────────────────────────────────────────────────────────
            Case "redink_sendtoword"
                If String.IsNullOrWhiteSpace(textBody) Then Return ""

                ' Everything that touches Word must run on the UI thread
                Await SwitchToUi(Sub()

                                     Dim wdApp As Microsoft.Office.Interop.Word.Application =
                    Globals.ThisAddIn.Application

                                     Dim sel As Microsoft.Office.Interop.Word.Selection = wdApp.Selection

                                     wdApp.ScreenUpdating = False

                                     ' Sanitize inputs and insert without using TypeText
                                     Dim safeBody As String = If(textBody, String.Empty).Replace(ChrW(0), String.Empty)
                                     Dim safeUrl As String = If(sourceUrl, String.Empty).Replace(ChrW(0), String.Empty)
                                     Dim finalText As String = safeBody & " (" & safeUrl & ")"

                                     Dim rng As Word.Range = sel.Range.Duplicate
                                     If rng.Start = rng.End Then
                                         rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                         rng.InsertAfter(finalText)
                                     Else
                                         rng.Text = finalText  ' replace selected content safely
                                     End If

                                     sel.SetRange(rng.End, rng.End) ' place caret after inserted text
                                     wdApp.ScreenUpdating = True

                                     ' Release COM objects explicitly (good hygiene)
                                     System.Runtime.InteropServices.Marshal.ReleaseComObject(sel)
                                 End Sub)

                Return ""      ' nothing needs to be sent back in this scenario
        End Select

        Return ""              ' unknown command → no-op
    End Function



End Class
