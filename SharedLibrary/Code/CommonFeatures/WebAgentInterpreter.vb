' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.


Option Strict On
Option Explicit On

Imports Newtonsoft.Json.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Net
Imports System.Reflection
Imports System.Threading
Imports FxResources.System
Imports Markdig
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    Public NotInheritable Class WebAgentInterpreter
        Implements System.IDisposable

        Private ReadOnly _http As System.Net.Http.HttpClient
        Private ReadOnly _cookieContainer As System.Net.CookieContainer
        Private ReadOnly _handler As System.Net.Http.HttpClientHandler

        Private _headersDefault As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private _vars As New System.Collections.Generic.Dictionary(Of System.String, System.Object)(System.StringComparer.OrdinalIgnoreCase)
        Private _secrets As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private _baseUrl As System.String = System.String.Empty
        Private _userAgent As System.String = "WebAgentInterpreter/1.0"
        Private _lastResponseBody As System.String = System.String.Empty
        Private _lastResponseUrl As System.String = System.String.Empty
        Private _lastDoc As HtmlAgilityPack.HtmlDocument = Nothing
        Private _defaultTimeoutMs As System.Int32 = 30000

        Private _log As New System.Text.StringBuilder()
        Private _finalMarkdown As System.String = System.String.Empty

        Private _context As ISharedContext = Nothing
        Private _useSecondAPI As Boolean = False
        Private _autoselectModel As Boolean = False

        Private _dynamicExpand As Boolean = False
        Private Const MAX_DYNAMIC_FETCH As Integer = 10


        Private ReadOnly _failHard As Boolean = False   ' set True to stop on first non‑handled http failure
        Public ReadOnly Property LogText As String
            Get
                Return _log.ToString()
            End Get
        End Property

        Private _currentStepId As System.String = System.String.Empty
        Shared Sub New()
            Try
                System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12 Or
                System.Net.SecurityProtocolType.Tls11 Or
                System.Net.SecurityProtocolType.Tls
            Catch
            End Try
        End Sub

        Public Sub New()
            _cookieContainer = New System.Net.CookieContainer()
            _handler = New System.Net.Http.HttpClientHandler() With {
            .AllowAutoRedirect = True,
            .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate,
            .UseCookies = True,
            .CookieContainer = _cookieContainer
        }

            Try
#If NET6_0_OR_GREATER Or NET7_0_OR_GREATER Or NET8_0_OR_GREATER Then
                _handler.SslProtocols = System.Security.Authentication.SslProtocols.Tls12 Or System.Security.Authentication.SslProtocols.Tls13
#Else
                _handler.SslProtocols = System.Security.Authentication.SslProtocols.Tls12
#End If
            Catch
                Try
                    _handler.SslProtocols = System.Security.Authentication.SslProtocols.Tls12
                Catch
                End Try
            End Try

            _handler.ServerCertificateCustomValidationCallback =
            Function(req As System.Net.Http.HttpRequestMessage,
                     cert As System.Security.Cryptography.X509Certificates.X509Certificate2,
                     chain As System.Security.Cryptography.X509Certificates.X509Chain,
                     errors As System.Net.Security.SslPolicyErrors)

                If errors <> System.Net.Security.SslPolicyErrors.None Then
                    Log($"[TLS] Host={req?.RequestUri?.Host} Errors={errors} Subject={cert?.Subject}")
                    If chain IsNot Nothing Then
                        For Each st In chain.ChainStatus
                            Log($"[TLS] ChainStatus: {st.Status} {st.StatusInformation}")
                        Next
                    End If
                End If

                Return errors = System.Net.Security.SslPolicyErrors.None
            End Function

            _http = New System.Net.Http.HttpClient(_handler)
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            _http.Dispose()
            _handler.Dispose()
        End Sub

        Private Function _FormatVarValue(value As Object) As String
            If value Is Nothing Then Return "(null)"
            Try
                If TypeOf value Is String Then
                    Dim s = DirectCast(value, String)
                    If s.Length > 400 Then Return s.Substring(0, 400) & "…"
                    Return s
                End If
                If TypeOf value Is Newtonsoft.Json.Linq.JToken Then
                    Dim t = DirectCast(value, Newtonsoft.Json.Linq.JToken)
                    Dim txt = t.ToString(Newtonsoft.Json.Formatting.None)
                    If txt.Length > 400 Then txt = txt.Substring(0, 400) & "…"
                    Return txt
                End If
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(value)
                If json.Length > 400 Then json = json.Substring(0, 400) & "…"
                Return json
            Catch
                Return value.ToString()
            End Try
        End Function


        Private Function GetDebugFlag(name As String, Optional defaultValue As Boolean = False) As Boolean
            Try
                Dim o As Object = Nothing
                If _vars IsNot Nothing AndAlso _vars.TryGetValue(name, o) AndAlso o IsNot Nothing Then
                    Dim b As Boolean
                    If Boolean.TryParse(o.ToString(), b) Then Return b
                End If
            Catch
            End Try
            Return defaultValue
        End Function

        Private _debugLogPath As String = Nothing
        Private _debugInitialized As Boolean = False

        Private Sub InitDebugLogIfNeeded()
            If _debugInitialized Then Return
            If Not GetDebugFlag("debug") AndAlso Not GetDebugFlag("debug_allAttempts") Then Return
            Try
                ' Prefer the current user's Desktop for the debug log
                Dim desktopPath As String = ""
                Try
                    desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Catch
                End Try

                If String.IsNullOrWhiteSpace(desktopPath) OrElse Not System.IO.Directory.Exists(desktopPath) Then
                    ' Fallback: application base directory if Desktop is unavailable
                    desktopPath = AppDomain.CurrentDomain.BaseDirectory
                End If

                _debugLogPath = System.IO.Path.Combine(desktopPath, "RI_Debug_Webagent.txt")

                System.IO.File.WriteAllText(_debugLogPath,
                                 $"# WebAgent Debug Log {DateTime.Now:O}{Environment.NewLine}" &
                                 $"# Machine: {Environment.MachineName}{Environment.NewLine}" &
                                 $"# User   : {Environment.UserName}{Environment.NewLine}" &
                                 $"# Desktop: {desktopPath}{Environment.NewLine}{Environment.NewLine}")

                _debugInitialized = True
            Catch ex As Exception
                ' Silently ignore any issues initializing the log
            End Try
        End Sub




        Private Function DebugEnabled() As Boolean
            Return GetDebugFlag("debug") OrElse GetDebugFlag("debug_allAttempts") _
        OrElse GetDebugFlag("debug_substeps") OrElse GetDebugFlag("debug_var_changes") _
        OrElse GetDebugFlag("debug_include_script") OrElse GetDebugFlag("debug_summary")
        End Function


        Private Sub DebugLogScriptOverview(root As Newtonsoft.Json.Linq.JObject)
            If Not GetDebugFlag("debug_include_script") Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            Try
                Dim scrub = CType(root.DeepClone(), Newtonsoft.Json.Linq.JObject)

                ' Mask secrets inside env.secrets
                Dim env = TryCast(scrub("env"), JObject)
                Dim secrets = TryCast(env?("secrets"), JObject)
                If secrets IsNot Nothing Then
                    For Each p In secrets.Properties()
                        p.Value = "***"
                    Next
                End If

                ' Build step table
                Dim steps = TryCast(scrub("steps"), JArray)
                Dim sb As New StringBuilder()
                sb.AppendLine("=== SCRIPT OVERVIEW ===")
                If steps IsNot Nothing Then
                    Dim i As Integer = 0
                    For Each st As JObject In steps
                        Dim sid = st.Value(Of String)("id")
                        Dim cmd = st.Value(Of String)("command")
                        sb.AppendLine($"{i,3}: id='{sid}' cmd='{cmd}'")
                        i += 1
                    Next
                End If
                sb.AppendLine("=== RAW (MASKED) SCRIPT JSON ===")
                Dim jsonText = scrub.ToString(Newtonsoft.Json.Formatting.Indented)
                If jsonText.Length > 120000 Then
                    jsonText = jsonText.Substring(0, 120000) & vbCrLf & "...(truncated for debug)..."
                End If
                WriteDebug(sb.ToString())
                WriteDebug(jsonText)
            Catch ex As Exception
                WriteDebug("[debug_include_script] error: " & ex.Message)
            End Try
        End Sub

        ' -------------------- NEW: Variable Change Logging --------------------
        Private Sub DebugLogVarChange(name As String, oldVal As Object, newVal As Object)
            If Not GetDebugFlag("debug_var_changes") Then Return
            If Not DebugEnabled() Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            Try
                Dim oldTxt = If(oldVal Is Nothing, "(null)", _FormatVarValue(oldVal))
                Dim newTxt = If(newVal Is Nothing, "(null)", _FormatVarValue(newVal))
                If oldTxt.Length > 300 Then oldTxt = oldTxt.Substring(0, 300) & "…"
                If newTxt.Length > 300 Then newTxt = newTxt.Substring(0, 300) & "…"
                WriteDebug($"[var_change] {name} :: {MaskSecrets(oldTxt)}  ==>  {MaskSecrets(newTxt)}")
            Catch
            End Try
        End Sub


        Private Sub SafeStoreVar_DebugPatch(varName As String, value As Object)
            If String.IsNullOrWhiteSpace(varName) Then Exit Sub
            Dim hadOld As Boolean = _vars.ContainsKey(varName)
            Dim oldVal As Object = If(hadOld, _vars(varName), Nothing)

            ' (Reuse original SafeStoreVar logic)
            If value Is Nothing Then
                If hadOld Then
                    DebugLogVarChange(varName, oldVal, Nothing)
                    _vars.Remove(varName)
                End If
                Exit Sub
            End If
            Try
                Select Case True
                    Case TypeOf value Is JToken
                        _vars(varName) = DirectCast(value, JToken).DeepClone()
                    Case TypeOf value Is String, TypeOf value Is Integer, TypeOf value Is Boolean, TypeOf value Is Long, TypeOf value Is Double, TypeOf value Is Decimal
                        _vars(varName) = value
                    Case Else
                        Dim jt = JToken.FromObject(value)
                        _vars(varName) = jt.DeepClone().ToObject(Of Object)()
                End Select
            Catch
                _vars(varName) = value
            End Try
            DebugLogVarChange(varName, oldVal, _vars(varName))
        End Sub

        Private Sub DebugLogSubStepStart(sid As String, cmd As String, attempt As Integer, maxRetry As Integer, parms As JObject)
            If Not GetDebugFlag("debug_substeps") Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            Try
                Dim pTxt As String = ""
                If parms IsNot Nothing Then
                    pTxt = parms.ToString(Newtonsoft.Json.Formatting.None)
                    If pTxt.Length > 800 Then pTxt = pTxt.Substring(0, 800) & "…"
                End If
                WriteDebug($"[substep:start] id={sid} cmd={cmd} attempt={attempt + 1}/{Math.Max(1, maxRetry + 1)} parms={MaskSecrets(pTxt)}")
            Catch
            End Try
        End Sub

        Private Sub DebugLogSubStepResult(sid As String, cmd As String, durationMs As Long, success As Boolean, ex As Exception, result As Object)
            If Not GetDebugFlag("debug_substeps") Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            Try
                Dim resTxt = ""
                If result IsNot Nothing Then
                    Try
                        If TypeOf result Is JToken Then
                            resTxt = DirectCast(result, JToken).ToString(Newtonsoft.Json.Formatting.None)
                        Else
                            resTxt = Newtonsoft.Json.JsonConvert.SerializeObject(result)
                        End If
                    Catch
                        resTxt = result.ToString()
                    End Try
                    If resTxt.Length > 800 Then resTxt = resTxt.Substring(0, 800) & "…"
                End If
                If ex IsNot Nothing Then
                    Dim exTxt = ex.ToString()
                    If exTxt.Length > 1200 Then exTxt = exTxt.Substring(0, 1200) & "…"
                    WriteDebug($"[substep:end] id={sid} cmd={cmd} success={success} dur_ms={durationMs} ERROR={MaskSecrets(ex.Message)} STACK={MaskSecrets(exTxt)}")
                Else
                    WriteDebug($"[substep:end] id={sid} cmd={cmd} success={success} dur_ms={durationMs} result={MaskSecrets(resTxt)}")
                End If
            Catch
            End Try
        End Sub


        Private Sub DebugLogStepStart(sid As String, cmd As String, index As Integer, total As Integer)
            If Not DebugEnabled() Then Return
            If Not (GetDebugFlag("debug") OrElse GetDebugFlag("debug_allAttempts") OrElse GetDebugFlag("debug_substeps")) Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            WriteDebug($"[step:start] {index + 1}/{total} id={sid} cmd={cmd}")
        End Sub

        Private Sub DebugLogStepEnd(sid As String, cmd As String, success As Boolean, durationMs As Long, ex As Exception)
            If Not DebugEnabled() Then Return
            If Not (GetDebugFlag("debug") OrElse GetDebugFlag("debug_allAttempts") OrElse GetDebugFlag("debug_substeps")) Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            If ex IsNot Nothing Then
                Dim exMsg = ex.Message
                WriteDebug($"[step:end] id={sid} cmd={cmd} success={success} dur_ms={durationMs} error={MaskSecrets(exMsg)}")
            Else
                WriteDebug($"[step:end] id={sid} cmd={cmd} success={success} dur_ms={durationMs}")
            End If
        End Sub


        Private Sub DebugFinalSummary(totalSteps As Integer,
                            executedSteps As Integer,
                            abortedStepId As String,
                            elapsedMs As Long)
            If Not GetDebugFlag("debug_summary") Then Return
            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Return
            Try
                WriteDebug("=== EXECUTION SUMMARY ===")
                WriteDebug($"total_steps        = {totalSteps}")
                WriteDebug($"executed_steps     = {executedSteps}")
                WriteDebug($"aborted_step_id    = {If(String.IsNullOrWhiteSpace(abortedStepId), "(none)", abortedStepId)}")
                WriteDebug($"total_elapsed_ms   = {elapsedMs}")
                ' Optionally list last known variables (masked & truncated)
                Dim listVars As New List(Of String)
                For Each kv In _vars
                    Dim vTxt = _FormatVarValue(kv.Value)
                    If vTxt.Length > 180 Then vTxt = vTxt.Substring(0, 180) & "…"
                    listVars.Add($"{kv.Key}={MaskSecrets(vTxt)}")
                Next
                WriteDebug("vars_snapshot: " & String.Join("; ", listVars))
                WriteDebug("=== END SUMMARY ===")
            Catch
            End Try
        End Sub




        Private Function GetFlag(name As String, defaultVal As Boolean) As Boolean
            Return GetDebugFlag(name, defaultVal)
        End Function

        Private Sub WriteDebug(lines As String)
            If Not _debugInitialized Then Return
            Try
                System.IO.File.AppendAllText(_debugLogPath, lines & Environment.NewLine)
            Catch
            End Try
        End Sub

        ' === Replace existing DebugSnapshot implementation with this file-writing version ===
        Private Sub DebugSnapshot(stepId As String,
                      command As String,
                      attemptNumber As Integer,
                      maxRetry As Integer,
                      success As Boolean,
                      willRetry As Boolean,
                      lastEx As Exception)

            ' Activate only if flags set
            If Not GetDebugFlag("debug") AndAlso Not GetDebugFlag("debug_allAttempts") Then Exit Sub

            ' If not all attempts requested, only log final outcome
            If Not GetDebugFlag("debug_allAttempts") Then
                If Not success AndAlso willRetry Then Exit Sub
            End If

            InitDebugLogIfNeeded()
            If Not _debugInitialized Then Exit Sub

            Try
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine(New String("-"c, 70))
                sb.AppendLine($"Timestamp : {DateTime.Now:O}")
                sb.AppendLine($"Step      : {stepId}")
                sb.AppendLine($"Command   : {command}")
                sb.AppendLine($"Attempt   : {attemptNumber}/{maxRetry + 1}")
                sb.AppendLine($"Success   : {success}")
                sb.AppendLine($"WillRetry : {willRetry}")

                If Not String.IsNullOrWhiteSpace(_lastResponseUrl) Then
                    sb.AppendLine($"LastURL   : {MaskSecrets(_lastResponseUrl)}")
                End If

                If Not String.IsNullOrEmpty(_lastResponseBody) Then
                    Dim bodyPreview = System.Text.RegularExpressions.Regex.Replace(_lastResponseBody, "\s+", " ").Trim()
                    If bodyPreview.Length > 800 Then bodyPreview = bodyPreview.Substring(0, 800) & "…"
                    sb.AppendLine("LastBody  : " & MaskSecrets(bodyPreview))
                End If

                ' LLM raw (if relevant) – stored by CmdLlmAnalyzeAsync as lastLlm_raw
                If command IsNot Nothing AndAlso command.IndexOf("llm", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Dim rawObj As Object = Nothing
                    If _vars.TryGetValue("lastLlm_raw", rawObj) AndAlso rawObj IsNot Nothing Then
                        Dim rawTxt = rawObj.ToString()
                        Dim preview = rawTxt
                        If preview.Length > 5000 Then preview = preview.Substring(0, 5000) & "…"
                        sb.AppendLine($"LLM Raw Len: {rawTxt.Length}")
                        sb.AppendLine("LLM Raw:")
                        sb.AppendLine(preview)
                    End If
                End If

                If lastEx IsNot Nothing Then
                    Dim exText = lastEx.ToString()
                    If exText.Length > 2000 Then exText = exText.Substring(0, 2000) & "…"
                    sb.AppendLine("Exception : " & lastEx.GetType().Name)
                    sb.AppendLine("ExMessage : " & lastEx.Message)
                    sb.AppendLine("ExDetail  : " & exText)
                End If

                Dim hideLlmVars As Boolean = (command Is Nothing) OrElse (command.IndexOf("llm", StringComparison.OrdinalIgnoreCase) < 0)

                sb.AppendLine("Variables:")
                For Each kv In _vars
                    If hideLlmVars AndAlso (String.Equals(kv.Key, "lastLlm", StringComparison.OrdinalIgnoreCase) _
                                  OrElse String.Equals(kv.Key, "lastLlm_page_url", StringComparison.OrdinalIgnoreCase)) Then
                        Continue For
                    End If
                    Dim valText = _FormatVarValue(kv.Value)
                    If valText.Length > 600 Then valText = valText.Substring(0, 600) & "…"
                    sb.AppendLine($"  {kv.Key} = {MaskSecrets(valText)}")
                Next

                WriteDebug(sb.ToString())
            Catch
                ' Ignore logging failures
            End Try
        End Sub

        Public Async Function RunAsync(scriptJson As System.String,
                                   context As ISharedContext,
                                   Optional useSecondAPI As Boolean = False,
                                   Optional autoselectModel As Boolean = False,
                                   Optional cancel As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task(Of System.String)
            _context = context
            _useSecondAPI = useSecondAPI
            _autoselectModel = autoselectModel
            Return Await RunAsync(scriptJson, cancel)
        End Function

        Public Async Function RunAsync(scriptJson As System.String,
                                   context As ISharedContext,
                                   Optional useSecondAPI As Boolean = False,
                                   Optional cancel As System.Threading.CancellationToken = Nothing,
                                   Optional silent As Boolean = False) As System.Threading.Tasks.Task(Of System.String)
            _context = context
            _useSecondAPI = useSecondAPI
            Return Await RunAsync(scriptJson, cancel, silent)
        End Function


        Public Async Function RunAsync(scriptJson As System.String,
                                   Optional cancel As System.Threading.CancellationToken = Nothing,
                                   Optional silent As Boolean = False) As System.Threading.Tasks.Task(Of System.String)

            Dim root As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(scriptJson)

            Dim meta = TryCast(root("meta"), Newtonsoft.Json.Linq.JObject)
            If meta IsNot Nothing Then
                Dim dflt = meta.Value(Of System.Nullable(Of System.Int32))("default_timeout_ms")
                If dflt.HasValue Then _defaultTimeoutMs = dflt.Value
                Dim ua = meta.Value(Of System.String)("user_agent")
                If ua IsNot Nothing Then _userAgent = ua
            End If

            Dim env = TryCast(root("env"), Newtonsoft.Json.Linq.JObject)
            If env IsNot Nothing Then
                Dim bu = env.Value(Of System.String)("base_url")
                If bu IsNot Nothing Then
                    _baseUrl = bu
                    _vars("base_url") = bu
                End If
                Dim headers = TryCast(env("headers"), Newtonsoft.Json.Linq.JObject)
                If headers IsNot Nothing Then
                    For Each prop In headers.Properties()
                        _headersDefault(prop.Name) = prop.Value.ToString()
                    Next
                End If
                Dim secrets = TryCast(env("secrets"), JObject)
                If secrets IsNot Nothing Then
                    For Each prop In secrets.Properties()
                        Dim raw = prop.Value.ToString()
                        Dim resolved = ResolveSecret(raw)
                        _secrets(prop.Name) = resolved
                        _vars(prop.Name) = resolved
                    Next
                End If
                Dim variables = TryCast(env("variables"), JObject)
                If variables IsNot Nothing Then
                    For Each prop In variables.Properties()
                        _vars(prop.Name) = prop.Value.ToObject(Of System.Object)()
                    Next
                End If
            End If

            _http.DefaultRequestHeaders.UserAgent.Clear()
            _http.DefaultRequestHeaders.UserAgent.ParseAdd(_userAgent)
            For Each kv In _headersDefault
                SafeSetHeader(_http.DefaultRequestHeaders, kv.Key, kv.Value)
            Next

            Dim steps = TryCast(root("steps"), Newtonsoft.Json.Linq.JArray)
            If steps Is Nothing OrElse steps.Count = 0 Then
                Throw New System.Exception("No 'steps' array found in script.")
            End If

            InitDebugLogIfNeeded()
            DebugLogScriptOverview(root)
            Dim totalStepsCount = steps.Count    ' (NEW) total step count
            Dim executedSteps = 0
            Dim globalSw = System.Diagnostics.Stopwatch.StartNew()

            Dim idToIndex As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
            For i As System.Int32 = 0 To steps.Count - 1
                Dim sidMap = steps(i).Value(Of System.String)("id")
                If Not System.String.IsNullOrWhiteSpace(sidMap) Then idToIndex(sidMap) = i
            Next

            Dim stepIndex As System.Int32 = 0

            Try
                While stepIndex < steps.Count
                    cancel.ThrowIfCancellationRequested()

                    Dim stepObj = TryCast(steps(stepIndex), Newtonsoft.Json.Linq.JObject)
                    Dim sid As System.String = stepObj.Value(Of System.String)("id")
                    Dim command As System.String = stepObj.Value(Of System.String)("command")
                    Dim timeoutMs As System.Int32 = stepObj.Value(Of System.Nullable(Of System.Int32))("timeout_ms").GetValueOrDefault(_defaultTimeoutMs)

                    _currentStepId = sid

                    ' Step timing + start log (NEW)
                    Dim stepSw = System.Diagnostics.Stopwatch.StartNew()
                    DebugLogStepStart(sid, command, stepIndex, totalStepsCount)
                    Dim stepSucceeded As Boolean = False

                    Try
                        If command IsNot Nothing AndAlso command.IndexOf("llm", StringComparison.OrdinalIgnoreCase) < 0 Then
                            If _vars.ContainsKey("lastLlm_raw") Then _vars.Remove("lastLlm_raw")
                            If GetFlag("debug_clear_llm_state", False) Then
                                If _vars.ContainsKey("lastLlm") Then _vars.Remove("lastLlm")
                                If _vars.ContainsKey("lastLlm_page_url") Then _vars.Remove("lastLlm_page_url")
                            End If
                        End If

                        Dim retry = TryCast(stepObj("retry"), Newtonsoft.Json.Linq.JObject)
                        Dim maxRetry As System.Int32 = If(retry IsNot Nothing, retry.Value(Of System.Nullable(Of System.Int32))("max").GetValueOrDefault(0), 0)
                        Dim retryDelay As System.Int32 = If(retry IsNot Nothing, retry.Value(Of System.Nullable(Of System.Int32))("delay_ms").GetValueOrDefault(1000), 1000)
                        Dim backoff As System.Double = If(retry IsNot Nothing, retry.Value(Of System.Nullable(Of System.Double))("backoff").GetValueOrDefault(2.0R), 2.0R)

                        Dim onError = TryCast(stepObj("on_error"), Newtonsoft.Json.Linq.JObject)
                        Dim onErrAction As System.String = If(onError IsNot Nothing, onError.Value(Of System.String)("action"), Nothing)
                        Dim onErrGoto As System.String = If(onError IsNot Nothing, onError.Value(Of System.String)("goto"), Nothing)

                        Dim guardObj = TryCast(stepObj("guard"), Newtonsoft.Json.Linq.JObject)
                        If guardObj IsNot Nothing Then
                            Dim condition = guardObj.Value(Of System.String)("if")
                            If Not System.String.IsNullOrWhiteSpace(condition) Then
                                Dim eval = EvalCondition(condition)
                                If Not eval Then
                                    Dim elseId = guardObj.Value(Of System.String)("else_goto")
                                    Log($"[{sid}] Guard false -> skip. else_goto={elseId}")
                                    stepSw.Stop()
                                    DebugLogStepEnd(sid, command, True, stepSw.ElapsedMilliseconds, Nothing) ' Skipped counts as successful skip
                                    executedSteps += 1
                                    If Not System.String.IsNullOrWhiteSpace(elseId) AndAlso idToIndex.ContainsKey(elseId) Then
                                        stepIndex = idToIndex(elseId)
                                        Continue While
                                    Else
                                        stepIndex += 1
                                        Continue While
                                    End If
                                End If
                            End If
                        End If

                        Dim waitFor = TryCast(stepObj("wait_for"), Newtonsoft.Json.Linq.JObject)
                        If waitFor IsNot Nothing AndAlso waitFor.Value(Of System.String)("type") = "time" Then
                            Dim msDelay = waitFor.Value(Of System.Nullable(Of System.Int32))("timeout_ms").GetValueOrDefault(0)
                            If msDelay > 0 Then
                                Await System.Threading.Tasks.Task.Delay(msDelay, cancel)
                            End If
                        End If

                        If Not silent Then
                            Try
                                Dim parmsPreview = TryCast(stepObj("params"), Newtonsoft.Json.Linq.JObject)
                                Dim progressMsg As String = Nothing
                                Select Case command.ToLowerInvariant()
                                    Case "open_url", "http_request"
                                        Dim u = ""
                                        If parmsPreview IsNot Nothing Then u = ExpandTemplates(parmsPreview.Value(Of System.String)("url"))
                                        progressMsg = If(System.String.IsNullOrWhiteSpace(u), "Loading the library ...", "Loading the library from " & u)
                                    Case "download_url"
                                        Dim u = ""
                                        If parmsPreview IsNot Nothing Then u = ExpandTemplates(parmsPreview.Value(Of System.String)("url"))
                                        progressMsg = If(System.String.IsNullOrWhiteSpace(u), "Downloading resource ...", "Downloading resource from " & u)
                                    Case "extract_text", "extract_html", "extract_attribute", "find"
                                        progressMsg = "Extracting data (" & command & ") ..."
                                    Case "llm_analyze", "llm", "llmanalyze"
                                        Dim urlDisplay As System.String = System.String.Empty
                                        If Not System.String.IsNullOrWhiteSpace(_lastResponseUrl) Then
                                            Dim shortUrl As System.String = _lastResponseUrl
                                            If shortUrl.Length > 120 Then shortUrl = shortUrl.Substring(0, 117) & "..."
                                            urlDisplay = " [" & shortUrl & "]"
                                        End If
                                        progressMsg = "Analyzing content (LLM) ..." & urlDisplay
                                    Case "render_report"
                                        progressMsg = "Rendering report ..."
                                    Case Else
                                        progressMsg = "Executing step '" & command & "' ..."
                                End Select
                                If Not System.String.IsNullOrWhiteSpace(sid) Then progressMsg = "[" & sid & "] " & progressMsg
                                InfoBox.ShowInfoBox(progressMsg)
                            Catch
                            End Try
                        End If

                        Dim attempt As Integer = 0
                        Dim success As Boolean = False
                        Dim resultValue As Object = Nothing
                        Dim lastEx As Exception = Nothing

                        Do
                            Dim plannedDelay As Integer = 0
                            lastEx = Nothing
                            Try
                                Log($"[{sid}] {command}")
                                Dim parms = TryCast(stepObj("params"), Newtonsoft.Json.Linq.JObject)
                                Dim sw = System.Diagnostics.Stopwatch.StartNew()

                                Select Case command.ToLowerInvariant()
                                    Case "set_user_agent" : resultValue = CmdSetUserAgent(parms)
                                    Case "set_headers" : resultValue = CmdSetHeaders(parms)
                                    Case "set_cookies" : resultValue = CmdSetCookies(parms)
                                    Case "open_url" : resultValue = Await CmdOpenUrlAsync(parms, timeoutMs, cancel)
                                    Case "navigate" : Throw New System.Exception("navigate is not available in HTTP mode. Use open_url.")
                                    Case "wait" : resultValue = Await CmdWaitAsync(parms, cancel)
                                    Case "find" : resultValue = CmdFind(parms)
                                    Case "extract_text" : resultValue = CmdExtractText(parms)
                                    Case "extract_html" : resultValue = CmdExtractHtml(parms)
                                    Case "extract_attribute" : resultValue = CmdExtractAttribute(parms)
                                    Case "download_url" : resultValue = Await CmdDownloadUrlAsync(parms, cancel)
                                    Case "save_file" : resultValue = CmdSaveFile(parms)
                                    Case "read_file" : resultValue = CmdReadFile(parms)
                                    Case "http_request" : resultValue = Await CmdHttpRequestAsync(parms, timeoutMs, cancel)
                                    Case "set_var" : resultValue = CmdSetVar(parms)
                                    Case "template" : resultValue = CmdTemplate(parms)
                                    Case "if" : resultValue = Await CmdIfAsync(parms, cancel)
                                    Case "foreach" : resultValue = Await CmdForEachAsync(parms, cancel)
                                    Case "render_report" : resultValue = CmdRenderReport(parms)
                                    Case "delete_file" : resultValue = CmdDeleteFile(parms)
                                    Case "send_email_report" : resultValue = CmdSendEmailReport(parms)
                                    Case "log" : resultValue = CmdLog(parms)
                                    Case "enable_dynamic" : CmdEnableDynamic(DirectCast(stepObj, JObject))
                                    Case "array_push" : resultValue = CmdArrayPush(parms)
                                    Case "llm_analyze", "llm", "llmanalyze" : resultValue = Await CmdLlmAnalyzeAsync(parms, timeoutMs, cancel)
                                    Case Else
                                        Throw New System.Exception($"Unknown command: {command}")
                                End Select

                                sw.Stop()
                                Log($"[{sid}] OK in {sw.ElapsedMilliseconds} ms")

                                Dim assign = TryCast(stepObj("assign"), Newtonsoft.Json.Linq.JObject)
                                If assign IsNot Nothing Then
                                    Dim varName = assign.Value(Of String)("var")
                                    Dim path = assign.Value(Of String)("path")
                                    If Not String.IsNullOrWhiteSpace(varName) Then
                                        Dim toStore As Object = resultValue
                                        If Not String.IsNullOrWhiteSpace(path) AndAlso resultValue IsNot Nothing Then
                                            Dim tokenObj = Newtonsoft.Json.Linq.JToken.FromObject(resultValue)
                                            Dim sel = tokenObj.SelectToken(path)
                                            If sel IsNot Nothing Then
                                                toStore = sel.ToObject(Of Object)()
                                            Else
                                                toStore = Nothing
                                            End If
                                        End If
                                        SafeStoreVar(varName, toStore)
                                    End If
                                End If

                                success = True

                            Catch exInner As Exception
                                lastEx = exInner
                                Dim detail = exInner.ToString()
                                Log($"[{sid}] Error (attempt {attempt + 1}/{maxRetry + 1}): {detail}")
                                Dim httpEx = TryCast(exInner, System.Net.Http.HttpRequestException)
                                If httpEx IsNot Nothing AndAlso httpEx.InnerException IsNot Nothing Then
                                    Log($"[{sid}] Inner: {httpEx.InnerException.GetType().Name}: {httpEx.InnerException.Message}")
                                End If

                                If attempt < maxRetry Then
                                    plannedDelay = CInt(retryDelay * Math.Pow(backoff, attempt))
                                    Log($"[{sid}] Will retry in {plannedDelay} ms …")
                                Else
                                    If onErrAction IsNot Nothing Then
                                        Select Case onErrAction
                                            Case "continue"
                                                success = True
                                            Case "goto"
                                                If Not String.IsNullOrWhiteSpace(onErrGoto) AndAlso idToIndex.ContainsKey(onErrGoto) Then
                                                    stepIndex = idToIndex(onErrGoto)
                                                    success = True
                                                Else
                                                    Throw New System.Exception($"on_error.goto target '{onErrGoto}' not found.", exInner)
                                                End If
                                            Case "abort"
                                                Throw
                                            Case "retry"
                                                Throw
                                            Case Else
                                                Throw
                                        End Select
                                    Else
                                        If _failHard Then Throw
                                    End If
                                End If
                            Finally
                                Dim willRetry As Boolean = (Not success AndAlso attempt < maxRetry)
                                DebugSnapshot(sid, command, attempt + 1, maxRetry, success, willRetry, lastEx)
                            End Try

                            attempt += 1
                            If Not success AndAlso attempt <= maxRetry AndAlso plannedDelay > 0 Then
                                Await System.Threading.Tasks.Task.Delay(plannedDelay, cancel)
                            End If
                        Loop While Not success AndAlso attempt <= maxRetry

                        Dim wf = TryCast(stepObj("wait_for"), Newtonsoft.Json.Linq.JObject)
                        If wf IsNot Nothing Then
                            Dim tType = wf.Value(Of System.String)("type")
                            If tType = "url" Then
                                Dim expected = wf.Value(Of System.String)("value")
                                If Not System.String.IsNullOrEmpty(expected) AndAlso _lastResponseUrl IsNot Nothing Then
                                    If _lastResponseUrl.IndexOf(expected, System.StringComparison.OrdinalIgnoreCase) = -1 Then
                                        Log($"[wait_for:url] Expected partial URL '{expected}' not found in '{_lastResponseUrl}'.")
                                    End If
                                End If
                            ElseIf tType = "selector" Then
                                Dim selObj = TryCast(wf("selector"), Newtonsoft.Json.Linq.JObject)
                                If selObj IsNot Nothing Then
                                    Dim nodes = ResolveSelector(selObj)
                                    If nodes Is Nothing OrElse nodes.Count = 0 Then
                                        Log("[wait_for:selector] Selector returned no matches.")
                                    End If
                                End If
                            End If
                        End If

                        stepSucceeded = True
                        executedSteps += 1
                        stepSw.Stop()
                        DebugLogStepEnd(sid, command, True, stepSw.ElapsedMilliseconds, Nothing)

                        stepIndex += 1

                    Catch exStep As Exception
                        stepSw.Stop()
                        DebugLogStepEnd(sid, command, False, stepSw.ElapsedMilliseconds, exStep)
                        Throw
                    End Try
                End While

                globalSw.Stop()
                DebugFinalSummary(totalStepsCount, executedSteps, Nothing, globalSw.ElapsedMilliseconds)

            Catch
                globalSw.Stop()
                DebugFinalSummary(totalStepsCount, executedSteps, _currentStepId, globalSw.ElapsedMilliseconds)
                Throw
            End Try

            If Not silent Then
                Try
                    InfoBox.ShowInfoBox("", 1)
                Catch
                End Try
            End If

            If Not System.String.IsNullOrWhiteSpace(_finalMarkdown) Then
                Return _finalMarkdown
            Else
                Return "# Execution finished" & System.Environment.NewLine & System.Environment.NewLine &
                   "````log" & System.Environment.NewLine & _log.ToString() & System.Environment.NewLine & "````"
            End If
        End Function

#Region "Command Implementations"

        Private Function CmdSetUserAgent(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim ua = ExpandTemplates(parms.Value(Of System.String)("user_agent"))
            If System.String.IsNullOrWhiteSpace(ua) Then Throw New System.Exception("user_agent missing.")
            _userAgent = ua
            _http.DefaultRequestHeaders.UserAgent.Clear()
            _http.DefaultRequestHeaders.UserAgent.ParseAdd(ua)
            Return New With {.user_agent = ua}
        End Function

        Private Function CmdSetHeaders(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim mode = parms.Value(Of System.String)("mode")
            Dim headers = TryCast(parms("headers"), Newtonsoft.Json.Linq.JObject)
            If headers Is Nothing Then Throw New System.Exception("headers missing.")
            If System.String.Equals(mode, "replace", System.StringComparison.OrdinalIgnoreCase) Then
                _headersDefault.Clear()
                _http.DefaultRequestHeaders.Clear()
                _http.DefaultRequestHeaders.UserAgent.ParseAdd(_userAgent)
            End If
            For Each p In headers.Properties()
                Dim val = ExpandTemplates(p.Value.ToString())
                _headersDefault(p.Name) = val
                SafeSetHeader(_http.DefaultRequestHeaders, p.Name, val)
            Next
            Return New With {.headers = _headersDefault}
        End Function

        Private Function CmdSetCookies(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim arr = TryCast(parms("cookies"), Newtonsoft.Json.Linq.JArray)
            If arr Is Nothing Then Throw New System.Exception("cookies missing.")
            For Each c As Newtonsoft.Json.Linq.JObject In arr
                Dim name = c.Value(Of System.String)("name")
                Dim value = ExpandTemplates(c.Value(Of System.String)("value"))
                Dim domain = c.Value(Of System.String)("domain")
                Dim path = c.Value(Of System.String)("path")
                If System.String.IsNullOrWhiteSpace(name) OrElse domain Is Nothing Then
                    Throw New System.Exception("Cookie name/domain required.")
                End If
                Dim ck As New System.Net.Cookie(name, value, If(path, "/"), domain) With {
                .Secure = c.Value(Of System.Nullable(Of System.Boolean))("secure").GetValueOrDefault(False),
                .HttpOnly = c.Value(Of System.Nullable(Of System.Boolean))("httpOnly").GetValueOrDefault(False)
            }
                _cookieContainer.Add(ck)
            Next
            Return New With {.count = arr.Count}
        End Function

        Private Function IsTransientStatus(code As Integer) As Boolean
            Select Case code
                Case 408, 425, 429, 500, 502, 503, 504
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Sub DebugDumpRequest(req As System.Net.Http.HttpRequestMessage)
            If Not GetDebugFlag("debug") AndAlso Not GetDebugFlag("debug_allAttempts") Then Return
            Try
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("[open_url][request] " & req.Method.Method & " " & req.RequestUri.ToString())
                For Each h In req.Headers
                    sb.AppendLine("  H: " & h.Key & ": " & String.Join(",", h.Value))
                Next
                If req.Content IsNot Nothing Then
                    For Each h In req.Content.Headers
                        sb.AppendLine("  HC: " & h.Key & ": " & String.Join(",", h.Value))
                    Next
                End If
                WriteDebug(sb.ToString())
            Catch
            End Try
        End Sub


        Private Async Function CmdOpenUrlAsync(parms As Newtonsoft.Json.Linq.JObject,
                                           timeoutMs As System.Int32,
                                           cancel As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)

            If parms Is Nothing Then Throw New System.Exception("open_url: params missing.")

            Dim rawUrl = parms.Value(Of System.String)("url")
            If String.IsNullOrWhiteSpace(rawUrl) Then Throw New System.Exception("open_url: 'url' missing.")

            Dim expandedRaw = ExpandTemplates(rawUrl)
            If String.IsNullOrWhiteSpace(expandedRaw) Then
                Throw New System.Exception($"open_url: URL expansion produced empty string (original='{rawUrl}').")
            End If
            If expandedRaw.Contains("{{") Then
                Throw New System.Exception($"open_url: URL contains unresolved template tokens after expansion: '{expandedRaw}'")
            End If

            Dim method = parms.Value(Of String)("method")
            If String.IsNullOrWhiteSpace(method) Then method = "GET"

            Dim stepTimeout = parms.Value(Of Integer?)("timeout_ms")
            If stepTimeout.HasValue AndAlso stepTimeout.Value > 0 Then
                timeoutMs = stepTimeout.Value
            ElseIf timeoutMs <= 0 Then
                timeoutMs = _defaultTimeoutMs
            End If

            Dim fullUrl = ResolveUrl(expandedRaw)
            If String.IsNullOrWhiteSpace(fullUrl) Then
                Throw New System.Exception($"open_url: ResolveUrl failed for '{expandedRaw}'.")
            End If
            fullUrl = SanitizePotentialMarkdownUrl(fullUrl)

            Dim uri As System.Uri = Nothing
            If Not System.Uri.TryCreate(fullUrl, System.UriKind.Absolute, uri) Then
                Throw New System.Exception("open_url: Not an absolute URL: " & fullUrl)
            End If

            Dim returnBody = parms.Value(Of Boolean?)("return_body").GetValueOrDefault(False)

            Dim hdrObj = TryCast(parms("headers"), Newtonsoft.Json.Linq.JObject)
            Dim bodyToken = parms("body")
            Dim bodyType = parms.Value(Of String)("body_type")

            Dim retryObj = TryCast(parms("retry"), Newtonsoft.Json.Linq.JObject)
            Dim maxRetry = If(retryObj IsNot Nothing, retryObj.Value(Of Integer?)("max").GetValueOrDefault(0), 0)
            Dim retryDelay = If(retryObj IsNot Nothing, retryObj.Value(Of Integer?)("delay_ms").GetValueOrDefault(1200), 1200)
            Dim backoff = If(retryObj IsNot Nothing, retryObj.Value(Of Double?)("backoff").GetValueOrDefault(2.0R), 2.0R)

            Dim attempt As Integer = 0
            Dim lastEx As Exception = Nothing
            Dim lastStatus As Integer = 0
            Dim usedHeadFallback As Boolean = False

            Dim swTotal = System.Diagnostics.Stopwatch.StartNew()

            Do
                cancel.ThrowIfCancellationRequested()
                lastEx = Nothing
                lastStatus = 0
                Dim plannedDelay As Integer = 0
                Dim success As Boolean = False
                Dim doHeadFallback As Boolean = False

                Using cts As New System.Threading.CancellationTokenSource(timeoutMs)
                    Using cancel.Register(Sub() cts.Cancel())
                        Try
                            Dim req = New System.Net.Http.HttpRequestMessage(New System.Net.Http.HttpMethod(method), uri)

                            For Each kv In _headersDefault
                                If Not req.Headers.Contains(kv.Key) Then
                                    req.Headers.TryAddWithoutValidation(kv.Key, kv.Value)
                                End If
                            Next
                            If hdrObj IsNot Nothing Then
                                For Each p In hdrObj.Properties()
                                    Dim hv = ExpandTemplates(p.Value.ToString())
                                    req.Headers.Remove(p.Name)
                                    req.Headers.TryAddWithoutValidation(p.Name, hv)
                                Next
                            End If

                            If bodyToken IsNot Nothing Then
                                Dim bodyStr = ExpandTemplates(bodyToken.ToString())
                                Select Case bodyType?.ToLowerInvariant()
                                    Case "json"
                                        req.Content = New System.Net.Http.StringContent(bodyStr, System.Text.Encoding.UTF8, "application/json")
                                    Case "form"
                                        Dim formPairs As New List(Of KeyValuePair(Of String, String))
                                        Dim jo = TryCast(bodyToken, Newtonsoft.Json.Linq.JObject)
                                        If jo IsNot Nothing Then
                                            For Each prop In jo.Properties()
                                                formPairs.Add(New KeyValuePair(Of String, String)(prop.Name, ExpandTemplates(prop.Value.ToString())))
                                            Next
                                        End If
                                        req.Content = New System.Net.Http.FormUrlEncodedContent(formPairs)
                                    Case Else
                                        req.Content = New System.Net.Http.StringContent(bodyStr, System.Text.Encoding.UTF8)
                                End Select
                            End If

                            DebugDumpRequest(req)

                            Dim swReq = System.Diagnostics.Stopwatch.StartNew()
                            Dim resp = Await _http.SendAsync(req, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, cts.Token)
                            swReq.Stop()

                            lastStatus = CInt(resp.StatusCode)
                            _vars("last_http_status") = lastStatus
                            _vars("last_http_elapsed_ms") = swReq.ElapsedMilliseconds

                            Dim bytes = Await resp.Content.ReadAsByteArrayAsync()
                            Dim bodyText = DecodeBody(bytes, resp.Content.Headers.ContentType)
                            _lastResponseUrl = resp.RequestMessage.RequestUri.ToString()
                            _lastResponseBody = bodyText
                            'LoadDocument(bodyText)
                            Await LoadDocumentAsync(bodyText, _lastResponseUrl, cancel).ConfigureAwait(False)

                            AutoExtractLinks()

                            If bodyText IsNot Nothing AndAlso bodyText.Length < 300 AndAlso
                           (bodyText.IndexOf("Incapsula", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                            bodyText.IndexOf("Access Denied", StringComparison.OrdinalIgnoreCase) >= 0) Then
                                Log("[open_url][warn] Possible WAF placeholder page detected (very short body).")
                            End If

                            If Not resp.IsSuccessStatusCode AndAlso IsTransientStatus(lastStatus) AndAlso attempt < maxRetry Then
                                plannedDelay = CInt(retryDelay * Math.Pow(backoff, attempt))
                                Log($"[open_url] Transient HTTP {lastStatus} → retry in {plannedDelay} ms (attempt {attempt + 1}/{maxRetry + 1}).")
                            Else
                                success = resp.IsSuccessStatusCode OrElse Not _failHard
                            End If

                            If success Then
                                swTotal.Stop()
                                If returnBody Then
                                    Return New With {.status = lastStatus, .url = _lastResponseUrl, .body = _lastResponseBody, .elapsed_ms = swReq.ElapsedMilliseconds}
                                Else
                                    Return New With {.status = lastStatus, .url = _lastResponseUrl, .elapsed_ms = swReq.ElapsedMilliseconds}
                                End If
                            End If

                        Catch tex As TaskCanceledException
                            lastEx = tex

                            Dim externalCancel = cancel.IsCancellationRequested
                            ' If outer token not cancelled we treat this as an internal timeout (cts or HTTP pipeline)
                            If Not externalCancel Then
                                Log($"[open_url] Timeout after {timeoutMs} ms (attempt {attempt + 1}/{maxRetry + 1}).")
                                If method.Equals("GET", StringComparison.OrdinalIgnoreCase) AndAlso Not usedHeadFallback AndAlso attempt = maxRetry Then
                                    usedHeadFallback = True
                                    doHeadFallback = True
                                End If
                                If attempt < maxRetry Then
                                    plannedDelay = CInt(retryDelay * Math.Pow(backoff, attempt))
                                    Log($"[open_url] Will retry after timeout in {plannedDelay} ms.")
                                ElseIf _failHard Then
                                    Throw
                                End If
                            Else
                                Log("[open_url] External cancellation signal received – aborting.")
                                Throw
                            End If

                        Catch ex As Exception
                            lastEx = ex
                            Log("[open_url] Error: " & ex.GetType().Name & ": " & ex.Message)
                            If attempt < maxRetry Then
                                plannedDelay = CInt(retryDelay * Math.Pow(backoff, attempt))
                                Log($"[open_url] Retry in {plannedDelay} ms (attempt {attempt + 1}/{maxRetry + 1}).")
                            ElseIf _failHard Then
                                Throw
                            End If
                        End Try

                        ' Perform deferred HEAD fallback (allowed to Await here, outside Catch)
                        If doHeadFallback Then
                            Try
                                Log("[open_url] Trying HEAD fallback to test reachability …")
                                Using headCts As New System.Threading.CancellationTokenSource(Math.Min(5000, timeoutMs))
                                    Dim headReq = New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Head, uri)
                                    Dim respHead = Await _http.SendAsync(headReq, headCts.Token)
                                    Log("[open_url] HEAD fallback status: " & CInt(respHead.StatusCode))
                                End Using
                            Catch ex2 As Exception
                                Log("[open_url] HEAD fallback failed: " & ex2.Message)
                            End Try
                        End If
                    End Using
                End Using

                attempt += 1
                If attempt <= maxRetry AndAlso plannedDelay > 0 Then
                    Await System.Threading.Tasks.Task.Delay(plannedDelay, cancel)
                End If
            Loop While attempt <= maxRetry

            swTotal.Stop()
            _vars("last_http_status") = lastStatus
            _vars("last_http_elapsed_ms") = swTotal.ElapsedMilliseconds

            Dim detail = If(lastEx IsNot Nothing, lastEx.Message, "Unknown failure")
            Throw New System.Exception($"open_url failed after {attempt} attempt(s). Status={lastStatus}. Detail={detail}")
        End Function

        Private Sub AutoExtractLinks()
            Try
                Dim enabled As Boolean = True
                If _vars.ContainsKey("auto_link_enable") Then
                    Dim o = _vars("auto_link_enable")
                    If TypeOf o Is Boolean Then enabled = DirectCast(o, Boolean)
                End If
                If Not enabled Then Exit Sub

                If _lastDoc Is Nothing OrElse String.IsNullOrEmpty(_lastResponseBody) Then Exit Sub

                Dim patterns As New List(Of String)
                If _vars.ContainsKey("auto_link_patterns") Then
                    Dim raw = _vars("auto_link_patterns")
                    If TypeOf raw Is IEnumerable(Of String) Then
                        patterns.AddRange(DirectCast(raw, IEnumerable(Of String)))
                    ElseIf TypeOf raw Is String Then
                        patterns.Add(DirectCast(raw, String))
                    End If
                End If
                ' Provide a sensible default if none specified
                If patterns.Count = 0 Then
                    patterns.Add("(?i)\b(show|detail|decision|case|id=|doc|file|download)\b")
                End If

                Dim minLen As Integer = 15
                If _vars.ContainsKey("auto_link_min") Then
                    Dim mObj = _vars("auto_link_min")
                    If TypeOf mObj Is Integer Then minLen = DirectCast(mObj, Integer)
                End If

                Dim anchors = _lastDoc.DocumentNode.SelectNodes("//a[@href]")
                Dim collected As New List(Of String)
                If anchors IsNot Nothing Then
                    For Each a In anchors
                        Dim href = a.GetAttributeValue("href", "")
                        If String.IsNullOrWhiteSpace(href) Then Continue For
                        If href.Length < minLen Then Continue For

                        Dim abs = ResolveUrl(href)
                        If String.IsNullOrEmpty(abs) Then Continue For

                        Dim score = 0
                        For Each p In patterns
                            If Regex.IsMatch(href, p) OrElse Regex.IsMatch(abs, p) Then
                                score += 1
                            End If
                        Next
                        ' Keep if matched at least one pattern OR caller allows a passive collection via flag
                        If score > 0 Then
                            If Not collected.Contains(abs) Then collected.Add(abs)
                        End If
                    Next
                End If

                _vars("auto_links") = collected
                Log($"[auto_extract_generic] Collected {collected.Count} link(s).")
            Catch ex As Exception
                Log("[auto_extract_generic][error] " & ex.Message)
            End Try
        End Sub

        Public Sub AddAutoLinkPattern(pattern As String)
            If String.IsNullOrWhiteSpace(pattern) Then Exit Sub
            If Not _vars.ContainsKey("auto_link_patterns") Then
                _vars("auto_link_patterns") = New List(Of String)
            End If
            Dim lst = TryCast(_vars("auto_link_patterns"), List(Of String))
            If lst Is Nothing Then
                lst = New List(Of String)
                _vars("auto_link_patterns") = lst
            End If
            If Not lst.Contains(pattern) Then lst.Add(pattern)
        End Sub

        Private Async Function CmdWaitAsync(parms As Newtonsoft.Json.Linq.JObject, cancel As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim ms = parms.Value(Of System.Nullable(Of System.Int32))("ms").GetValueOrDefault(0)
            If ms > 0 Then Await System.Threading.Tasks.Task.Delay(ms, cancel)
            Return New With {.slept = ms}
        End Function

        Private Function CmdFind(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            ' Expected params: { "in": "<varName>", "text": "needle", "assign": {"var":"found"} }
            If parms Is Nothing Then Return Nothing

            Dim varName = parms.Value(Of String)("in")
            Dim needle = parms.Value(Of String)("text")
            If String.IsNullOrEmpty(varName) OrElse String.IsNullOrEmpty(needle) Then
                Return New With {.count = 0}
            End If

            Dim hayObj As Object = Nothing
            If _vars.TryGetValue(varName, hayObj) Then
                ' Normalize to string
                Dim hay As String
                If TypeOf hayObj Is String Then
                    hay = DirectCast(hayObj, String)
                ElseIf TypeOf hayObj Is Newtonsoft.Json.Linq.JToken Then
                    hay = DirectCast(hayObj, Newtonsoft.Json.Linq.JToken).ToString()
                ElseIf TypeOf hayObj Is Newtonsoft.Json.Linq.JObject Then
                    hay = DirectCast(hayObj, Newtonsoft.Json.Linq.JObject).ToString(Newtonsoft.Json.Formatting.None)
                ElseIf TypeOf hayObj Is Newtonsoft.Json.Linq.JArray Then
                    hay = DirectCast(hayObj, Newtonsoft.Json.Linq.JArray).ToString(Newtonsoft.Json.Formatting.None)
                Else
                    hay = hayObj.ToString()
                End If

                Dim idx = hay.IndexOf(needle, StringComparison.OrdinalIgnoreCase)
                Dim found = idx >= 0
                Dim assignObj = TryCast(parms("assign"), Newtonsoft.Json.Linq.JObject)
                If assignObj IsNot Nothing Then
                    Dim targetVar = assignObj.Value(Of String)("var")
                    If Not String.IsNullOrEmpty(targetVar) Then
                        _vars(targetVar) = found
                    End If
                End If
                Return New With {.found = found, .index = idx}
            End If

            Return New With {.found = False, .index = -1}
        End Function
        Private Function CmdExtractText(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim sel = TryCast(parms("selector"), Newtonsoft.Json.Linq.JObject)
            Dim all = parms.Value(Of System.Nullable(Of System.Boolean))("all").GetValueOrDefault(False)
            Dim normalize = parms.Value(Of System.Nullable(Of System.Boolean))("normalize_whitespace").GetValueOrDefault(True)
            Dim pattern = parms.Value(Of System.String)("regex")
            Dim groupIdx = parms.Value(Of System.Nullable(Of System.Int32))("group").GetValueOrDefault(0)
            Dim nodes = ResolveSelector(sel)
            If nodes Is Nothing Then Return If(all, New System.Collections.Generic.List(Of System.String)(), Nothing)
            Dim texts As New System.Collections.Generic.List(Of System.String)
            For Each n In nodes
                Dim t = GetInnerText(n, normalize)
                If Not System.String.IsNullOrWhiteSpace(pattern) Then
                    Dim m = System.Text.RegularExpressions.Regex.Match(t, pattern, System.Text.RegularExpressions.RegexOptions.Singleline)
                    If m.Success AndAlso groupIdx <= m.Groups.Count - 1 Then
                        t = m.Groups(groupIdx).Value
                    Else
                        t = System.String.Empty
                    End If
                End If
                If Not all Then Return t
                texts.Add(t)
            Next
            Return texts
        End Function

        Private Function CmdExtractHtml(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim sel = TryCast(parms("selector"), Newtonsoft.Json.Linq.JObject)
            Dim outer = parms.Value(Of System.Nullable(Of System.Boolean))("outer").GetValueOrDefault(False)
            Dim nodes = ResolveSelector(sel)
            If nodes Is Nothing OrElse nodes.Count = 0 Then Return System.String.Empty
            Return If(outer, nodes(0).OuterHtml, nodes(0).InnerHtml)
        End Function

        Private Function CmdExtractAttribute(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim nodesVar = parms.Value(Of String)("nodes_var")
            Dim attrName = parms.Value(Of String)("attribute")
            Dim targetVar = parms.Value(Of String)("var")
            If String.IsNullOrEmpty(nodesVar) OrElse String.IsNullOrEmpty(attrName) OrElse String.IsNullOrEmpty(targetVar) Then
                Throw New Exception("extract_attribute: missing parameters")
            End If

            If Not _vars.ContainsKey(nodesVar) Then
                _vars(targetVar) = New List(Of String)
                Return Nothing
            End If

            Dim raw = _vars(nodesVar)
            Dim nodeList As New List(Of Object)

            ' Normalize raw into a list of node objects
            If TypeOf raw Is IEnumerable(Of Object) Then
                nodeList.AddRange(DirectCast(raw, IEnumerable(Of Object)))
            ElseIf TypeOf raw Is IEnumerable Then
                For Each it In DirectCast(raw, IEnumerable)
                    nodeList.Add(it)
                Next
            Else
                nodeList.Add(raw) ' single object fallback
            End If

            Dim hrefs As New List(Of String)
            For Each n In nodeList
                Try
                    Dim jo = Newtonsoft.Json.Linq.JObject.FromObject(n)
                    Dim attrs = jo("attributes")
                    If attrs IsNot Nothing AndAlso attrs(attrName) IsNot Nothing Then
                        Dim v = attrs.Value(Of String)(attrName)
                        If Not String.IsNullOrWhiteSpace(v) Then hrefs.Add(v)
                    End If
                Catch
                    ' ignore serialization issues
                End Try
            Next

            _vars(targetVar) = hrefs
            Return Nothing
        End Function

        Private Async Function CmdDownloadUrlAsync(parms As Newtonsoft.Json.Linq.JObject, cancel As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim url = ResolveUrl(ExpandTemplates(parms.Value(Of System.String)("url")))
            Dim targetDir = ExpandTemplates(parms.Value(Of System.String)("target_dir"))
            Dim filename = ExpandTemplates(parms.Value(Of System.String)("filename"))
            If System.String.IsNullOrWhiteSpace(targetDir) Then Throw New System.Exception("target_dir missing.")
            If System.String.IsNullOrWhiteSpace(filename) Then filename = "download.bin"
            If Not System.IO.Directory.Exists(targetDir) Then System.IO.Directory.CreateDirectory(targetDir)
            Dim path = System.IO.Path.Combine(targetDir, filename)

            Dim method = parms.Value(Of System.String)("method")
            If System.String.IsNullOrWhiteSpace(method) Then method = "GET"

            Dim headers = TryCast(parms("headers"), Newtonsoft.Json.Linq.JObject)
            Dim req As New System.Net.Http.HttpRequestMessage(New System.Net.Http.HttpMethod(method), url)
            If headers IsNot Nothing Then
                For Each h In headers.Properties()
                    req.Headers.TryAddWithoutValidation(h.Name, ExpandTemplates(h.Value.ToString()))
                Next
            End If

            Dim body = parms("body")
            Dim bodyType = parms.Value(Of System.String)("body_type")
            If body IsNot Nothing Then
                Select Case bodyType
                    Case "json"
                        req.Content = New System.Net.Http.StringContent(ExpandTemplates(body.ToString()), System.Text.Encoding.UTF8, "application/json")
                    Case "form"
                        Dim kv As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of System.String, System.String))
                        For Each p In TryCast(body, Newtonsoft.Json.Linq.JObject).Properties()
                            kv.Add(New System.Collections.Generic.KeyValuePair(Of System.String, System.String)(p.Name, ExpandTemplates(p.Value.ToString())))
                        Next
                        req.Content = New System.Net.Http.FormUrlEncodedContent(kv)
                    Case Else
                        req.Content = New System.Net.Http.StringContent(ExpandTemplates(body.ToString()), System.Text.Encoding.UTF8)
                End Select
            End If

            Dim resp = Await _http.SendAsync(req, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, cancel)
            Using fs As New System.IO.FileStream(path, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None)
                Await resp.Content.CopyToAsync(fs)
            End Using
            Return New With {.path = path, .status = CInt(resp.StatusCode)}
        End Function

        Private Function CmdSaveFile(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim content = ExpandTemplates(parms.Value(Of System.String)("content"))
            Dim filePath = ExpandTemplates(parms.Value(Of System.String)("path"))
            Dim encoding = parms.Value(Of System.String)("encoding")
            If System.String.IsNullOrWhiteSpace(filePath) Then Throw New System.Exception("path missing.")
            Dim dir = System.IO.Path.GetDirectoryName(filePath)
            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
            If System.String.Equals(encoding, "binary", System.StringComparison.OrdinalIgnoreCase) Then
                Dim bytes = System.Convert.FromBase64String(content)
                System.IO.File.WriteAllBytes(filePath, bytes)
            Else
                System.IO.File.WriteAllText(filePath, If(content, System.String.Empty), System.Text.Encoding.UTF8)
            End If
            Return New With {.path = filePath}
        End Function

        Private Function CmdReadFile(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim filePath = ExpandTemplates(parms.Value(Of System.String)("path"))
            Dim encoding = parms.Value(Of System.String)("encoding")
            If System.String.IsNullOrWhiteSpace(filePath) Then Throw New System.Exception("path missing.")
            Dim txt As System.String
            If System.String.Equals(encoding, "binary", System.StringComparison.OrdinalIgnoreCase) Then
                Dim bytes = System.IO.File.ReadAllBytes(filePath)
                txt = System.Convert.ToBase64String(bytes)
            Else
                txt = System.IO.File.ReadAllText(filePath, System.Text.Encoding.UTF8)
            End If
            Return txt
        End Function

        Private Async Function CmdHttpRequestAsync(parms As Newtonsoft.Json.Linq.JObject, timeoutMs As System.Int32, cancel As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim url = ResolveUrl(ExpandTemplates(parms.Value(Of System.String)("url")))
            Dim absolute As System.Uri
            If Not System.Uri.TryCreate(url, System.UriKind.Absolute, absolute) Then
                Throw New System.Exception("URL must be absolute (include scheme). Got: " & url)
            End If
            Dim method = parms.Value(Of System.String)("method")
            If System.String.IsNullOrWhiteSpace(method) Then method = "GET"
            Dim req As New System.Net.Http.HttpRequestMessage(New System.Net.Http.HttpMethod(method), url)

            Dim headers = TryCast(parms("headers"), Newtonsoft.Json.Linq.JObject)
            If headers IsNot Nothing Then
                For Each h In headers.Properties()
                    req.Headers.TryAddWithoutValidation(h.Name, ExpandTemplates(h.Value.ToString()))
                Next
            End If

            Dim query = TryCast(parms("query"), Newtonsoft.Json.Linq.JObject)
            If query IsNot Nothing Then
                Dim ub As New System.UriBuilder(url)
                Dim qParts As New System.Collections.Generic.List(Of System.String)
                For Each p In query.Properties()
                    qParts.Add(System.Uri.EscapeDataString(p.Name) & "=" & System.Uri.EscapeDataString(ExpandTemplates(p.Value.ToString())))
                Next
                If qParts.Count > 0 Then
                    ub.Query = System.String.Join("&", qParts)
                    req.RequestUri = ub.Uri
                End If
            End If

            Dim body = parms("body")
            Dim bodyType = parms.Value(Of System.String)("body_type")
            If body IsNot Nothing Then
                Select Case bodyType
                    Case "json"
                        req.Content = New System.Net.Http.StringContent(ExpandTemplates(body.ToString()), System.Text.Encoding.UTF8, "application/json")
                    Case "form"
                        Dim kv As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of System.String, System.String))
                        For Each p In TryCast(body, Newtonsoft.Json.Linq.JObject).Properties()
                            kv.Add(New System.Collections.Generic.KeyValuePair(Of System.String, System.String)(p.Name, ExpandTemplates(p.Value.ToString())))
                        Next
                        req.Content = New System.Net.Http.FormUrlEncodedContent(kv)
                    Case Else
                        req.Content = New System.Net.Http.StringContent(ExpandTemplates(body.ToString()), System.Text.Encoding.UTF8)
                End Select
            End If

            Using cts As New System.Threading.CancellationTokenSource(timeoutMs)
                Using cancel.Register(Sub() cts.Cancel())
                    Dim resp = Await _http.SendAsync(req, cts.Token)
                    Dim bytes = Await resp.Content.ReadAsByteArrayAsync()
                    Dim bodyText = DecodeBody(bytes, resp.Content.Headers.ContentType)
                    _lastResponseUrl = resp.RequestMessage.RequestUri.ToString()
                    _lastResponseBody = bodyText
                    'LoadDocument(bodyText)
                    Await LoadDocumentAsync(bodyText, _lastResponseUrl, cancel).ConfigureAwait(False)
                    Return New With {
                    .status = CInt(resp.StatusCode),
                    .headers = resp.Headers.ToString(),
                    .body = bodyText,
                    .url = _lastResponseUrl
                }
                End Using
            End Using
        End Function

        Private Function CmdSetVar(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim name = parms.Value(Of System.String)("name")
            Dim valueToken = parms("value")
            If System.String.IsNullOrWhiteSpace(name) Then Throw New System.Exception("variable name missing.")
            Dim valueObj As System.Object
            If valueToken Is Nothing Then
                valueObj = Nothing
            ElseIf valueToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                valueObj = ExpandTemplates(valueToken.ToString())
            Else
                valueObj = valueToken.ToObject(Of System.Object)()
            End If
            SafeStoreVar(name, valueObj)
            Return New With {.name = name, .value = valueObj}
        End Function

        Private Function CmdTemplate(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim tpl = parms.Value(Of System.String)("template")
            Dim ctxToken = parms("context")
            Dim ctxObj As System.Object = Nothing
            If ctxToken IsNot Nothing Then
                Dim mt = TryCast(ctxToken, JObject)
                If mt IsNot Nothing Then
                    For Each p In mt.Properties().ToList()
                        If p.Value.Type = JTokenType.String Then
                            Dim raw = p.Value.ToString()
                            Dim expanded = ExpandTemplates(raw)
                            Dim parsed As JToken = Nothing
                            Try
                                parsed = JToken.Parse(expanded)
                            Catch
                            End Try
                            If parsed IsNot Nothing Then
                                p.Value = parsed
                            Else
                                p.Value = expanded
                            End If
                        End If
                    Next
                End If
                ' Kontext an Renderer weitergeben
                ctxObj = ctxToken
            End If

            ' Mustache render (nur Sections & bekannte Variablen aus context)
            Dim rendered = SimpleMustacheRender(tpl, ctxObj)

            ' Zweiter Pass: globale Variablen (_vars) expandieren
            rendered = ExpandTemplates(rendered)

            Return rendered
        End Function

        Private Function CmdDeleteFile(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            If parms Is Nothing Then Return False
            Dim rawPath = parms.Value(Of String)("path")
            If String.IsNullOrWhiteSpace(rawPath) Then Return False
            Dim expanded = ExpandTemplates(rawPath)
            Try
                If System.IO.File.Exists(expanded) Then
                    System.IO.File.Delete(expanded)
                    Log($"[delete_file] Deleted {expanded}")
                    Return True
                Else
                    Log($"[delete_file] Not found: {expanded}")
                    Return False
                End If
            Catch ex As Exception
                Log($"[delete_file] Error deleting {expanded}: {ex.Message}")
                Return False
            End Try
        End Function


        Private Function CmdSendEmailReport(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            If parms Is Nothing Then Return False
            Try
                Dim ExpandSafe As Func(Of String, String) =
                Function(s As String)
                    If String.IsNullOrWhiteSpace(s) Then Return s
                    Try
                        Return ExpandTemplates(s)
                    Catch
                        Return s
                    End Try
                End Function

                ' Raw parameters
                Dim rawTo = parms.Value(Of String)("to")
                Dim rawSubject = parms.Value(Of String)("subject")
                Dim rawSmtpHost = parms.Value(Of String)("smtp_host")
                Dim rawPort = parms.Value(Of String)("smtp_port")
                Dim rawFromEmail = parms.Value(Of String)("from_email")
                Dim rawFromName = parms.Value(Of String)("from_name")
                Dim rawBodyMd = parms.Value(Of String)("body_markdown")
                Dim rawIpOverride = If(parms.Value(Of String)("ip_override"), parms.Value(Of String)("ip"))
                Dim rawHeloDomain = parms.Value(Of String)("helo_domain")
                Dim rawNet = parms.Value(Of String)("net")

                ' Expand templates
                Dim toAddr = ExpandSafe(rawTo)
                Dim subject = ExpandSafe(rawSubject)
                Dim smtpHost = ExpandSafe(rawSmtpHost)
                Dim fromEmail = ExpandSafe(rawFromEmail)
                Dim fromName = ExpandSafe(rawFromName)
                Dim bodyMarkdownTemplate = ExpandSafe(rawBodyMd)
                Dim ipOverride = ExpandSafe(rawIpOverride)
                Dim heloDomain = ExpandSafe(rawHeloDomain)
                Dim netOverride = ExpandSafe(rawNet)

                Dim smtpPort As Integer = 25
                If Not String.IsNullOrWhiteSpace(rawPort) Then
                    Integer.TryParse(ExpandSafe(rawPort), smtpPort)
                End If

                Dim smtpSsl = String.Equals(parms.Value(Of String)("smtp_ssl"), "true", StringComparison.OrdinalIgnoreCase)
                Dim smtpAuth = String.Equals(parms.Value(Of String)("smtp_auth"), "true", StringComparison.OrdinalIgnoreCase)
                Dim smtpUser = ExpandSafe(parms.Value(Of String)("smtp_user"))
                Dim smtpPass = ExpandSafe(parms.Value(Of String)("smtp_pass"))

                If String.IsNullOrWhiteSpace(toAddr) OrElse String.IsNullOrWhiteSpace(smtpHost) Then
                    Log("[send_email_report] Missing required fields (to / smtp_host).")
                    Return False
                End If

                If String.IsNullOrWhiteSpace(fromEmail) Then fromEmail = SharedMethods.WebAgent_DefaultFromEMail
                If String.IsNullOrWhiteSpace(fromName) Then fromName = SharedMethods.WebAgent_DefaultFromName
                If String.IsNullOrWhiteSpace(subject) Then subject = "Report"

                ' Prepare plain source text (after template expansion)
                Dim plainBody As String = If(String.IsNullOrWhiteSpace(bodyMarkdownTemplate), "(leer)", bodyMarkdownTemplate)

                ' Decide if we should treat input as Markdown or already HTML
                Dim looksLikeHtml = plainBody.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                plainBody.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                plainBody.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0

                Dim htmlBodyCore As String
                If looksLikeHtml Then
                    ' Use as-is (still append footer later)
                    htmlBodyCore = plainBody
                Else
                    ' Convert Markdown -> HTML
                    Try
                        Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                        htmlBodyCore = Markdig.Markdown.ToHtml(plainBody, pipeline)
                    Catch
                        htmlBodyCore = System.Net.WebUtility.HtmlEncode(plainBody).Replace(vbCrLf, "<br/>")
                    End Try
                End If

                ' Footer
                Dim ip = If(Not String.IsNullOrWhiteSpace(ipOverride), ipOverride, GetFirstLocalIPv4())
                Dim netDomain = If(Not String.IsNullOrWhiteSpace(netOverride), netOverride, GetCurrentNetworkDomain())
                Dim agentName = SharedMethods.AN
                Dim footerHtml = $"<p style=""font-size:11px;color:#666;margin-top:16px"">(created using {agentName} WebAgent at {ip} from {netDomain})</p>"
                Dim footerPlain = Environment.NewLine & "(created using " & agentName & " WebAgent at " & ip & " from " & netDomain & ")"

                ' If input already had full HTML document, inject footer before closing body/html if possible
                Dim fullHtml As String
                If looksLikeHtml Then
                    Dim injected = False
                    Dim sb = New System.Text.StringBuilder(htmlBodyCore)
                    ' Try typical insertion points
                    Dim bodyCloseIdx = sb.ToString().IndexOf("</body>", StringComparison.OrdinalIgnoreCase)
                    If bodyCloseIdx >= 0 Then
                        sb.Insert(bodyCloseIdx, footerHtml)
                        injected = True
                    End If
                    If Not injected Then
                        Dim htmlCloseIdx = sb.ToString().IndexOf("</html>", StringComparison.OrdinalIgnoreCase)
                        If htmlCloseIdx >= 0 Then
                            sb.Insert(htmlCloseIdx, footerHtml)
                            injected = True
                        End If
                    End If
                    If Not injected Then
                        sb.Append(footerHtml)
                    End If
                    fullHtml = sb.ToString()
                Else
                    ' Build minimal standards-compliant HTML doc
                    fullHtml =
                            $"<!DOCTYPE html>
                                <html>
                                <head>
                                <meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />
                                <title>{System.Net.WebUtility.HtmlEncode(subject)}</title>
                                </head>
                                <body>
                                {htmlBodyCore}
                                {footerHtml}
                                </body>
                                </html>"
                    plainBody &= footerPlain
                End If

                Dim msg = New System.Net.Mail.MailMessage()
                msg.From = New System.Net.Mail.MailAddress(fromEmail, fromName, System.Text.Encoding.UTF8)
                For Each addr In toAddr.Split({";", ","}, StringSplitOptions.RemoveEmptyEntries)
                    msg.To.Add(addr.Trim())
                Next
                msg.Subject = subject
                msg.SubjectEncoding = Encoding.UTF8
                msg.BodyEncoding = Encoding.UTF8

                ' Build proper multipart/alternative
                Dim plainView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(plainBody, Encoding.UTF8, "text/plain")
                plainView.TransferEncoding = System.Net.Mime.TransferEncoding.QuotedPrintable

                Dim htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(fullHtml, Encoding.UTF8, "text/html")
                htmlView.TransferEncoding = System.Net.Mime.TransferEncoding.QuotedPrintable

                ' Order: plain first, then html
                msg.AlternateViews.Add(plainView)
                msg.AlternateViews.Add(htmlView)

                ' Do NOT also set msg.Body / IsBodyHtml when using AlternateViews (avoids some relays forcing text/plain)
                ' SMTP client
                Dim client = New System.Net.Mail.SmtpClient(smtpHost, smtpPort) With {
                .EnableSsl = smtpSsl,
                .DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                .Timeout = 20000
            }

                If smtpAuth AndAlso Not String.IsNullOrWhiteSpace(smtpUser) Then
                    client.Credentials = New System.Net.NetworkCredential(smtpUser, smtpPass)
                Else
                    client.UseDefaultCredentials = False
                End If

                ' HELO / EHLO domain generation
                If String.IsNullOrWhiteSpace(heloDomain) Then
                    Dim hostLocal = System.Net.Dns.GetHostName()
                    If Not hostLocal.Contains(".") AndAlso Not String.IsNullOrWhiteSpace(netDomain) AndAlso netDomain.Contains(".") Then
                        heloDomain = hostLocal & "." & netDomain
                    Else
                        heloDomain = hostLocal
                    End If
                End If
                If Not heloDomain.Contains(".") Then heloDomain &= ".localdomain"

                ' Try overriding internal clientDomain field (best effort)
                Try
                    Dim f = GetType(System.Net.Mail.SmtpClient).GetField("clientDomain", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
                    If f IsNot Nothing Then
                        f.SetValue(client, heloDomain)
                        Log($"[send_email_report] Using HELO domain: {heloDomain}")
                    End If
                Catch ex As Exception
                    Log($"[send_email_report] HELO override failed: {ex.Message}")
                End Try

                client.Send(msg)
                Log("[send_email_report] Mail sent (multipart/alternative).")
                Return True

            Catch ex As System.Net.Mail.SmtpException
                Log("[send_email_report] SMTP error: " & ex.Message)
                Return False
            Catch ex As Exception
                Log("[send_email_report] General error: " & ex.Message)
                Return False
            End Try
        End Function


        Private Function GetSetting(parms As Newtonsoft.Json.Linq.JObject, key As System.String, defaultValue As System.String) As System.String
            ' 1) params
            If parms IsNot Nothing Then
                Dim tok = parms(key)
                If tok IsNot Nothing Then
                    Dim val As System.String = ExpandTemplates(tok.ToString())
                    If Not System.String.IsNullOrWhiteSpace(val) Then Return val
                End If
            End If
            ' 2) _vars
            Dim v As System.Object = Nothing
            If _vars IsNot Nothing AndAlso _vars.TryGetValue(key, v) AndAlso v IsNot Nothing Then
                Dim s As System.String = v.ToString()
                If Not System.String.IsNullOrWhiteSpace(s) Then Return s
            End If
            ' 3) _secrets
            Dim s2 As System.String = Nothing
            If _secrets IsNot Nothing AndAlso _secrets.TryGetValue(key, s2) AndAlso Not System.String.IsNullOrWhiteSpace(s2) Then
                Return s2
            End If
            Return defaultValue
        End Function

        Private Function GetFirstLocalIPv4() As System.String
            Try
                Dim host As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName())
                For Each ip As System.Net.IPAddress In host.AddressList
                    If ip IsNot Nothing AndAlso ip.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                        If Not System.Net.IPAddress.IsLoopback(ip) Then
                            Return ip.ToString()
                        End If
                    End If
                Next
            Catch ex As System.Exception
                Try : Log("[ip] " & ex.Message) : Catch : End Try
            End Try
            Return "0.0.0.0"
        End Function

        Private Function GetCurrentNetworkDomain() As System.String
            Try
                Dim ipgp As System.Net.NetworkInformation.IPGlobalProperties =
        System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties()
                If ipgp IsNot Nothing Then
                    Dim dnsDom As System.String = ipgp.DomainName
                    If Not System.String.IsNullOrWhiteSpace(dnsDom) Then
                        Return dnsDom
                    End If
                End If

                Dim envDom As System.String = System.Environment.UserDomainName
                If Not System.String.IsNullOrWhiteSpace(envDom) Then
                    Return envDom
                End If
            Catch ex As System.Exception
                Try : Log("[net] " & ex.Message) : Catch : End Try
            End Try
            Return "UNKNOWN"
        End Function



        Private Async Function CmdIfAsync(parms As JObject,
                              cancel As CancellationToken) As Task(Of Object)
            ' Expected params JSON:
            ' {
            '   "condition": "<expression understood by EvalCondition>",
            '   "steps": [ ... ],
            '   "else_steps": [ ... ]   (optional)
            ' }
            Dim conditionExpr As String = If(parms?.Value(Of String)("condition"), "")
            Dim thenSteps As JArray = parms?.Value(Of JArray)("steps")
            Dim elseSteps As JArray = parms?.Value(Of JArray)("else_steps")

            Dim condResult As Boolean = False
            If conditionExpr <> "" Then
                Try
                    condResult = EvalCondition(conditionExpr)
                Catch ex As Exception
                    WriteDebug("[if] Condition error: " & ex.Message)
                    condResult = False
                End Try
            Else
                WriteDebug("[if] Missing condition – treated as False.")
            End If

            Dim branch As JArray = If(condResult, thenSteps, elseSteps)
            If branch Is Nothing OrElse branch.Count = 0 Then
                WriteDebug("[if] No steps in selected branch.")
                Return Nothing
            End If

            Await RunSubStepsAsync(branch, cancel)
            Return Nothing
        End Function


        Private Function DeepCloneObject(obj As Object) As Object
            If obj Is Nothing Then Return Nothing
            Try
                Dim jt = JToken.FromObject(obj)
                Return jt.DeepClone().ToObject(Of Object)()
            Catch
                Return obj
            End Try
        End Function


        Private Sub SafeStoreVar(varName As String, value As Object)
            ' Wrapper now delegates to debug-aware variant
            SafeStoreVar_DebugPatch(varName, value)
        End Sub


        Private Async Function CmdForEachAsync(parms As JObject,
                                       cancel As CancellationToken) As Task(Of Object)
            If parms Is Nothing Then Return Nothing

            Dim listVar = parms.Value(Of String)("list")
            Dim itemVar = parms.Value(Of String)("item_var")
            Dim steps = TryCast(parms("steps"), JArray)
            If String.IsNullOrWhiteSpace(listVar) OrElse String.IsNullOrWhiteSpace(itemVar) OrElse steps Is Nothing Then
                Throw New Exception("foreach: missing list / item_var / steps")
            End If

            Dim continueOnError As Boolean = True
            Dim stopOnError As Boolean = False
            Dim maxItems = parms.Value(Of Integer?)("max_items")
            Dim softBreakVar = parms.Value(Of String)("break_if_var_true") ' optional runtime break flag

            If parms("continue_on_error") IsNot Nothing Then
                Boolean.TryParse(parms("continue_on_error").ToString(), continueOnError)
            End If
            If parms("stop_on_error") IsNot Nothing Then
                Boolean.TryParse(parms("stop_on_error").ToString(), stopOnError)
            End If
            ' stop_on_error overrides continue_on_error explicitly
            If stopOnError Then continueOnError = False

            Dim src As Object = Nothing
            If Not _vars.TryGetValue(listVar, src) OrElse src Is Nothing Then
                Log($"[foreach] list '{listVar}' not found or null → skipping.")
                Return New With {.count = 0, .executed = 0}
            End If

            Dim enumerable As IEnumerable(Of Object)
            If TypeOf src Is JArray Then
                enumerable = DirectCast(src, JArray).Select(Function(t) CType(t, JToken))
            ElseIf TypeOf src Is IEnumerable(Of Object) Then
                enumerable = DirectCast(src, IEnumerable(Of Object))
            ElseIf TypeOf src Is IEnumerable Then
                enumerable = DirectCast(src, IEnumerable).Cast(Of Object)()
            Else
                enumerable = {src}
            End If

            Dim idx As Integer = 0
            Dim executed As Integer = 0
            For Each item In enumerable
                cancel.ThrowIfCancellationRequested()

                If maxItems.HasValue AndAlso idx >= maxItems.Value Then
                    Log($"[foreach] max_items {maxItems.Value} reached.")
                    Exit For
                End If

                _vars(itemVar) = item
                _vars(itemVar & "_index") = idx

                Dim localToken = cancel ' (kept for clarity – could create linked CTS here if needed)

                Try
                    Await RunSubStepsAsync(steps, localToken)
                    executed += 1

                Catch ex As OperationCanceledException
                    Dim treatAsExternal = cancel.IsCancellationRequested
                    Dim tolerate = False
                    If _vars.ContainsKey("continue_on_llm_timeout") AndAlso _IsTruthy(_vars("continue_on_llm_timeout")) Then
                        tolerate = True
                    End If

                    If treatAsExternal AndAlso Not tolerate Then
                        Log("[foreach] External cancellation requested → abort loop.")
                        Throw
                    Else
                        Log($"[foreach] Internal or tolerated cancellation at index={idx}: {ex.Message}")
                        If stopOnError Then
                            Log("[foreach] stop_on_error=True → aborting.")
                            Exit For
                        End If
                        If Not continueOnError Then
                            Log("[foreach] continue_on_error=False → breaking.")
                            Exit For
                        End If
                        ' continue with next item
                    End If

                Catch ex As Exception
                    Log($"[foreach] Exception at index={idx}: {ex.GetType().Name} {ex.Message}")
                    If stopOnError Then
                        Log("[foreach] stop_on_error=True → aborting.")
                        Exit For
                    End If
                    If Not continueOnError Then
                        Log("[foreach] continue_on_error=False → breaking.")
                        Exit For
                    End If
                End Try

                ' Soft runtime break variable
                If Not String.IsNullOrWhiteSpace(softBreakVar) Then
                    If _vars.ContainsKey(softBreakVar) AndAlso _IsTruthy(_vars(softBreakVar)) Then
                        Log($"[foreach] break_if_var_true '{softBreakVar}' = true → exiting loop.")
                        Exit For
                    End If
                End If

                idx += 1
            Next

            Return New With {.count = idx, .executed = executed}
        End Function

        Private Async Function RunSubStepsAsync(steps As JArray,
                                            cancel As CancellationToken) As System.Threading.Tasks.Task
            For Each st In steps
                cancel.ThrowIfCancellationRequested()

                Dim stepObj = TryCast(st, JObject)
                If stepObj Is Nothing Then Continue For

                Dim sid = stepObj.Value(Of String)("id")
                Dim cmd = stepObj.Value(Of String)("command")

                Try
                    Dim timeoutMs = stepObj.Value(Of Integer?)("timeout_ms").GetValueOrDefault(_defaultTimeoutMs)
                    Dim retry = TryCast(stepObj("retry"), JObject)
                    Dim maxRetry = If(retry IsNot Nothing, retry.Value(Of Integer?)("max").GetValueOrDefault(0), 0)
                    Dim retryDelay = If(retry IsNot Nothing, retry.Value(Of Integer?)("delay_ms").GetValueOrDefault(1000), 1000)
                    Dim backoff = If(retry IsNot Nothing, retry.Value(Of Double?)("backoff").GetValueOrDefault(2.0), 2.0)

                    Dim attempt As Integer = 0
                    Dim success As Boolean = False
                    Dim lastEx As Exception = Nothing
                    Dim resultValue As Object = Nothing   ' kept (result of sub-step)

                    Do
                        lastEx = Nothing
                        Dim delayMs As Integer = 0
                        success = False
                        resultValue = Nothing

                        Dim parms = TryCast(stepObj("params"), JObject)

                        DebugLogSubStepStart(sid, cmd, attempt, maxRetry, parms)
                        Dim subSw = Diagnostics.Stopwatch.StartNew()

                        Try
                            _currentStepId = sid
                            Dim swExec = Diagnostics.Stopwatch.StartNew()

                            Select Case cmd.ToLowerInvariant()
                                Case "llm_analyze", "llm", "llmanalyze"
                                    resultValue = Await CmdLlmAnalyzeAsync(parms, timeoutMs, cancel)
                                Case "open_url"
                                    resultValue = Await CmdOpenUrlAsync(parms, timeoutMs, cancel)
                                Case "http_request"
                                    resultValue = Await CmdHttpRequestAsync(parms, timeoutMs, cancel)
                                Case "download_url"
                                    resultValue = Await CmdDownloadUrlAsync(parms, cancel)
                                Case "wait"
                                    resultValue = Await CmdWaitAsync(parms, cancel)
                                Case "foreach"
                                    resultValue = Await CmdForEachAsync(parms, cancel)
                                Case "if"
                                    resultValue = Await CmdIfAsync(parms, cancel)
                                Case "template"
                                    resultValue = CmdTemplate(parms)
                                Case "set_var"
                                    resultValue = CmdSetVar(parms)
                                Case "array_push"
                                    resultValue = CmdArrayPush(parms)
                                Case "extract_text"
                                    resultValue = CmdExtractText(parms)
                                Case "extract_html"
                                    resultValue = CmdExtractHtml(parms)
                                Case "extract_attribute"
                                    resultValue = CmdExtractAttribute(parms)
                                Case "render_report"
                                    resultValue = CmdRenderReport(parms)
                                Case "log"
                                    resultValue = CmdLog(parms)
                                Case Else
                                    Throw New Exception($"RunSubStepsAsync: unknown command '{cmd}'")
                            End Select

                            swExec.Stop()
                            success = True


                            Dim assign = TryCast(stepObj("assign"), JObject)
                            If assign IsNot Nothing Then
                                Dim varName = assign.Value(Of String)("var")
                                Dim path = assign.Value(Of String)("path")
                                If Not String.IsNullOrWhiteSpace(varName) Then
                                    Dim toStore As Object = resultValue
                                    If Not String.IsNullOrWhiteSpace(path) AndAlso resultValue IsNot Nothing Then
                                        Try
                                            Dim tokenObj = JToken.FromObject(resultValue)
                                            Dim sel = tokenObj.SelectToken(path)
                                            If sel IsNot Nothing Then
                                                toStore = sel.ToObject(Of Object)()
                                            Else
                                                toStore = Nothing
                                            End If
                                        Catch
                                            toStore = Nothing
                                        End Try
                                    End If
                                    SafeStoreVar(varName, toStore)
                                End If
                            End If

                        Catch ex As Exception
                            lastEx = ex
                            If TypeOf ex Is OperationCanceledException AndAlso cancel.IsCancellationRequested Then
                                subSw.Stop()
                                DebugLogSubStepResult(sid, cmd, subSw.ElapsedMilliseconds, False, lastEx, Nothing)
                                Throw
                            End If
                            If attempt < maxRetry Then
                                delayMs = CInt(retryDelay * Math.Pow(backoff, attempt))
                                Log($"[substep:{sid}] attempt={attempt + 1} error={ex.Message} retry in {delayMs}ms")
                            Else
                                delayMs = 0
                            End If
                        Finally
                            subSw.Stop()

                            DebugLogSubStepResult(sid, cmd, subSw.ElapsedMilliseconds, success, lastEx, resultValue)
                        End Try

                        attempt += 1
                        If Not success AndAlso attempt <= maxRetry AndAlso delayMs > 0 Then
                            Await System.Threading.Tasks.Task.Delay(delayMs, cancel)
                        End If

                    Loop While Not success AndAlso attempt <= maxRetry

                    If Not success Then
                        Throw New Exception($"Substep '{sid}' failed after {attempt} attempt(s).", lastEx)
                    End If

                    AccumulateDecisionLinksIfNeeded(sid)

                Catch oce As OperationCanceledException
                    If cancel.IsCancellationRequested Then
                        Throw
                    End If
                    Throw New Exception($"Internal cancellation in substep '{_currentStepId}'", oce)
                End Try
            Next
        End Function


        Private Function CmdRenderReport(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim engine = parms.Value(Of System.String)("engine")
            Dim tpl = parms.Value(Of System.String)("template")
            Dim ctxToken = parms("context")
            Dim outputPathRaw = parms.Value(Of System.String)("output_path")

            Dim ctx As System.Object = Nothing
            If ctxToken IsNot Nothing Then
                ctx = MaterializeContextPlaceholders(ctxToken)
            End If

            Dim md As System.String
            md = SimpleMustacheRender(tpl, ctx)
            md = ExpandTemplates(md)

            Dim outputPath As String = Nothing
            If Not String.IsNullOrWhiteSpace(outputPathRaw) Then
                outputPath = ExpandTemplates(outputPathRaw)
                ' If still contains "{{" warn & skip writing
                If outputPath?.Contains("{{") = True Then
                    Log("[render_report] Unresolved template tokens in output_path; skipping file write: " & outputPath)
                    outputPath = Nothing
                End If
            End If

            If Not String.IsNullOrWhiteSpace(outputPath) Then
                Try
                    Dim dir = System.IO.Path.GetDirectoryName(outputPath)
                    If Not System.String.IsNullOrEmpty(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
                        System.IO.Directory.CreateDirectory(dir)
                    End If
                    System.IO.File.WriteAllText(outputPath, md, System.Text.Encoding.UTF8)
                Catch ex As Exception
                    Log("[render_report] Failed to write file: " & ex.Message)
                End Try
            End If

            _finalMarkdown = md
            Return New With {.output = If(outputPath, "(memory)")}
        End Function

        Private Function CmdLog(parms As Newtonsoft.Json.Linq.JObject) As System.Object
            Dim level = parms.Value(Of System.String)("level")
            Dim message = ExpandTemplates(parms.Value(Of System.String)("message"))
            Log($"[{level}] {message}")
            Return Nothing
        End Function



        '   retry_on_invalid: Bool          -> Throw when output invalid (so outer step retry triggers)
        '   require_key: "key1,key2"        -> Comma list of required top-level JSON keys
        '   require_array_key: "decisions"  -> Single array key that must exist (or comma list)
        '   require_min_items: Integer      -> If require_array_key set, array must have at least this many items
        '   reject_if_empty: Bool           -> Treat empty (post-sanitize) output as invalid
        '   reject_if_plaintext: Bool       -> Treat non-JSON (parsed = Nothing) as invalid (default True)
        '   allow_non_json: Bool            -> Override reject_if_plaintext (forces accept even if not JSON)
        '   max_preview: Int                -> UI preview length (default 250)
        '   log_raw: Bool                   -> If True, write raw (trimmed) to debug log
        '   require_key_all: Bool           -> If True, ALL listed keys must exist; otherwise any missing => invalid
        '   timeoutMs (already supported)
        '
        ' OUTER RETRY:
        ' Add to step:
        '   "retry": { "max": 2, "delay_ms": 2000, "backoff": 1.5 }
        ' plus params { "retry_on_invalid": true }
        '
        Private Async Function CmdLlmAnalyzeAsync(parms As Newtonsoft.Json.Linq.JObject,
                                              timeoutMs As Int32,
                                              cancel As Threading.CancellationToken) As Task(Of Object)
            ' Hardened + retry-aware + quick inner attempts.
            If _context Is Nothing Then Return ""

            Dim statusVar = parms?.Value(Of String)("status_var")
            If Not String.IsNullOrEmpty(statusVar) AndAlso _vars.ContainsKey(statusVar) Then
                Dim st = _vars(statusVar)?.ToString()
                If st = "404" AndAlso Not GetDebugFlag("allow_llm_on_404") Then
                    WriteDebug("[INFO llm_analyze] Skipped due to 404 status (allow_llm_on_404=true to override).")
                    Return Nothing
                End If
            End If

            ' Helpers
            Dim GetBool =
            Function(name As String, def As Boolean) As Boolean
                Dim tok = parms(name)
                If tok Is Nothing Then Return def
                If tok.Type = JTokenType.Boolean Then Return tok.Value(Of Boolean)()
                Dim s = tok.ToString().Trim()
                Dim b As Boolean
                If Boolean.TryParse(s, b) Then Return b
                If s = "1" Then Return True
                If s = "0" Then Return False
                Return def
            End Function

            Dim GetInt =
            Function(name As String, def As Integer) As Integer
                Dim tok = parms(name)
                If tok Is Nothing Then Return def
                Dim i As Integer
                If Integer.TryParse(tok.ToString().Trim(), i) Then Return i
                Return def
            End Function

            Dim retryOnInvalid = GetBool("retry_on_invalid", False)
            Dim rejectIfEmpty = GetBool("reject_if_empty", False)
            Dim rejectIfPlain = GetBool("reject_if_plaintext", True)
            Dim allowNonJson = GetBool("allow_non_json", False)
            If allowNonJson Then rejectIfPlain = False
            Dim logRaw = GetBool("log_raw", False)
            Dim requireAll = GetBool("require_key_all", True)
            Dim maxPreview = GetInt("max_preview", 250)
            Dim requireMinItems = GetInt("require_min_items", 0)

            Dim requiredKeysRaw = parms.Value(Of String)("require_key")
            Dim requiredArrayKeysRaw = parms.Value(Of String)("require_array_key")

            Dim requiredKeys As New List(Of String)
            If Not String.IsNullOrWhiteSpace(requiredKeysRaw) Then
                requiredKeys.AddRange(requiredKeysRaw.Split(","c).Select(Function(s) s.Trim()).Where(Function(s) s.Length > 0))
            End If

            Dim requiredArrayKeys As New List(Of String)
            If Not String.IsNullOrWhiteSpace(requiredArrayKeysRaw) Then
                requiredArrayKeys.AddRange(requiredArrayKeysRaw.Split(","c).Select(Function(s) s.Trim()).Where(Function(s) s.Length > 0))
            End If

            Try
                ' Prompts
                Dim systemPrompt As String =
                If(parms("system")?.ToString(),
                   If(parms("systemPrompt")?.ToString(), ""))

                Dim userPrompt As String =
                If(parms("user")?.ToString(),
                   If(parms("prompt")?.ToString(),
                      If(parms("input")?.ToString(),
                         If(parms("arguments")?.ToString(), ""))))

                systemPrompt = ExpandTemplates(systemPrompt)
                userPrompt = ExpandTemplates(userPrompt)

                Dim currentUrl As String = If(_lastResponseUrl, String.Empty)
                Log($"[llm-meta] step={_currentStepId} url={currentUrl}")
                Log($"[llm] system length={systemPrompt.Length} user length={userPrompt.Length}")

                ' Temperature
                Dim tempStr As String =
                If(parms("temperature") IsNot Nothing,
                   System.Convert.ToString(parms("temperature"), System.Globalization.CultureInfo.InvariantCulture),
                   If(_useSecondAPI, _context.INI_Temperature_2, _context.INI_Temperature))

                ' Timeout
                Dim timeoutLong As Long
                If parms("timeoutMs") IsNot Nothing Then
                    timeoutLong = CLng(parms("timeoutMs"))
                ElseIf timeoutMs > 0 Then
                    timeoutLong = timeoutMs
                Else
                    timeoutLong = If(_useSecondAPI, _context.INI_Timeout_2, _context.INI_Timeout)
                End If

                ' Dynamic model 
                If _useSecondAPI AndAlso _autoselectModel Then
                    If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                        If Not GetSpecialTaskModel(_context, _context.INI_AlternateModelPath, "WebAgent") Then
                            originalConfigLoaded = False
                            _useSecondAPI = False
                        Else
                            _useSecondAPI = True
                        End If
                    End If
                End If

                Dim innerAttempts = GetInt("inner_attempts", 1)
                If innerAttempts < 1 Then innerAttempts = 1
                Dim innerDelayMs = GetInt("inner_delay_ms", 800)

                Dim rawResult As String = Nothing
                Dim inner As Integer = 0
                Dim overallSw = Diagnostics.Stopwatch.StartNew()

                Do
                    Dim attemptSw = Diagnostics.Stopwatch.StartNew()
                    Try

                        If userPrompt?.Contains("{{") Then
                            Log($"[llm warn] unresolved placeholders in user prompt (step={_currentStepId}).")
                        End If

                        Using perCallCts As New CancellationTokenSource(CInt(timeoutLong))
                            Using linked = CancellationTokenSource.CreateLinkedTokenSource(cancel, perCallCts.Token)
                                rawResult = Await SharedMethods.LLM(
                                    _context,
                                    systemPrompt,
                                    userPrompt,
                                    "",
                                    tempStr,
                                    timeoutLong,
                                    _useSecondAPI,
                                    True,
                                    "",
                                    "",
                                    linked.Token)
                            End Using
                        End Using


                    Catch tce As TaskCanceledException
                        Log("[llm timeout] " & tce.Message)
                        Throw
                    Catch oce As OperationCanceledException
                        Log("[llm canceled] " & oce.Message)
                        Throw
                    Catch exHttp As Http.HttpRequestException
                        Log("[llm http] " & exHttp.Message)
                        Throw
                    End Try
                    attemptSw.Stop()

                    If rawResult Is Nothing Then rawResult = ""
                    Dim probeSanitized = SanitizeLlmResult(rawResult)
                    Dim probeParsed = TryParseJson(probeSanitized)

                    ' Stop if JSON parsed OK OR we exhausted attempts
                    If probeParsed IsNot Nothing OrElse inner + 1 >= innerAttempts Then
                        Exit Do
                    End If

                    inner += 1
                    Log($"[llm inner retry] invalid/plain attempt={inner}/{innerAttempts} waiting {innerDelayMs}ms")
                    If innerDelayMs > 0 Then
                        Await System.Threading.Tasks.Task.Delay(innerDelayMs, cancel)
                    End If
                Loop While inner < innerAttempts

                overallSw.Stop()
                _vars("lastLlm_latency_ms") = overallSw.ElapsedMilliseconds
                _vars("lastLlm_raw") = If(rawResult, "")


                ' Restore model if temporarily switched
                If _autoselectModel AndAlso _useSecondAPI AndAlso originalConfigLoaded Then
                    RestoreDefaults(_context, originalConfig)
                    originalConfigLoaded = False
                End If

                If rawResult Is Nothing Then rawResult = ""

                Dim sanitized = SanitizeLlmResult(rawResult)
                Dim parsed As JToken = TryParseJson(sanitized)

                Dim invalid As Boolean = False
                Dim invalidReasons As New List(Of String)

                If rejectIfEmpty AndAlso String.IsNullOrWhiteSpace(sanitized) Then
                    invalid = True
                    invalidReasons.Add("empty_output")
                End If

                If parsed Is Nothing AndAlso rejectIfPlain Then
                    invalid = True
                    invalidReasons.Add("non_json")
                End If

                Dim rootObj As JObject = Nothing
                If parsed IsNot Nothing Then
                    rootObj = TryCast(parsed, JObject)
                End If

                If requiredKeys.Count > 0 AndAlso rootObj IsNot Nothing Then
                    Dim missing = requiredKeys.Where(Function(k) rootObj(k) Is Nothing).ToList()
                    If requireAll AndAlso missing.Count > 0 Then
                        invalid = True
                        invalidReasons.Add("missing_keys:" & String.Join(",", missing))
                    ElseIf Not requireAll AndAlso missing.Count = requiredKeys.Count Then
                        invalid = True
                        invalidReasons.Add("no_required_keys_present")
                    End If
                ElseIf requiredKeys.Count > 0 AndAlso parsed Is Nothing Then
                    invalid = True
                    invalidReasons.Add("missing_required_keys_non_object")
                End If

                If requiredArrayKeys.Count > 0 AndAlso rootObj IsNot Nothing Then
                    For Each arrKey In requiredArrayKeys
                        Dim tk = rootObj(arrKey)
                        If tk Is Nothing OrElse tk.Type <> JTokenType.Array Then
                            invalid = True
                            invalidReasons.Add($"array_key_missing_or_not_array:{arrKey}")
                        ElseIf requireMinItems > 0 AndAlso tk.Count() < requireMinItems Then
                            invalid = True
                            invalidReasons.Add($"array_key_min_items_not_met:{arrKey}(<{requireMinItems})")
                        End If
                    Next
                End If

                Dim finalObj As Object
                If parsed Is Nothing Then
                    finalObj = New JObject(
                    New JProperty("_invalid", True),
                    New JProperty("_reason", If(String.Join(";", invalidReasons), "parse_failed")),
                    New JProperty("raw_length", sanitized.Length),
                    New JProperty("step_id", _currentStepId),
                    New JProperty("page_url", currentUrl)
                )
                ElseIf TypeOf parsed Is JObject Then
                    Dim jobj = DirectCast(parsed.DeepClone(), JObject)
                    jobj("step_id") = _currentStepId
                    jobj("page_url") = currentUrl
                    If invalid Then
                        jobj("_invalid") = True
                        jobj("_reason") = String.Join(";", invalidReasons)
                    End If
                    finalObj = jobj
                Else
                    Dim wrap = New JObject(
                    New JProperty("value", parsed.ToString()),
                    New JProperty("step_id", _currentStepId),
                    New JProperty("page_url", currentUrl)
                )
                    If invalid Then
                        wrap("_invalid") = True
                        wrap("_reason") = String.Join(";", invalidReasons)
                    End If
                    finalObj = wrap
                End If

                SafeStoreVar("lastLlm", finalObj)
                SafeStoreVar("lastLlm_page_url", currentUrl)
                SafeStoreVar("last_step_id", _currentStepId)

                If Not String.IsNullOrWhiteSpace(sanitized) Then
                    Dim preview = "[step:" & _currentStepId & "] [url:" & currentUrl & "] " &
                              If(sanitized.Length > maxPreview, sanitized.Substring(0, maxPreview) & "...", sanitized)
                    Try : InfoBox.ShowInfoBox(preview, 1) : Catch : End Try
                End If

                If logRaw AndAlso (GetDebugFlag("debug") OrElse GetDebugFlag("debug_allAttempts")) Then
                    Try
                        InitDebugLogIfNeeded()
                        If _debugInitialized Then
                            Dim trimmed = rawResult
                            If trimmed Is Nothing Then trimmed = ""
                            If trimmed.Length > 8000 Then trimmed = trimmed.Substring(0, 8000) & "…"
                            WriteDebug(New String("-"c, 50))
                            WriteDebug($"[llm raw] step={_currentStepId} url={currentUrl} len={rawResult.Length}")
                            WriteDebug(trimmed)
                        End If
                    Catch
                    End Try
                End If

                Dim forceThrowInvalid = retryOnInvalid
                If _vars.ContainsKey("allow_llm_invalid") AndAlso _IsTruthy(_vars("allow_llm_invalid")) Then
                    forceThrowInvalid = False
                End If

                If invalid AndAlso forceThrowInvalid Then
                    Log($"[llm invalid] step={_currentStepId} reasons={String.Join(";", invalidReasons)} retry_on_invalid=True")
                    Throw New Exception("llm_analyze invalid: " & String.Join(";", invalidReasons))
                End If

                Return finalObj

            Catch ex As TaskCanceledException
                Try : Log("[llm exception] TaskCanceled: " & ex.Message) : Catch : End Try
                Throw
            Catch ex As OperationCanceledException
                Try : Log("[llm exception] OpCanceled: " & ex.Message) : Catch : End Try
                Throw
            Catch ex As Exception
                Dim forceThrow = GetDebugFlag("llm_rethrow_all")
                Try : Log("[llm error] " & ex.Message) : Catch : End Try
                If forceThrow Then Throw
                Return New JObject(
                New JProperty("_invalid", True),
                New JProperty("_exception", ex.GetType().Name),
                New JProperty("_message", ex.Message),
                New JProperty("step_id", _currentStepId)
            )
            End Try
        End Function

        Public Shared Function SanitizeLlmResult(raw As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(raw) Then Return ""
            Dim codeBlocks As New System.Collections.Generic.List(Of System.String)
            Try
                Dim reFences As New System.Text.RegularExpressions.Regex("(?s)```(?:[A-Za-z0-9_+-]*)\s*(.*?)```")
                Dim matches = reFences.Matches(raw)
                For Each m As System.Text.RegularExpressions.Match In matches
                    If m.Groups.Count > 1 Then
                        codeBlocks.Add(m.Groups(1).Value.Trim())
                    End If
                Next
            Catch
            End Try
            For Each blk In codeBlocks
                Dim j = TryParseJson(blk)
                If j IsNot Nothing Then
                    Return blk.Trim()
                End If
            Next
            If codeBlocks.Count > 0 Then
                raw = System.Text.RegularExpressions.Regex.Replace(raw,
                "(?s)```(?:[A-Za-z0-9_+-]*)\s*(.*?)```",
                Function(m) m.Groups(1).Value.Trim() & System.Environment.NewLine)
            End If
            Dim candidate = raw.Trim()
            If TryParseJson(candidate) IsNot Nothing Then
                Return candidate
            End If
            Dim extracted = ExtractFirstJsonStructure(candidate)
            If Not System.String.IsNullOrWhiteSpace(extracted) AndAlso TryParseJson(extracted) IsNot Nothing Then
                Return extracted.Trim()
            End If
            candidate = candidate.Trim("`"c, Chr(10), Chr(13)).Trim()
            Return candidate
        End Function

        Private Shared Function TryParseJson(text As System.String) As Newtonsoft.Json.Linq.JToken
            If System.String.IsNullOrWhiteSpace(text) Then Return Nothing
            Try
                Return Newtonsoft.Json.Linq.JToken.Parse(text)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ExtractFirstJsonStructure(text As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(text) Then Return Nothing
            Dim functionExtract =
            Function(openChar As Char, closeChar As Char) As System.String
                Dim depth As System.Int32 = 0
                Dim startIdx As System.Int32 = -1
                For i = 0 To text.Length - 1
                    Dim c = text(i)
                    If c = openChar Then
                        If depth = 0 Then startIdx = i
                        depth += 1
                    ElseIf c = closeChar Then
                        If depth > 0 Then
                            depth -= 1
                            If depth = 0 AndAlso startIdx >= 0 Then
                                Dim seg = text.Substring(startIdx, i - startIdx + 1)
                                Return seg
                            End If
                        End If
                    End If
                Next
                Return Nothing
            End Function
            Dim best = functionExtract("{"c, "}"c)
            If Not System.String.IsNullOrWhiteSpace(best) Then Return best
            best = functionExtract("["c, "]"c)
            Return best
        End Function

#End Region

#Region "Selector Resolution"

        Private Function ResolveSelector(sel As Newtonsoft.Json.Linq.JObject) As System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)
            If _lastDoc Is Nothing Then Throw New System.Exception("No document loaded. Call open_url or http_request first.")
            Dim strategy = sel.Value(Of System.String)("strategy")
            Dim value = ExpandTemplates(sel.Value(Of System.String)("value"))
            Dim container As System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode) = Nothing
            Dim within = TryCast(sel("within"), Newtonsoft.Json.Linq.JObject)
            If within IsNot Nothing Then
                container = ResolveSelector(within)
            End If

            Dim baseNodes As System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode) =
            If(container, New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)({_lastDoc.DocumentNode}))

            Dim matches As New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)()

            Select Case strategy
                Case "xpath"
                    For Each root In baseNodes
                        Dim ns = root.SelectNodes(value)
                        If ns IsNot Nothing Then matches.AddRange(ns)
                    Next
                Case "css"
                    Dim xp = CssToXPath(value)
                    For Each root In baseNodes
                        Dim ns = root.SelectNodes(xp)
                        If ns IsNot Nothing Then matches.AddRange(ns)
                    Next
                Case "text"
                    Dim exact As System.Boolean = False
                    If value.StartsWith("exact:", System.StringComparison.OrdinalIgnoreCase) Then
                        exact = True
                        value = value.Substring(6)
                    End If
                    For Each root In baseNodes
                        Dim ns = root.SelectNodes(".//*")
                        If ns Is Nothing Then Continue For
                        For Each n In ns
                            Dim t = GetInnerText(n, True)
                            If exact Then
                                If System.String.Equals(t, value, System.StringComparison.OrdinalIgnoreCase) Then matches.Add(n)
                            Else
                                If Not System.String.IsNullOrEmpty(t) AndAlso t.IndexOf(value, System.StringComparison.OrdinalIgnoreCase) >= 0 Then matches.Add(n)
                            End If
                        Next
                    Next
                Case "regex"
                    Dim re As New System.Text.RegularExpressions.Regex(value, System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
                    For Each root In baseNodes
                        Dim ns = root.SelectNodes(".//*")
                        If ns Is Nothing Then Continue For
                        For Each n In ns
                            Dim t = GetInnerText(n, True)
                            If re.IsMatch(t) Then matches.Add(n)
                        Next
                    Next
                Case Else
                    Throw New System.Exception($"Unknown strategy: {strategy}")
            End Select

            Dim relative = TryCast(sel("relative"), Newtonsoft.Json.Linq.JObject)
            If relative IsNot Nothing AndAlso matches.Count > 0 Then
                Dim position = relative.Value(Of System.String)("position")
                If System.String.Equals(position, "first", System.StringComparison.OrdinalIgnoreCase) Then
                    matches = New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)({matches(0)})
                ElseIf System.String.Equals(position, "last", System.StringComparison.OrdinalIgnoreCase) Then
                    matches = New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)({matches(matches.Count - 1)})
                ElseIf System.String.Equals(position, "nth", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim nth = relative.Value(Of System.Nullable(Of System.Int32))("nth").GetValueOrDefault(1)
                    Dim idx = System.Math.Max(1, nth) - 1
                    If idx >= 0 AndAlso idx < matches.Count Then
                        matches = New System.Collections.Generic.List(Of HtmlAgilityPack.HtmlNode)({matches(idx)})
                    Else
                        matches.Clear()
                    End If
                End If
            End If

            Return matches
        End Function

        Private Function CssToXPath(css As System.String) As System.String
            Dim parts = css.Split({" "c}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim xpath As New System.Text.StringBuilder()
            xpath.Append("//*")
            Dim directChild As System.Boolean = False

            For Each raw In parts
                Dim token = raw.Trim()
                If token = ">" Then
                    directChild = True
                    Continue For
                End If
                Dim segment = CssSimpleSelectorToXPath(token)
                If directChild Then
                    xpath.Append("/").Append(segment)
                Else
                    xpath.Append("//").Append(segment)
                End If
                directChild = False
            Next

            Return xpath.ToString()
        End Function

        Private Function CssSimpleSelectorToXPath(token As System.String) As System.String
            Dim mTag = System.Text.RegularExpressions.Regex.Match(token, "^[a-zA-Z][a-zA-Z0-9_-]*")
            Dim tag = If(mTag.Success, mTag.Value, "*")
            Dim rest = token.Substring(tag.Length)
            Dim xp As New System.Text.StringBuilder(tag)

            For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(rest, "#([a-zA-Z0-9_-]+)")
                xp.AppendFormat("[@id='{0}']", m.Groups(1).Value)
            Next
            For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(rest, "\.([a-zA-Z0-9_-]+)")
                xp.AppendFormat("[contains(concat(' ', normalize-space(@class), ' '), ' {0} ')]", m.Groups(1).Value)
            Next
            For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(rest, "\[([a-zA-Z0-9_-]+)(?:=(?:'([^']*)'|""([^""]*)""|([^\]]+)))?\]")
                Dim attr = m.Groups(1).Value
                Dim val = If(m.Groups(2).Success, m.Groups(2).Value,
                         If(m.Groups(3).Success, m.Groups(3).Value,
                            If(m.Groups(4).Success, m.Groups(4).Value, Nothing)))
                If val Is Nothing Then
                    xp.AppendFormat("[@{0}]", attr)
                Else
                    xp.AppendFormat("[@{0}='{1}']", attr, val)
                End If
            Next
            Dim nth = System.Text.RegularExpressions.Regex.Match(rest, ":(?:nth-child)\((\d+)\)")
            If nth.Success Then
                xp.AppendFormat("[position()={0}]", nth.Groups(1).Value)
            End If
            Return xp.ToString()
        End Function

        Private Function GetInnerText(node As HtmlAgilityPack.HtmlNode, normalize As System.Boolean) As System.String
            Dim t = node.InnerText
            If normalize Then
                t = System.Text.RegularExpressions.Regex.Replace(t, "\s+", " ").Trim()
            End If
            Return t
        End Function

        Private Function SerializeNode(n As HtmlAgilityPack.HtmlNode) As System.Object
            Dim dict As New System.Collections.Generic.Dictionary(Of System.String, System.Object)()
            dict("name") = n.Name
            dict("innerText") = GetInnerText(n, True)
            Dim atts As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each a In n.Attributes
                atts(a.Name) = a.Value
            Next
            dict("attributes") = atts
            dict("outerHtml") = n.OuterHtml
            Return dict
        End Function

        Private Sub LoadDocument(html As String)
            ' Fallback (avoid using on UI thread if dynamic enabled)
            LoadDocumentAsync(html, _lastResponseUrl, CancellationToken.None).GetAwaiter().GetResult()
        End Sub

        Private Async Function LoadDocumentAsync(html As String,
                                     sourceUrl As String,
                                     cancel As CancellationToken) As System.Threading.Tasks.Task
            _lastResponseBody = html
            _lastResponseUrl = sourceUrl
            _lastDoc = New HtmlAgilityPack.HtmlDocument()
            _lastDoc.LoadHtml(_lastResponseBody)

            If _dynamicExpand Then
                Log("[dynamic] starting expansion")
                Using cts As New CancellationTokenSource(_defaultTimeoutMs)
                    Using linked As CancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(cancel, cts.Token)
                        Try
                            Dim expanded = Await ExpandDynamicContentAsync(sourceUrl, _lastResponseBody, linked.Token).ConfigureAwait(False)
                            If Not String.IsNullOrEmpty(expanded) AndAlso Not ReferenceEquals(expanded, _lastResponseBody) Then
                                _lastResponseBody = expanded
                                _lastDoc = New HtmlAgilityPack.HtmlDocument()
                                _lastDoc.LoadHtml(_lastResponseBody)
                                Log("[dynamic] expansion complete & re-parsed")
                            Else
                                Log("[dynamic] expansion produced no changes")
                            End If
                        Catch ex As OperationCanceledException
                            Log("[dynamic] expansion cancelled/timeout: " & ex.Message)
                        Catch ex As Exception
                            Log("[dynamic] expansion error: " & ex.Message)
                        End Try
                    End Using
                End Using
            End If
        End Function


        ' Params:
        '   {
        '     "array": "summary_array",          ' Ziel-Variablenname (wird JArray)
        '     "item_var": "decision_summary_obj" ' (ODER) item_var: Name einer bestehenden Variable
        '     "item": { ... }                    ' (ODER) inline JSON Objekt / Wert
        '   }
        '   item_var hat Vorrang; fallback auf inline "item".
        ' Rückgabe: { pushed = True, count = <Anzahl>, array = "<name>" }
        Private Function CmdArrayPush(parms As JObject) As Object
            If parms Is Nothing Then Throw New Exception("array_push: params missing")
            Dim arrayName = parms.Value(Of String)("array")
            Dim itemVar = parms.Value(Of String)("item_var")

            If String.IsNullOrWhiteSpace(arrayName) Then
                Throw New Exception("array_push: 'array' missing")
            End If

            ' 1) Item bestimmen
            Dim itemObj As Object = Nothing
            If Not String.IsNullOrWhiteSpace(itemVar) Then
                If Not _vars.TryGetValue(itemVar, itemObj) OrElse itemObj Is Nothing Then
                    ' Nichts zu pushen – kein Fehler, nur noop
                    Return New With {.pushed = False, .count = GetExistingArrayCount(arrayName), .array = arrayName}
                End If
            Else
                Dim inlineToken = parms("item")
                If inlineToken Is Nothing Then
                    Throw New Exception("array_push: neither 'item_var' nor inline 'item' present")
                End If
                itemObj = inlineToken.ToObject(Of Object)()
            End If

            ' 2) Existierendes Array holen / herstellen (immer auf JArray normalisieren)
            Dim arr As JArray = Nothing
            If _vars.ContainsKey(arrayName) Then
                Select Case True
                    Case TypeOf _vars(arrayName) Is JArray
                        arr = DirectCast(_vars(arrayName), JArray)
                    Case TypeOf _vars(arrayName) Is String
                        ' Versuchen zu parsen falls String wie "[]"
                        Dim s = DirectCast(_vars(arrayName), String).Trim()
                        If s.StartsWith("[") AndAlso s.EndsWith("]") Then
                            Try : arr = JArray.Parse(s) : Catch : End Try
                        End If
                    Case Else
                        ' Versuch generisch zu konvertieren
                        Try
                            arr = JArray.FromObject(_vars(arrayName))
                        Catch
                        End Try
                End Select
            End If
            If arr Is Nothing Then arr = New JArray()

            ' 3) Item als JToken klonen und anhängen
            Dim jt As JToken
            If TypeOf itemObj Is JToken Then
                jt = DirectCast(itemObj, JToken).DeepClone()
            Else
                jt = JToken.FromObject(itemObj)
            End If
            arr.Add(jt)

            ' 4) Zurückspeichern (direkt als JArray, kein Doppelkonvert via SafeStoreVar nötig)
            _vars(arrayName) = arr

            Return New With {
    .pushed = True,
    .count = arr.Count,
    .array = arrayName
}
        End Function

        Private Sub AccumulateDecisionLinksIfNeeded(stepId As String)
            If Not String.Equals(stepId, "extract_decision_links", StringComparison.OrdinalIgnoreCase) Then Exit Sub
            Dim pageObj As Object = Nothing
            If Not _vars.TryGetValue("decision_links", pageObj) OrElse Not TypeOf pageObj Is JArray Then Exit Sub
            Dim allObj As Object = Nothing
            If Not _vars.TryGetValue("all_decisions", allObj) OrElse Not TypeOf allObj Is JArray Then
                allObj = New JArray()
                _vars("all_decisions") = allObj
            End If
            For Each d In CType(pageObj, JArray)
                CType(allObj, JArray).Add(d.DeepClone())
            Next
        End Sub

        ' Hilfsfunktion für frühen Rückgabewert
        Private Function GetExistingArrayCount(name As String) As Integer
            If Not _vars.ContainsKey(name) Then Return 0
            If TypeOf _vars(name) Is JArray Then Return DirectCast(_vars(name), JArray).Count
            Return 0
        End Function

        Private Function CmdEnableDynamic(parms As JObject) As Object
            EnableDynamicExpansion()
            Return New JObject(
            New JProperty("status", "ok"),
            New JProperty("dynamic", True)
        )
        End Function

        Private Shared ReadOnly DynamicUrlRegexes As Regex() = {
            New Regex("(https?://[^\s'""<>]+index_aza\.php[^\s'""<>]*)", RegexOptions.IgnoreCase),
            New Regex("url\s*:\s*['""]([^'""]*index_aza\.php[^'""]*)['""]", RegexOptions.IgnoreCase),
            New Regex("\$\.(?:get|post)\s*\(\s*['""]([^'""]*index_aza\.php[^'""]*)['""]", RegexOptions.IgnoreCase),
            New Regex("fetch\s*\(\s*['""]([^'""]*index_aza\.php[^'""]*)['""]", RegexOptions.IgnoreCase)
        }

        Private Sub EnableDynamicExpansion()
            _dynamicExpand = True
            Log("Dynamic expansion ENABLED")
        End Sub

        Private Function CmdDisableDynamic(parms As JObject) As Object
            _dynamicExpand = False
            Log("Dynamic expansion DISABLED")
            Return New JObject(New JProperty("status", "ok"), New JProperty("dynamic", False))
        End Function

        ' Core dynamic expansion after initial HTML load
        Private Async Function ExpandDynamicContentAsync(baseUrl As String,
                                             originalHtml As String,
                                             cancel As CancellationToken) As Task(Of String)
            Dim composite As New StringBuilder(originalHtml)
            Dim discovered As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim queue As New Queue(Of String)()
            Dim fetchCount As Integer = 0

            If originalHtml.Length > 1_500_000 Then
                Log("[dynamic] skipping expansion (page too large)")
                Return originalHtml
            End If

            Try
                If _lastDoc IsNot Nothing Then
                    Dim scriptNodes = _lastDoc.DocumentNode.SelectNodes("//script[@src]")
                    If scriptNodes IsNot Nothing Then
                        For Each s In scriptNodes
                            Dim src As String = s.GetAttributeValue("src", "")
                            If Not String.IsNullOrWhiteSpace(src) Then
                                Dim abs = ResolveUrl(src)
                                If ShouldFetchDynamic(abs, discovered, queue) Then queue.Enqueue(abs)
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                Log("[dynamic] script src scan error: " & ex.Message)
            End Try

            Try
                If _lastDoc IsNot Nothing Then
                    Dim inlineScripts = _lastDoc.DocumentNode.SelectNodes("//script[not(@src)]")
                    If inlineScripts IsNot Nothing Then
                        For Each s In inlineScripts
                            Dim code = s.InnerText
                            If String.IsNullOrWhiteSpace(code) Then Continue For
                            For Each rx In DynamicUrlRegexes
                                Dim matches = rx.Matches(code)
                                For Each m As Match In matches
                                    Dim cand = m.Groups(m.Groups.Count - 1).Value
                                    If String.IsNullOrWhiteSpace(cand) Then Continue For
                                    Dim abs = ResolveUrl(cand.Trim())
                                    If ShouldFetchDynamic(abs, discovered, queue) Then queue.Enqueue(abs)
                                Next
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Log("[dynamic] inline script scan error: " & ex.Message)
            End Try

            While queue.Count > 0 AndAlso fetchCount < MAX_DYNAMIC_FETCH AndAlso Not cancel.IsCancellationRequested
                Dim u = queue.Dequeue()
                fetchCount += 1
                Try
                    Log($"[dynamic] fetching {u}")
                    Using ctsSingle As New CancellationTokenSource(_defaultTimeoutMs)
                        Using linked = CancellationTokenSource.CreateLinkedTokenSource(cancel, ctsSingle.Token)
                            Dim resp = Await _http.GetAsync(u, linked.Token).ConfigureAwait(False)
                            Dim body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                            composite.AppendLine().AppendLine($"<!-- DYNAMIC_FETCH: {u} -->").AppendLine(body)
                        End Using
                    End Using
                Catch ex As OperationCanceledException
                    Log("[dynamic] fetch timeout/cancel: " & u)
                Catch ex As Exception
                    Log("[dynamic] fetch error " & u & ": " & ex.Message)
                End Try
            End While

            Log($"[dynamic] fetched={fetchCount} (queue left={queue.Count})")
            Return composite.ToString()
        End Function

        Private Function ShouldFetchDynamic(url As String,
                                discovered As HashSet(Of String),
                                queue As Queue(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(url) Then Return False
            If discovered.Contains(url) Then Return False
            ' Domain constraint (basic)
            Try
                Dim u = New Uri(url)
                ' Optionally enforce alloweddomains if configured
                ' If Not String.IsNullOrEmpty(SharedMethods.alloweddomains) Then ...
            Catch
                Return False
            End Try
            discovered.Add(url)
            Return True
        End Function


#End Region

#Region "Templating / Helpers"

        ' Replace existing ResolveSecret with:
        Private Function ResolveSecret(reference As String) As String
            If String.IsNullOrEmpty(reference) Then Return reference
            If reference.StartsWith("secret://", StringComparison.OrdinalIgnoreCase) Then
                Dim key = reference.Substring(9)
                Dim val As String = Nothing
                If _secrets.TryGetValue(key, val) Then
                    Return val
                End If
                Return ""
            End If
            Return reference
        End Function
        Private Function ResolveUrl(url As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(url) Then Return url

            url = SanitizePotentialMarkdownUrl(url)

            If url.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase) OrElse
           url.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then
                Return url
            End If

            If Not System.String.IsNullOrWhiteSpace(_lastResponseUrl) Then
                Try
                    Dim baseUri As New System.Uri(_lastResponseUrl)
                    Return New System.Uri(baseUri, url).ToString()
                Catch ex As System.Exception
                    Log("[ResolveUrl] Failed combining with lastResponseUrl: " & ex.Message)
                End Try
            End If

            If Not System.String.IsNullOrWhiteSpace(_baseUrl) Then
                Try
                    Dim baseUri As New System.Uri(_baseUrl)
                    Return New System.Uri(baseUri, url).ToString()
                Catch ex As System.Exception
                    Log("[ResolveUrl] Failed combining with base_url: " & ex.Message)
                End Try
            End If

            Return url
        End Function

        Private Function SanitizePotentialMarkdownUrl(raw As System.String) As System.String
            Dim s = raw.Trim()
            ' Pattern: [visible](actual) – prefer the target inside parentheses if valid
            Dim mdMatch = System.Text.RegularExpressions.Regex.Match(s, "^\[(?<vis>[^\]]+)\]\((?<url>https?://[^)]+)\)$",
                                                                 System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If mdMatch.Success Then
                Dim realUrl = mdMatch.Groups("url").Value
                If realUrl.StartsWith("http", System.StringComparison.OrdinalIgnoreCase) Then
                    Return realUrl
                End If
            End If
            ' Case where only [https://...](https://...) was pasted but expansion removed brackets wrongly:
            ' remove stray leading '[' or trailing ')' if both present
            If s.Contains("](") AndAlso s.EndsWith(")") Then
                ' Try to pick substring after "]("
                Dim idx = s.IndexOf("](")
                If idx >= 0 Then
                    Dim candidate = s.Substring(idx + 2, s.Length - idx - 3)
                    If candidate.StartsWith("http", System.StringComparison.OrdinalIgnoreCase) Then
                        Return candidate
                    End If
                End If
            End If
            ' Strip surrounding angle brackets <...>
            If s.StartsWith("<") AndAlso s.EndsWith(">") Then
                s = s.Substring(1, s.Length - 2)
            End If
            Return s
        End Function

        Private Function DecodeBody(bytes As System.Byte(), contentType As System.Net.Http.Headers.MediaTypeHeaderValue) As System.String
            Dim charset As System.String = Nothing
            If contentType IsNot Nothing AndAlso Not System.String.IsNullOrEmpty(contentType.CharSet) Then
                charset = contentType.CharSet
            End If
            Dim enc As System.Text.Encoding
            Try
                enc = If(Not System.String.IsNullOrWhiteSpace(charset), System.Text.Encoding.GetEncoding(charset), System.Text.Encoding.UTF8)
            Catch
                enc = System.Text.Encoding.UTF8
            End Try
            Return enc.GetString(bytes)
        End Function

        Private Sub SafeSetHeader(col As System.Net.Http.Headers.HttpRequestHeaders, name As System.String, value As System.String)
            If System.String.Equals(name, "User-Agent", System.StringComparison.OrdinalIgnoreCase) Then
                col.UserAgent.Clear()
                col.UserAgent.ParseAdd(value)
                Return
            End If
            col.Remove(name)
            col.TryAddWithoutValidation(name, value)
        End Sub

        Private Function MaskSecrets(line As String) As String
            If _secrets Is Nothing OrElse _secrets.Count = 0 Then Return line
            Dim masked = line
            For Each kv In _secrets
                If Not String.IsNullOrEmpty(kv.Value) Then
                    masked = masked.Replace(kv.Value, "***")
                End If
            Next
            Return masked
        End Function

        Private Sub Log(msg As String)
            Dim safe = MaskSecrets(msg)
            Dim line = $"[{System.DateTime.Now:O}] {safe}"
            _log.AppendLine(line)
            Try : System.Diagnostics.Debug.WriteLine(line) : Catch : End Try
        End Sub

        Private Function ExpandTemplates(input As System.String) As System.String
            If input Is Nothing Then Return Nothing
            Dim unresolved As New System.Collections.Generic.List(Of System.String)

            Dim result = System.Text.RegularExpressions.Regex.Replace(
    input,
    "{{\s*([^}]+)\s*}}",
    Function(m)
        Dim expr = m.Groups(1).Value
        Dim v = ResolveValue("{{" & expr & "}}")
        If v Is Nothing Then
            unresolved.Add(expr)
            ' Keep the original marker so downstream code can detect unresolved tokens
            Return "{{" & expr & "}}"
        End If
        Return v.ToString()
    End Function)

            If unresolved.Count > 0 Then
                For Each u In unresolved
                    Log($"[template] Unresolved placeholder '{u}'")
                Next
            End If
            Return result
        End Function

        Private Function ResolveValue(exprOrLiteral As System.Object) As System.Object
            If exprOrLiteral Is Nothing Then Return Nothing
            Dim s = TryCast(exprOrLiteral, System.String)
            If s Is Nothing Then Return exprOrLiteral

            If s.StartsWith("{{") AndAlso s.EndsWith("}}") Then
                Dim keyPath = s.Substring(2, s.Length - 4).Trim()

                If keyPath.StartsWith("env.", StringComparison.OrdinalIgnoreCase) Then
                    Dim envName = keyPath.Substring(4)
                    If String.Equals(envName, "DESKTOP", StringComparison.OrdinalIgnoreCase) Then
                        Return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    End If
                    Dim ev = Environment.GetEnvironmentVariable(envName)
                    If Not String.IsNullOrEmpty(ev) Then Return ev
                    Return Nothing
                End If

                If keyPath.Equals("base_url", System.StringComparison.OrdinalIgnoreCase) AndAlso Not System.String.IsNullOrEmpty(_baseUrl) Then
                    Return _baseUrl
                End If
                Dim parts = keyPath.Split("."c)
                If parts.Length = 0 Then Return Nothing
                Dim startIndex As System.Int32 = 0
                If parts(0).Equals("variables", System.StringComparison.OrdinalIgnoreCase) Then
                    startIndex = 1
                End If
                If startIndex >= parts.Length Then
                    Return Nothing
                End If
                Dim current As System.Object = Nothing
                Dim rootName = parts(startIndex)
                If Not _vars.TryGetValue(rootName, current) Then
                    If _vars.TryGetValue(keyPath, current) Then
                        Return current
                    End If
                    Return Nothing
                End If
                For i = startIndex + 1 To parts.Length - 1
                    If current Is Nothing Then Return Nothing
                    Dim segment = parts(i)

                    Dim jt = TryCast(current, Newtonsoft.Json.Linq.JToken)
                    If jt IsNot Nothing Then
                        Dim sel = jt(segment)
                        If sel Is Nothing Then Return Nothing
                        If TypeOf sel Is Newtonsoft.Json.Linq.JValue Then
                            current = DirectCast(sel, Newtonsoft.Json.Linq.JValue).Value
                        Else
                            current = sel
                        End If
                        Continue For
                    End If

                    Dim dict = TryCast(current, System.Collections.Generic.IDictionary(Of System.String, System.Object))
                    If dict IsNot Nothing Then
                        If Not dict.TryGetValue(segment, current) Then Return Nothing
                        Continue For
                    End If

                    Dim t = current.GetType()
                    Dim pi = t.GetProperty(segment,
                                       System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.IgnoreCase)
                    If pi IsNot Nothing Then
                        current = pi.GetValue(current)
                        Continue For
                    End If

                    Return Nothing
                Next
                Return current
            Else
                Return ExpandTemplates(s)
            End If
        End Function

        Private Function _IsTruthy(val As Object) As Boolean
            If val Is Nothing Then Return False
            If TypeOf val Is Boolean Then Return DirectCast(val, Boolean)

            If TypeOf val Is String Then
                Dim s = DirectCast(val, String).Trim()
                If s.Length = 0 Then Return False
                Select Case s.ToLowerInvariant()
                    Case "false", "0", "null", "none", "nil" : Return False
                End Select
                Return True
            End If

            If TypeOf val Is IEnumerable AndAlso Not TypeOf val Is String Then
                Dim en = DirectCast(val, IEnumerable).GetEnumerator()
                Try
                    Return en.MoveNext()
                Finally
                    Dim disp = TryCast(en, IDisposable)
                    If disp IsNot Nothing Then disp.Dispose()
                End Try
            End If
            Return True
        End Function

        Private Function SimpleMustacheRender(template As String, context As Object) As String
            If String.IsNullOrEmpty(template) Then Return String.Empty

            Dim ctxToken As JToken = Nothing
            If TypeOf context Is JToken Then
                ctxToken = DirectCast(context, JToken)
            ElseIf context IsNot Nothing Then
                Try : ctxToken = JToken.FromObject(context) : Catch : End Try
            End If

            Dim ResolveVar As Func(Of String, Object) =
    Function(path As String) As Object
        Dim key = path.Trim()
        If ctxToken IsNot Nothing Then
            Try
                Dim t = ctxToken.SelectToken(key)
                If t IsNot Nothing Then
                    Select Case t.Type
                        Case JTokenType.Array, JTokenType.Object
                            Return t
                        Case Else
                            Return t.ToString()
                    End Select
                End If
            Catch
            End Try
        End If

        If context IsNot Nothing Then
            Try
                Dim pi = context.GetType().GetProperty(key, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)
                If pi IsNot Nothing Then Return pi.GetValue(context)
            Catch
            End Try
        End If
        Return Nothing
    End Function

            ' Process simple (# / ^) sections iteratively (limited depth)
            Dim sectionRegex As New Regex("\{\{(?<sig>[#^])\s*(?<name>[^\}]+?)\s*\}\}(?<body>.*?)\{\{/\s*\k<name>\s*\}\}",
                              RegexOptions.Singleline)
            Dim loopGuard = 0
            Dim output = template
            While loopGuard < 50
                loopGuard += 1
                Dim m = sectionRegex.Match(output)
                If Not m.Success Then Exit While

                Dim kind = m.Groups("sig").Value
                Dim name = m.Groups("name").Value
                Dim body = m.Groups("body").Value
                Dim val = ResolveVar(name)
                Dim truthy = _IsTruthy(val)
                Dim repl As String = String.Empty

                If kind = "#" Then
                    If truthy Then
                        If TypeOf val Is JArray Then
                            Dim sb As New StringBuilder()
                            For Each item As JToken In DirectCast(val, JArray)
                                sb.Append(SimpleMustacheRender(body, item))
                            Next
                            repl = sb.ToString()
                        ElseIf TypeOf val Is IEnumerable AndAlso Not TypeOf val Is String Then
                            Dim sb As New StringBuilder()
                            For Each item In DirectCast(val, IEnumerable)
                                Dim tok As JToken = TryCast(item, JToken)
                                If tok Is Nothing AndAlso item IsNot Nothing Then
                                    Try : tok = JToken.FromObject(item) : Catch : End Try
                                End If
                                sb.Append(SimpleMustacheRender(body, If(tok, item)))
                            Next
                            repl = sb.ToString()
                        Else
                            repl = SimpleMustacheRender(body, context)
                        End If
                    End If
                Else
                    ' inverted ^
                    If Not truthy Then
                        repl = SimpleMustacheRender(body, context)
                    End If
                End If

                output = output.Substring(0, m.Index) & repl & output.Substring(m.Index + m.Length)
            End While

            ' Triple mustache {{{var}}} (raw – wenn nicht im lokalen Kontext, Placeholder stehen lassen)
            output = Regex.Replace(
                output,
                "\{\{\{\s*([^\}]+?)\s*\}\}\}",
                Function(mt As Match) As String
                    Dim name = mt.Groups(1).Value.Trim()
                    Dim v = ResolveVar(name)
                    If v Is Nothing Then
                        ' Für den zweiten Pass in normale Form überführen:
                        Return "{{" & name & "}}"
                    End If
                    Return v.ToString()
                End Function)

            ' Normale Variablen {{var}} – bei Nichtauflösung Platzhalter erhalten
            output = Regex.Replace(
                output,
                "\{\{\s*([^\}#/\^][^\}]*)\s*\}\}",
                Function(mt As Match) As String
                    Dim rawName = mt.Groups(1).Value
                    Dim name = rawName.Trim()
                    Dim v = ResolveVar(name)
                    If v Is Nothing Then
                        ' Unverändert (bereinigt) stehen lassen
                        Return "{{" & name & "}}"
                    End If
                    Return v.ToString()
                End Function)

            Return output
        End Function

        Private Function MaterializeContextPlaceholders(ctxToken As Newtonsoft.Json.Linq.JToken) As System.Object
            If ctxToken Is Nothing Then Return Nothing
            If ctxToken.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then
                Return ctxToken.ToObject(Of System.Object)()
            End If
            Dim jobj = DirectCast(ctxToken, Newtonsoft.Json.Linq.JObject)
            For Each prop In jobj.Properties().ToList()
                If prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    Dim s = prop.Value.ToString().Trim()
                    If s.StartsWith("{{") AndAlso s.EndsWith("}}") Then
                        Dim resolved = ResolveValue(s)
                        If resolved IsNot Nothing Then
                            Dim jt = TryCast(resolved, Newtonsoft.Json.Linq.JToken)
                            If jt IsNot Nothing Then
                                prop.Value = jt
                            Else
                                prop.Value = Newtonsoft.Json.Linq.JToken.FromObject(resolved)
                            End If
                        End If
                    End If
                End If
            Next
            Return jobj.ToObject(Of System.Object)()
        End Function

        Private Function ResolveContextValue(ctx As System.Object, path As System.String) As System.Object
            If ctx Is Nothing Then Return Nothing
            Dim jt = TryCast(ctx, Newtonsoft.Json.Linq.JToken)
            If jt IsNot Nothing Then
                Dim sel = jt.SelectToken(path)
                If sel Is Nothing Then Return Nothing
                If TypeOf sel Is Newtonsoft.Json.Linq.JValue Then
                    Return DirectCast(sel, Newtonsoft.Json.Linq.JValue).Value
                ElseIf sel.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    Return sel.ToString()
                Else
                    Return sel.ToObject(Of System.Object)()
                End If
            End If
            Dim dict = TryCast(ctx, System.Collections.Generic.IDictionary(Of System.String, System.Object))
            If dict IsNot Nothing Then
                If dict.ContainsKey(path) Then Return dict(path)
                Dim parts = path.Split("."c)
                Dim cur As System.Object = dict
                For Each p In parts
                    Dim d = TryCast(cur, System.Collections.Generic.IDictionary(Of System.String, System.Object))
                    If d IsNot Nothing AndAlso d.ContainsKey(p) Then
                        cur = d(p)
                    Else
                        Return Nothing
                    End If
                Next
                Return cur
            End If
            Dim t = ctx.GetType()
            Dim prop = t.GetProperty(path)
            If prop IsNot Nothing Then
                Return prop.GetValue(ctx)
            End If
            Return Nothing
        End Function

        Private Function EvalCondition(condition As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(condition) Then Return False
            Dim c = condition.Trim()

            ' OR support
            If c.Contains("||") Then
                For Each part In c.Split(New String() {"||"}, StringSplitOptions.RemoveEmptyEntries)
                    If EvalCondition(part.Trim()) Then Return True
                Next
                Return False
            End If

            ' Empty array equality: {{var}} == []
            Dim emptyArr = Regex.Match(c, "^\s*({{.*}})\s*==\s*\[\s*\]\s*$")
            If emptyArr.Success Then
                Dim left = ResolveValue(emptyArr.Groups(1).Value)
                If left Is Nothing Then Return True
                If TypeOf left Is IEnumerable AndAlso Not TypeOf left Is String Then
                    Dim en = DirectCast(left, IEnumerable).GetEnumerator()
                    Dim any As Boolean = en.MoveNext()
                    Dim disp = TryCast(en, IDisposable) : If disp IsNot Nothing Then disp.Dispose()
                    Return Not any
                End If
                If TypeOf left Is String Then Return String.IsNullOrWhiteSpace(DirectCast(left, String))
                Return False
            End If

            Dim ex = System.Text.RegularExpressions.Regex.Match(c, "^\s*exists\s+({{.*}})\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If ex.Success Then
                Dim v = ResolveValue(ex.Groups(1).Value)
                Return v IsNot Nothing AndAlso Not System.String.IsNullOrEmpty(v.ToString())
            End If

            Dim eq = System.Text.RegularExpressions.Regex.Match(c, "^\s*({{.*}})\s*==\s*""?(.*?)""?\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If eq.Success Then
                Dim left = ResolveValue(eq.Groups(1).Value)
                Dim right = eq.Groups(2).Value
                Return System.String.Equals(If(left?.ToString(), System.String.Empty), right, System.StringComparison.OrdinalIgnoreCase)
            End If

            Dim co = System.Text.RegularExpressions.Regex.Match(c, "^\s*({{.*}})\s*contains\s*""(.*?)""\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If co.Success Then
                Dim left = ResolveValue(co.Groups(1).Value)?.ToString()
                Dim subStr = co.Groups(2).Value
                If left Is Nothing Then Return False
                Return left.IndexOf(subStr, System.StringComparison.OrdinalIgnoreCase) >= 0
            End If

            Dim rx = System.Text.RegularExpressions.Regex.Match(c, "^\s*({{.*}})\s*~=\s*""(.*?)""\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
            If rx.Success Then
                Dim left = ResolveValue(rx.Groups(1).Value)?.ToString()
                Dim pat = rx.Groups(2).Value
                If left Is Nothing Then Return False
                Return System.Text.RegularExpressions.Regex.IsMatch(left, pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
            End If

            Return False
        End Function

#End Region

    End Class

End Namespace