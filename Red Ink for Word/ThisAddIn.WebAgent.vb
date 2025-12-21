' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebAgent.vb
' Purpose: Implements WebAgent script selection, execution, and authoring capabilities
'          for automated web interaction workflows driven by JSON scripts.
'
' Architecture:
'  - Script Discovery: Scans local and global directories (INI_WebAgentPathLocal, INI_WebAgentPath)
'    for script files matching pattern '{AN2}-ag-*.json'. Local scripts take precedence.
'  - Parameter & Secret Substitution: Pre-execution processing via ProcessParameterPlaceholders
'    (handled by SharedLibrary) and ProcessWebAgentSecretPlaceholders (prompts user for secret values
'    matching {{secret:Description; type}} pattern in env.secrets).
'  - Pre-flight Validation: JSON parsing, confirmation dialogs for email send steps (ConfirmEmailSendSteps),
'    relative URL warnings when base_url missing.
'  - Interpreter Execution: WebAgentInterpreter.RunAsync executes script steps sequentially with support
'    for retries, error handling, conditional branching, loops, LLM calls, HTTP operations, file I/O,
'    templating, and email reporting.
'  - Cancellation: ESC key monitoring via PollEscForCancel runs in background thread; cancels CancellationTokenSource
'    on key press (debounced to prevent accidental triggers).
'  - Output Handling: Final Markdown report displayed in custom dialog; user can insert into document,
'    transfer to pane, or copy to clipboard (with optional Markdown-to-RTF conversion via InsertTextWithMarkdown).
'  - Script Authoring (CreateModifyWebAgentScript): LLM-assisted creation or amendment of scripts.
'    Provides WebAgentJSONInstruct and WebAgentParameterSpec as system prompts. Supports JSON validation
'    with optional auto-correction loop. Saves to configured directories with .bak backup on overwrite.
'  - Model Selection: Optional secondary API/model usage via INI_SecondAPI and INI_AlternateModelPath.
'  - External Dependencies: WebAgentInterpreter class (interpreter logic), SharedLibrary.SharedMethods
'    (UI dialogs, clipboard, text editor, model selection), Newtonsoft.Json (JSON parsing), Markdig
'    (Markdown rendering for email HTML conversion).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Data
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Discovers and prompts user to select a WebAgent script file from configured local/global directories.
    ''' Local scripts (INI_WebAgentPathLocal) are listed first, followed by global (INI_WebAgentPath).
    ''' </summary>
    ''' <returns>Full path to selected script file, or Nothing if canceled or no scripts found.</returns>
    Private Function GetWebAgentScriptFile() As String
        Dim globalPath As String = ExpandEnvironmentVariables(INI_WebAgentPath)
        Dim localPath As String = ExpandEnvironmentVariables(INI_WebAgentPathLocal)

        Dim candidates As New List(Of (Display As String, FullPath As String, IsLocal As Boolean))

        If String.IsNullOrWhiteSpace(localPath) AndAlso String.IsNullOrWhiteSpace(globalPath) Then
            ShowCustomMessageBox("No WebAgent paths configured. Please set 'WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file.", $"{AN} WebAgent")
            Return Nothing
        End If

        Try
            If Not String.IsNullOrWhiteSpace(localPath) AndAlso IO.Directory.Exists(localPath) Then
                For Each f In IO.Directory.GetFiles(localPath, $"{AN2}-ag-*.json", IO.SearchOption.TopDirectoryOnly)
                    candidates.Add(($"{IO.Path.GetFileName(f)} (local)", f, True))
                Next
            End If
        Catch ex As Exception
            ShowCustomMessageBox("Error enumerating local WebAgent scripts: " & ex.Message, $"{AN} WebAgent")
        End Try

        Try
            If Not String.IsNullOrWhiteSpace(globalPath) AndAlso IO.Directory.Exists(globalPath) Then
                For Each f In IO.Directory.GetFiles(globalPath, $"{AN2}-ag-*.json", IO.SearchOption.TopDirectoryOnly)
                    candidates.Add(($"{IO.Path.GetFileName(f)}", f, False))
                Next
            End If
        Catch ex As Exception
            ShowCustomMessageBox("Error enumerating global WebAgent scripts: " & ex.Message, $"{AN} WebAgent")
        End Try

        If candidates.Count = 0 Then
            ShowCustomMessageBox($"No WebAgent scripts found. Please set 'WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file and place files named '{AN2}-ag-*.json'.", $"{AN} WebAgent")
            Return Nothing
        End If

        ' Local first, then global; alphabetical within group
        Dim ordered = candidates.
            OrderByDescending(Function(c) c.IsLocal).
            ThenBy(Function(c) c.Display, StringComparer.OrdinalIgnoreCase).
            ToList()

        Dim displayItems = ordered.Select(Function(c) c.Display).ToArray()
        Dim sel As String = ShowSelectionForm("Select the WebAgent script you want to run:", $"{AN} WebAgent", displayItems)
        If String.IsNullOrEmpty(sel) Then Return Nothing

        Dim chosen = ordered.FirstOrDefault(Function(o) o.Display = sel)
        Return chosen.FullPath
    End Function


    ''' <summary>
    ''' Detects secret placeholders in script's env.secrets section (pattern: {{secret:Description; type}}).
    ''' Prompts user for secret values via ShowCustomVariableInputForm (values not logged).
    ''' Substitutes raw values into JSON tree before interpreter sees the script.
    ''' </summary>
    ''' <param name="script">JSON script string passed by reference; mutated in place with secret values.</param>
    ''' <returns>True if successful or no secrets found; False if user cancels or processing fails.</returns>
    Private Function ProcessWebAgentSecretPlaceholders(ByRef script As String) As Boolean
        Try
            Dim root As JObject = Nothing
            Try
                root = JObject.Parse(script)
            Catch
                ' Let normal JSON error handling later handle issues
                Return True
            End Try
            Dim envObj = TryCast(root("env"), JObject)
            If envObj Is Nothing Then Return True
            Dim secretsObj = TryCast(envObj("secrets"), JObject)
            If secretsObj Is Nothing OrElse secretsObj.Count = 0 Then Return True

            Dim secretDefs As New List(Of (Name As String, Placeholder As String, Desc As String))()

            Dim rx As New Regex("^\{\{\s*secret\s*:(.+?)\}\}$", RegexOptions.IgnoreCase)
            For Each p In secretsObj.Properties()
                Dim v = p.Value.ToString().Trim()
                Dim m = rx.Match(v)
                If m.Success Then
                    Dim meta = m.Groups(1).Value.Trim()
                    ' Metadata format supports "Description; type" segments; only description is used for prompts
                    Dim segs = meta.Split(";"c).Select(Function(s) s.Trim()).Where(Function(s) s <> "").ToArray()
                    Dim desc = segs(0)
                    secretDefs.Add((p.Name, v, desc))
                End If
            Next
            If secretDefs.Count = 0 Then Return True

            Dim ipList As New List(Of SharedLibrary.SharedLibrary.SharedMethods.InputParameter)
            For Each sd In secretDefs
                ipList.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter($"Secret: {sd.Name} - {sd.Desc}", ""))
            Next

            Dim arr = ipList.ToArray()
            If Not ShowCustomVariableInputForm("Provide secret values (they are not logged):", "WebAgent Secrets", arr) Then
                Return False
            End If

            ' Substitute into JSON tree (not by raw string replace to avoid accidental collisions)
            For i = 0 To secretDefs.Count - 1
                Dim name = secretDefs(i).Name
                Dim userVal = CStr(arr(i).Value)
                secretsObj(name) = userVal
            Next

            script = root.ToString(Formatting.None)
            Return True
        Catch ex As Exception
            ShowCustomMessageBox("Secret placeholder processing failed: " & ex.Message, AN & " WebAgent")
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Cancellation token source for the currently running WebAgent execution.
    ''' </summary>
    Private _webAgentRunCts As CancellationTokenSource

    ''' <summary>
    ''' Cancels the currently running WebAgent execution by triggering the cancellation token.
    ''' Safe to call even if no run is active.
    ''' </summary>
    Public Sub CancelWebAgentRun()
        Try : _webAgentRunCts?.Cancel() : Catch : End Try
    End Sub

    ''' <summary>
    ''' Background polling thread that monitors ESC key state via GetAsyncKeyState.
    ''' Cancels provided CancellationTokenSource when ESC pressed (debounced to require key release).
    ''' Runs until cancellation token signaled or ESC pressed.
    ''' </summary>
    ''' <param name="cts">CancellationTokenSource to cancel when ESC detected.</param>
    Private Sub PollEscForCancel(cts As CancellationTokenSource)
        Try
            Dim escLatched As Boolean = False
            While Not cts.IsCancellationRequested
                Thread.Sleep(150)
                Dim state = GetAsyncKeyState(VK_ESCAPE)
                If (state And &H8000) <> 0 Then
                    ' Debounce: ensure key released once before re-trigger
                    If Not escLatched Then
                        escLatched = True
                        cts.Cancel()
                        Exit While
                    End If
                Else
                    escLatched = False
                End If
            End While
        Catch
            ' Ignore polling errors
        End Try
    End Sub

    ''' <summary>
    ''' Scans script steps for 'send_email_report' commands and displays confirmation dialog
    ''' listing recipients (to, subject, smtp_host) before execution proceeds.
    ''' Security gate to prevent unintentional email sends.
    ''' </summary>
    ''' <param name="root">Parsed JSON root object of the script.</param>
    ''' <returns>1 if user confirms or no email steps found; 0 if user cancels or error occurs.</returns>
    Private Function ConfirmEmailSendSteps(root As JObject) As Integer
        Try
            Dim stepsArr = TryCast(root("steps"), JArray)
            If stepsArr Is Nothing OrElse stepsArr.Count = 0 Then Return 1

            Dim emailSteps As New List(Of JObject)
            For Each st As JObject In stepsArr
                Dim cmd = st.Value(Of String)("command")
                If Not String.IsNullOrWhiteSpace(cmd) AndAlso
                   String.Equals(cmd, "send_email_report", StringComparison.OrdinalIgnoreCase) Then
                    emailSteps.Add(st)
                End If
            Next
            ' No email steps found
            If emailSteps.Count = 0 Then Return 1

            ' Collect recipient info and optionally smtp_host to display
            Dim lines As New List(Of String)
            For Each st In emailSteps
                Dim p = TryCast(st("params"), JObject)
                Dim toRaw As String = p?("to")?.ToString()
                Dim smtpHost As String = p?("smtp_host")?.ToString()
                Dim subject As String = p?("subject")?.ToString()

                Dim recipients As String = ""
                If Not String.IsNullOrWhiteSpace(toRaw) Then
                    ' Normalize delimiters and trim each recipient
                    Dim parts = toRaw.Split(New String() {",", ";"}, StringSplitOptions.RemoveEmptyEntries).
                                      Select(Function(s) s.Trim()).
                                      Where(Function(s) s <> "")
                    recipients = String.Join(", ", parts)
                Else
                    recipients = "(missing)"
                End If

                Dim meta As New List(Of String)
                If Not String.IsNullOrWhiteSpace(subject) Then meta.Add($"Subject: {subject}")
                If Not String.IsNullOrWhiteSpace(smtpHost) Then meta.Add($"SMTP: {smtpHost}")

                If meta.Count > 0 Then
                    lines.Add($"- To: {recipients}  [{String.Join(" | ", meta)}]")
                Else
                    lines.Add($"- To: {recipients}")
                End If
            Next

            Dim msg As String =
                "Security check: This WebAgent script contains step(s) that will send an e-mail." & vbCrLf & vbCrLf &
                "Recipients detected:" & vbCrLf &
                String.Join(vbCrLf, lines) & vbCrLf & vbCrLf &
                "Do you wish to continue?"

            ' Ask for confirmation
            Dim ok As Integer = ShowCustomYesNoBox(msg, "Yes", "No, abort")
            Return ok
        Catch ex As Exception
            ' On any unexpected error, be safe and require explicit approval by failing the check
            ShowCustomMessageBox("Email send confirmation failed: " & ex.Message, $"{AN} WebAgent")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Main entry point for WebAgent script execution workflow.
    ''' Sequence: Script selection → Parameter/secret substitution → JSON validation → Pre-flight checks
    ''' (email confirmation, relative URL warnings) → Interpreter execution with ESC cancellation monitoring
    ''' → Output display with clipboard/insert/pane options.
    ''' Handles errors and provides user feedback at each stage.
    ''' </summary>
    Public Async Sub WebAgent()

        If INILoadFail() Then Return

        ' 1) Let user pick a script file (local first) 
        Dim selectedFile As String = GetWebAgentScriptFile()
        If String.IsNullOrWhiteSpace(selectedFile) Then Exit Sub

        Dim script As String = Nothing
        Try
            script = IO.File.ReadAllText(selectedFile, System.Text.Encoding.UTF8)
        Catch ex As Exception
            ShowCustomMessageBox("Failed to read script file:" & vbCrLf & selectedFile & vbCrLf & ex.Message, $"{AN} WebAgent")
            Exit Sub
        End Try

        If String.IsNullOrWhiteSpace(script) Then
            ShowCustomMessageBox("Script file is empty: " & selectedFile, $"{AN} WebAgent")
            Exit Sub
        End If

        ' 2) Detect and process parameter placeholders {{parameterN = ...}}
        If Not SLib.ProcessParameterPlaceholders(script) Then
            ShowCustomMessageBox("WebAgent canceled (no parameters applied).", $"{AN} WebAgent")
            Exit Sub
        End If

        If Not ProcessWebAgentSecretPlaceholders(script) Then
            ShowCustomMessageBox("WebAgent canceled (secrets).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 3) Parse (after substitution) to validate and for pre-flight checks
        Dim rootObj As JObject = Nothing
        Try
            rootObj = JObject.Parse(script)
        Catch ex As Exception
            SLib.PutInClipboard(ex.Message)
            ShowCustomMessageBox("Your script contains an incorrect JSON string (error copied to clipboard):" & vbCrLf & ex.Message, $"{AN} WebAgent")
            Exit Sub
        End Try

        Dim stepsArr = TryCast(rootObj("steps"), JArray)
        If stepsArr Is Nothing OrElse stepsArr.Count = 0 Then
            ShowCustomMessageBox("No steps found in script.", $"{AN} WebAgent")
            Exit Sub
        End If

        If ConfirmEmailSendSteps(rootObj) <> 1 Then
            ShowCustomMessageBox("WebAgent canceled (email send not approved).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 4) Pre-flight relative URL warnings
        Dim baseUrl As String = rootObj.SelectToken("env.base_url")?.ToString()
        Dim relativeIssues As New List(Of String)
        For Each st As JObject In stepsArr
            Dim cmd = st.Value(Of String)("command")
            If cmd = "open_url" OrElse cmd = "http_request" Then
                Dim p = TryCast(st("params"), JObject)
                Dim rawUrl = p?("url")?.ToString()
                If Not String.IsNullOrWhiteSpace(rawUrl) Then
                    If Not rawUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not rawUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                        If String.IsNullOrWhiteSpace(baseUrl) Then
                            relativeIssues.Add($"Step id='{st.Value(Of String)("id")}' uses relative or scheme-less URL '{rawUrl}' but env.base_url is not set.")
                        End If
                    End If
                Else
                    relativeIssues.Add($"Step id='{st.Value(Of String)("id")}' has empty or missing params.url.")
                End If
            End If
        Next
        If relativeIssues.Count > 0 Then
            Dim warnText As New StringBuilder()
            warnText.AppendLine("Pre-flight URL warnings:")
            For Each w In relativeIssues
                warnText.AppendLine(w)
            Next
            ShowCustomMessageBox(warnText.ToString().TrimEnd(), $"{AN} WebAgent", 20)
        End If

        ' 5) Run interpreter
        Dim finalMd As String = ""
        Dim webAgentCompleted As Boolean = False
        Dim abortedByEsc As Boolean = False

        _webAgentRunCts = New CancellationTokenSource()
        Dim ctsWeb = _webAgentRunCts

        ' Fully qualify Task to avoid ambiguity with Microsoft.Office.Interop.Word.Task
        Dim escMonitor As System.Threading.Tasks.Task =
            System.Threading.Tasks.Task.Run(Sub() PollEscForCancel(ctsWeb))

        Using interp As New WebAgentInterpreter()
            Dim sw As Stopwatch = Stopwatch.StartNew()
            Try
                finalMd = Await interp.RunAsync(
                    rootObj.ToString(Formatting.None),
                    _context,
                    True,   ' useSecondAPI
                    True,   ' autoselectModel
                    ctsWeb.Token)

                webAgentCompleted = True

            Catch ex As OperationCanceledException
                abortedByEsc = ctsWeb.IsCancellationRequested
                If abortedByEsc Then
                    finalMd = "# WebAgent aborted (ESC)" & vbCrLf & "User pressed ESC."
                Else
                    finalMd = "# WebAgent run canceled" & vbCrLf & ex.Message
                End If

            Catch ex As Exception
                Dim coreMsg = If(String.IsNullOrWhiteSpace(ex.Message), ex.GetType().Name, ex.Message)
                finalMd = "# WebAgent run failed" & vbCrLf & coreMsg
            End Try
            sw.Stop()
        End Using

        ' Stop ESC monitor task (fully qualified Task to suppress ambiguity)
        Try
            ctsWeb.Cancel()
            System.Threading.Tasks.Task.WaitAny(New System.Threading.Tasks.Task() {escMonitor}, 250)
        Catch
        End Try

        Dim isFailure As Boolean =
                    String.IsNullOrWhiteSpace(finalMd) OrElse
                    finalMd.StartsWith("# WebAgent run failed", StringComparison.OrdinalIgnoreCase) OrElse
                    finalMd.StartsWith("# WebAgent run canceled", StringComparison.OrdinalIgnoreCase) OrElse
                    finalMd.StartsWith("# WebAgent aborted", StringComparison.OrdinalIgnoreCase)

        If Not isFailure Then
            If finalMd.IndexOf("{{", StringComparison.Ordinal) >= 0 Then
                finalMd &= vbCrLf & vbCrLf & "_Warning: Unresolved template placeholders remain in output._"
            End If
        ElseIf String.IsNullOrWhiteSpace(finalMd) Then
            finalMd = "# WebAgent produced no output."
        End If

        InfoBox.ShowInfoBox("")

        Dim DialogResult As String = ShowCustomWindow("The WebAgent produced the following report:", finalMd, "You can edit the report. If you select OK, it will be copied to the clipboard in the original or edited form. You can also have the original inserted or transferred to the pane.", $"{AN} WebAgent", False, False, True, True)

        If DialogResult <> "" AndAlso DialogResult <> "Pane" Then
            If DialogResult = "Markdown" Then
                Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                If NewDocChoice = 1 Then
                    Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                    Dim currentSelection As Word.Selection = newDoc.Application.Selection
                    currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    InsertTextWithMarkdown(currentSelection, finalMd, True, True)
                ElseIf NewDocChoice = 2 Then
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & finalMd, False)
                Else
                    ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                    SLib.PutInClipboard(finalMd)
                End If
            Else
                SLib.PutInClipboard(DialogResult)
            End If
        ElseIf DialogResult = "Pane" Then
            ShowPaneAsync(
                "The WebAgent produced the following report:",
                finalMd,
                "",
                AN,
                noRTF:=False,
                insertMarkdown:=True
            )
        End If
    End Sub

    ''' <summary>
    ''' Complete specification document for WebAgent JSON script format, provided to LLMs for reliable authoring.
    ''' Covers: top-level structure (meta/env/steps), variable system, condition syntax, command reference,
    ''' retry mechanics, error handling, dynamic expansion, auto link extraction, template rendering,
    ''' selector objects, LLM result sanitization, common pitfalls, and authoring guidelines.
    ''' Version: Generated from current WebAgentInterpreter implementation.
    ''' </summary>
    Public Const WebAgentJSONInstruct As String =
        "WEB AGENT SCRIPT SPECIFICATION (for WebAgentInterpreter) " & vbCrLf &
        "Version: Generated from current interpreter code. Provide this spec verbatim to an LLM so it can reliably author valid scripts." & vbCrLf &
        "==================================================================================================================" & vbCrLf &
        "1. TOP-LEVEL JSON STRUCTURE" & vbCrLf &
        "{""meta"":{...},""env"":{...},""steps"":[ {StepObject}, ... ] }" & vbCrLf &
        "All fields optional unless marked required." & vbCrLf &
        "" & vbCrLf &
        "meta:" & vbCrLf &
        "  default_timeout_ms : Int (ms for steps without own timeout_ms)" & vbCrLf &
        "  user_agent         : String (default 'WebAgentInterpreter/1.0')" & vbCrLf &
        "" & vbCrLf &
        "env:" & vbCrLf &
        "  base_url   : String (used for resolving relative URLs)" & vbCrLf &
        "  headers    : { headerName: value, ... } (applied globally first)" & vbCrLf &
        "  secrets    : { name: valueOrReference } (if value starts with 'secret://', interpreter tries lookup in internal _secrets; returns '' if missing)" & vbCrLf &
        "  variables  : { name: initialValue } (arbitrary JSON values become initial _vars entries)" & vbCrLf &
        "" & vbCrLf &
        "steps: Array of Step Objects executed sequentially (with possible jumps via on_error.goto or guard.else_goto)." & vbCrLf &
        "Each Step Object keys:" & vbCrLf &
        "  id          : String (recommended, used by jumps and reporting)" & vbCrLf &
        "  command     : String (one of COMMAND LIST below)" & vbCrLf &
        "  params      : Object (command-specific)" & vbCrLf &
        "  timeout_ms  : Int (overrides default_timeout_ms for this step)" & vbCrLf &
        "  retry       : { max:Int, delay_ms:Int, backoff:Double }  (exponential delay = delay_ms * backoff^(attemptIndex))" & vbCrLf &
        "  on_error    : { action:""continue|goto|abort|retry"", goto:""StepId"" } (if step permanently fails after retries)" & vbCrLf &
        "                action meanings:" & vbCrLf &
        "                  continue = swallow error, proceed to next step" & vbCrLf &
        "                  goto     = jump to specified step id (requires goto field)" & vbCrLf &
        "                  abort    = rethrow (halts script)" & vbCrLf &
        "                  retry    = (already exhausted normal retry) will rethrow (use only if you wrap externally)" & vbCrLf &
        "                if on_error absent and _failHard=False (default), failures do not stop unless thrown" & vbCrLf &
        "  assign      : { var: ""varName"", path:""jsonPath"" } (stores result; if path given, selects token by JToken.SelectToken)" & vbCrLf &
        "  guard       : { if:""ConditionExpr"", else_goto:""StepId"" } (skip this step if condition FALSE; optional branch jump)" & vbCrLf &
        "  wait_for    : { type:""time|url|selector"", ... } (post-step passive check/log – DOES NOT WAIT actively except type=time before step execution)" & vbCrLf &
        "               time     => { type:""time"", timeout_ms:Int } (pre-step delay)" & vbCrLf &
        "               url      => { type:""url"", value:""substring expected in last URL"" } (logs if not found)" & vbCrLf &
        "               selector => { type:""selector"", selector:{...same selector object...} } (logs if missing)" & vbCrLf &
        "" & vbCrLf &
        "2. VARIABLE SYSTEM / TEMPLATE EXPANSION" & vbCrLf &
        "Variables stored in internal _vars dictionary. Placeholders: {{varName}} or nested paths like {{object.prop}}." & vbCrLf &
        "Special resolution rules:" & vbCrLf &
        "  - {{env.DESKTOP}} returns user desktop path." & vbCrLf &
        "  - {{base_url}} returns env.base_url (if set)." & vbCrLf &
        "  - If placeholder unresolved it is left intact (still '{{...}}') and logged as unresolved." & vbCrLf &
        "  - Templates are expanded in most string params via ExpandTemplates BEFORE execution." & vbCrLf &
        "  - 'template' command performs Mustache-like render (supports #section, ^inverted, simple vars, triple {{{var}}}); then a second global variable expansion pass runs." & vbCrLf &
        "Truthiness (_IsTruthy) for conditions: False if null, empty string, 'false','0','null','none','nil' (case-insensitive); empty non-string IEnumerable => False; everything else True." & vbCrLf &
        "" & vbCrLf &
        "3. CONDITION SYNTAX (used in guard.if and if.params.condition)" & vbCrLf &
        "Supported forms (whitespace tolerant):" & vbCrLf &
        "  exists {{var}}" & vbCrLf &
        "  {{var}} == ""literal""            (string comparison, case-insensitive)" & vbCrLf &
        "  {{var}} contains ""substring""    (case-insensitive)" & vbCrLf &
        "  {{var}} ~= ""regex""              (Regex, Singleline + IgnoreCase)" & vbCrLf &
        "  {{var}} == []                     (True if var is null, empty string, or empty enumerable)" & vbCrLf &
        "Logical OR: expr1 || expr2 || expr3 (evaluated left→right, returns True on first True)." & vbCrLf &
        "No AND operator explicitly; emulate with nested guards or De Morgan via OR." & vbCrLf &
        "" & vbCrLf &
        "4. IMPLICIT / COMMON VARIABLES" & vbCrLf &
        "  base_url                (from env)" & vbCrLf &
        "  last_http_status        (Int) after open_url / http_request" & vbCrLf &
        "  last_http_elapsed_ms    (Int)" & vbCrLf &
        "  last_step_id            (String)" & vbCrLf &
        "  lastLlm                 (JObject or wrapper) parsed/sanitized LLM response (+ step/page metadata)" & vbCrLf &
        "  lastLlm_raw             (Raw LLM text)" & vbCrLf &
        "  lastLlm_page_url        (URL at time of LLM step)" & vbCrLf &
        "  lastLlm_latency_ms      (Int)" & vbCrLf &
        "  auto_links              (List<String>) extracted anchor URLs (auto extraction)" & vbCrLf &
        "  auto_link_patterns      (List<String>) optional patterns (set via set_var or AddAutoLinkPattern; defaults included)" & vbCrLf &
        "  auto_link_enable        (Boolean) enable/disable automatic link collection" & vbCrLf &
        "  auto_link_min           (Int) minimal href length (default 15)" & vbCrLf &
        "  all_decisions / decision_links (used if a custom step populates them; code supports accumulation pattern)" & vbCrLf &
        "" & vbCrLf &
        "5. SELECTOR OBJECT (used in extract_text / extract_html / extract_attribute / internal waits)" & vbCrLf &
        "{ ""strategy"": ""xpath|css|text|regex"", ""value"": ""selectorValue"", optional:" & vbCrLf &
        "  ""within"": {SelectorObject}  (scopes selection to nodes returned by nested selector)" & vbCrLf &
        "  ""relative"": { ""position"": ""first|last|nth"", ""nth"": Int }" & vbCrLf &
        "}" & vbCrLf &
        "Strategy behaviors:" & vbCrLf &
        "  xpath: value is raw XPath" & vbCrLf &
        "  css  : limited CSS -> XPath translator (supports tag, .class, #id, [attr=value], :nth-child(n), descendant ' ', '>' direct child)" & vbCrLf &
        "  text : value OR 'exact:literal' (case-insensitive substring unless 'exact:' prefix)" & vbCrLf &
        "  regex: regex matched against normalized inner text" & vbCrLf &
        "" & vbCrLf &
        "6. COMMAND REFERENCE" & vbCrLf &
        "Command names are case-insensitive; llm_analyze, llm, llmanalyze map to same implementation." & vbCrLf &
        "" & vbCrLf &
        "a) set_user_agent" & vbCrLf &
        "  params: { ""user_agent"": ""string"" }" & vbCrLf &
        "  result: { user_agent }" & vbCrLf &
        "" & vbCrLf &
        "b) set_headers" & vbCrLf &
        "  params: { ""mode"":""replace|merge"", ""headers"": { name:value,... } }" & vbCrLf &
        "  'replace' clears previous; 'merge'(default) adds/overwrites." & vbCrLf &
        "  result: { headers: { ...effectiveHeaders } }" & vbCrLf &
        "" & vbCrLf &
        "c) set_cookies" & vbCrLf &
        "  params: { ""cookies"": [ { name, value, domain, path, secure:Bool?, httpOnly:Bool? }, ... ] }" & vbCrLf &
        "  result: { count: Int }" & vbCrLf &
        "" & vbCrLf &
        "d) open_url" & vbCrLf &
        "  params: { url:String (required), method:String (default GET), headers:Object?, body:Any?, body_type:""json|form|raw"", return_body:Bool?, timeout_ms:Int?, retry:{max,delay_ms,backoff} }" & vbCrLf &
        "  Loads page, sets DOM, auto extracts links, sets lastResponseUrl/Body." & vbCrLf &
        "  result: { status, url, elapsed_ms, (body if return_body=True) }" & vbCrLf &
        "" & vbCrLf &
        "e) wait" & vbCrLf &
        "  params: { ms:Int }" & vbCrLf &
        "  result: { slept: ms }" & vbCrLf &
        "" & vbCrLf &
        "f) find" & vbCrLf &
        "  params: { ""in"":""varName"", ""text"":""needle"", ""assign"":{""var"":""destVar""} }" & vbCrLf &
        "  Case-insensitive substring search. result: { found:Boolean, index:Int|-1 }" & vbCrLf &
        "" & vbCrLf &
        "g) extract_text" & vbCrLf &
        "  params: { selector:{...}, all:Bool?(default False), normalize_whitespace:Bool?(default True), regex:String?, group:Int? }" & vbCrLf &
        "  If all=False ⇒ returns first matched text (after optional regex extraction). If all=True ⇒ returns List<String>." & vbCrLf &
        "" & vbCrLf &
        "h) extract_html" & vbCrLf &
        "  params: { selector:{...}, outer:Bool?(default False) }" & vbCrLf &
        "  Returns inner (default) or outer HTML of first match or """" if none." & vbCrLf &
        "" & vbCrLf &
        "i) extract_attribute" & vbCrLf &
        "  params: { nodes_var:""varHoldingSerializedNodes"", attribute:""attrName"", var:""targetVar"" }" & vbCrLf &
        "  Expects nodes_var to contain an enumerable of serialized node dicts (with 'attributes'). Collects attribute values list to targetVar." & vbCrLf &
        "  result: null (stores list)" & vbCrLf &
        "" & vbCrLf &
        "j) download_url" & vbCrLf &
        "  params: { url, target_dir, filename?, method?, headers?, body?, body_type? }" & vbCrLf &
        "  Writes file. result: { path, status }" & vbCrLf &
        "" & vbCrLf &
        "k) save_file" & vbCrLf &
        "  params: { path, content, encoding? }" & vbCrLf &
        "  encoding==""binary"" ⇒ content must be Base64; else UTF-8 text." & vbCrLf &
        "  result: { path }" & vbCrLf &
        "" & vbCrLf &
        "l) read_file" & vbCrLf &
        "  params: { path, encoding? }" & vbCrLf &
        "  encoding==""binary"" ⇒ returns Base64; else UTF-8 text." & vbCrLf &
        "" & vbCrLf &
        "m) http_request (generic, does NOT auto link-extract unless you parse body via subsequent open_url if needed)" & vbCrLf &
        "  params: { url, method?, headers?, query:{k:v}, body, body_type, timeout_ms? }" & vbCrLf &
        "  result: { status, headers:StringDump, body, url } (also sets DOM for body like open_url)" & vbCrLf &
        "" & vbCrLf &
        "n) set_var" & vbCrLf &
        "  params: { name, value } (value may be any JSON; strings get template-expanded)" & vbCrLf &
        "  result: { name, value }" & vbCrLf &
        "" & vbCrLf &
        "o) template" & vbCrLf &
        "  params: { template:String(Mustache-lite), context:Object }" & vbCrLf &
        "  Section rules (#name repeats over arrays; ^name for inverted). Triple {{{var}}} leaves raw (then second pass expansion)." & vbCrLf &
        "  result: rendered String" & vbCrLf &
        "" & vbCrLf &
        "p) if" & vbCrLf &
        "  params: { condition:String, steps:[StepObjects], else_steps:[StepObjects]? }" & vbCrLf &
        "  Executes branch inline (sub-steps share same variable scope)." & vbCrLf &
        "" & vbCrLf &
        "q) foreach" & vbCrLf &
        "  params: { list:""varNameHoldingArray"", item_var:""loopItemVar"", steps:[...], continue_on_error:Bool?, stop_on_error:Bool?, max_items:Int?, break_if_var_true:""varName"" }" & vbCrLf &
        "  Loop variables added: item_var, item_var_index." & vbCrLf &
        "  Error policy precedence: stop_on_error=True overrides continue_on_error." & vbCrLf &
        "  Optional break_if_var_true variable checked each iteration." & vbCrLf &
        "  result: { count:<iterationsSeen>, executed:<successfulBodies> }" & vbCrLf &
        "" & vbCrLf &
        "r) render_report" & vbCrLf &
        "  params: { engine:? (ignored currently), template:String, context:Object?, output_path:String? }" & vbCrLf &
        "  Produces Markdown; writes file if output_path (after expansion and ensuring no '{{' remain). Sets _finalMarkdown." & vbCrLf &
        "  result: { output: ""(memory)"" or filePath }" & vbCrLf &
        "" & vbCrLf &
        "s) delete_file" & vbCrLf &
        "  params: { path } result: Bool (true if deleted, false if absent/error)" & vbCrLf &
        "" & vbCrLf &
        "t) send_email_report" & vbCrLf &
        "  params (most template-expanded safely):" & vbCrLf &
        "    to (semicolon or comma separated)" & vbCrLf &
        "    subject?  body_markdown? (if not full HTML, converted via Markdig)" & vbCrLf &
        "    smtp_host, smtp_port?, smtp_ssl?(""true|false""), smtp_auth?, smtp_user?, smtp_pass?" & vbCrLf &
        "    from_email?, from_name?, ip_override?, helo_domain?, net? (network domain override)" & vbCrLf &
        "  Adds footer '(created using Red Ink WebAgent at <ip> from <domain>)' both HTML & plain." & vbCrLf &
        "  Sends multipart/alternative (text/plain + text/html)." & vbCrLf &
        "  result: Bool" & vbCrLf &
        "" & vbCrLf &
        "u) log" & vbCrLf &
        "  params: { level:String, message:String } (message expanded) → appends to internal log." & vbCrLf &
        "" & vbCrLf &
        "v) enable_dynamic" & vbCrLf &
        "  params: (ignored - may be sent same object as step) Enables dynamic fetch expansion (limited to MAX_DYNAMIC_FETCH=10)." & vbCrLf &
        "" & vbCrLf &
        "w) array_push" & vbCrLf &
        "  params: { array:""targetArrayVar"", item_var:""existingVar"" | item:<inlineValue> }" & vbCrLf &
        "  If array var absent or not JArray, creates new JArray. Appends deep-cloned token." & vbCrLf &
        "  result: { pushed:Boolean, count:Int, array:String }" & vbCrLf &
        "" & vbCrLf &
        "x) llm / llm_analyze / llmanalyze" & vbCrLf &
        "  params accepted synonyms:" & vbCrLf &
        "    system | systemPrompt   (system role)" & vbCrLf &
        "    user | prompt | input | arguments (user prompt text)" & vbCrLf &
        "    temperature            (string or number; falls back to context INI)" & vbCrLf &
        "    timeoutMs              (ms for this call)" & vbCrLf &
        "    status_var             (if set & equals ""404"" and allow_llm_on_404 not true, step skipped)" & vbCrLf &
        "    inner_attempts:Int     (internal re-tries if non-JSON or invalid, default 1)" & vbCrLf &
        "    inner_delay_ms:Int     (delay between inner attempts)" & vbCrLf &
        "  Validation / gating flags:" & vbCrLf &
        "    retry_on_invalid:Bool       (throw if invalid so outer step retry triggers)" & vbCrLf &
        "    reject_if_empty:Bool        (empty sanitized output invalid)" & vbCrLf &
        "    reject_if_plaintext:Bool    (non-JSON invalid unless allow_non_json)" & vbCrLf &
        "    allow_non_json:Bool         (overrides reject_if_plaintext)" & vbCrLf &
        "    require_key:""k1,k2""        (all keys must exist unless require_key_all=False)" & vbCrLf &
        "    require_key_all:Bool        (default True)" & vbCrLf &
        "    require_array_key:""arr1,arr2"" (each must be JSON array)" & vbCrLf &
        "    require_min_items:Int       (applies to each require_array_key array)" & vbCrLf &
        "    log_raw:Bool                (dump raw output to debug file if debug enabled)" & vbCrLf &
        "    max_preview:Int             (UI preview length, default 250)" & vbCrLf &
        "  Output handling:" & vbCrLf &
        "    - Removes fenced code blocks, prefers JSON in them, else tries first JSON substring." & vbCrLf &
        "    - Wraps plaintext if needed with markers (_invalid reasons) if validation fails." & vbCrLf &
        "    - Stores structured result in lastLlm and metadata fields; raw text in lastLlm_raw." & vbCrLf &
        "" & vbCrLf &
        "7. RETRY MECHANICS" & vbCrLf &
        "At STEP level: 'retry' object controls entire step body incl. network call(s)." & vbCrLf &
        "At SUB-STEPS inside foreach/if: internal retry loops replicate same pattern (attempt logging)." & vbCrLf &
        "Backoff formula: effectiveDelay = delay_ms * (backoff ^ attemptIndex). attemptIndex starts at 0." & vbCrLf &
        "" & vbCrLf &
        "8. ERROR HANDLING" & vbCrLf &
        "Network transient statuses recognized (IsTransientStatus): 408,425,429,500,502,503,504 cause automatic retry if step retry configured." & vbCrLf &
        "on_error applied after all retries fail." & vbCrLf &
        "failHard flag FALSE (compiled constant) so unhandled non-network failures usually logged and may proceed unless thrown." & vbCrLf &
        "" & vbCrLf &
        "9. DYNAMIC EXPANSION" & vbCrLf &
        "enable_dynamic triggers post-load scanning of scripts & inline code for index_aza.php patterns and fetches up to MAX_DYNAMIC_FETCH (10) additional resources, appending their content to DOM HTML." & vbCrLf &
        "" & vbCrLf &
        "10. AUTO LINK EXTRACTION" & vbCrLf &
        "After open_url/http_request load, AutoExtractLinks() collects anchors (//a[@href]) whose href length >= auto_link_min and matches any regex in auto_link_patterns (default supplied). Stored as auto_links list." & vbCrLf &
        "Control flags via set_var: auto_link_enable (Bool), auto_link_patterns (List<String> or String), auto_link_min (Int)." & vbCrLf &
        "" & vbCrLf &
        "11. TEMPLATE SYSTEM (SimpleMustacheRender)" & vbCrLf &
        "Supported: {{var}}, {{{var}}}, sections: {{#name}}...{{/name}}, inverted {{^name}}...{{/name}}. No lambdas. Nested keys via JSON paths inside context or top-level context properties." & vbCrLf &
        "Unresolved vars are left as {{var}} for second pass global expansion." & vbCrLf &
        "" & vbCrLf &
        "12. ASSIGN PATH SEMANTICS" & vbCrLf &
        "If assign.path present, interpreter converts result object to JToken, then SelectToken(path). If token missing, stores null." & vbCrLf &
        "Use standard JSONPath-like JToken paths (e.g. 'data.items[0].title')." & vbCrLf &
        "" & vbCrLf &
        "13. FILE ENCODING NOTES" & vbCrLf &
        "save_file / read_file with encoding 'binary' use Base64 payloads." & vbCrLf &
        "download_url always writes raw bytes to file; does not auto base64 them." & vbCrLf &
        "" & vbCrLf &
        "14. URL RESOLUTION" & vbCrLf &
        "Relative URL precedence: lastResponseUrl (if any) → base_url → unchanged string." & vbCrLf &
        "Markdown-style [text](url) patterns sanitized; angle brackets <...> stripped if present." & vbCrLf &
        "" & vbCrLf &
        "15. LLM RESULT SANITIZATION LOGIC (SanitizeLlmResult)" & vbCrLf &
        "1) Extract fenced code block(s) ```lang ...```; if any block parses as JSON, first one returned." & vbCrLf &
        "2) Remove code fences if present; attempt full parse; else extract first balanced {...} or [...] substring; else strip stray backticks." & vbCrLf &
        "" & vbCrLf &
        "16. LOOP / FOREACH CAUTIONS" & vbCrLf &
        "Inside foreach, modifications to item_var can affect subsequent logic but original enumerable enumerated from snapshot at loop start." & vbCrLf &
        "break_if_var_true evaluated after iteration body." & vbCrLf &
        "" & vbCrLf &
        "17. DEBUG FLAGS (set via variables to True/False):" & vbCrLf &
        "  debug, debug_allAttempts, debug_substeps, debug_var_changes, debug_include_script, debug_summary, debug_clear_llm_state," & vbCrLf &
        "  allow_llm_on_404, llm_rethrow_all" & vbCrLf &
        "Enable by setting corresponding variable to 'true' (string) or Boolean True (set_var)." & vbCrLf &
        "" & vbCrLf &
        "18. EXAMPLE MINIMAL SCRIPT" & vbCrLf &
        "{""meta"":{""default_timeout_ms"":15000}," & vbCrLf &
        """env"":{""base_url"":""https://example.com"",""variables"":{""searchTerm"":""widgets""}}," & vbCrLf &
        """steps"":[ " & vbCrLf &
        " {""id"":""open"",""command"":""open_url"",""params"":{""url"":""/products?q={{searchTerm}}""}, ""retry"":{""max"":2,""delay_ms"":1000,""backoff"":2}}," & vbCrLf &
        " {""id"":""extract"",""command"":""extract_text"",""params"":{""selector"":{""strategy"":""css"",""value"":""h1""}}, ""assign"":{""var"":""pageTitle""}}," & vbCrLf &
        " {""id"":""decide"",""command"":""if"",""params"":{""condition"":""{{pageTitle}} contains \""Widget\"""",""steps"":[{""id"":""ok"",""command"":""log"",""params"":{""level"":""info"",""message"":""Title ok: {{pageTitle}}""}}],""else_steps"":[{""id"":""warn"",""command"":""log"",""params"":{""level"":""warn"",""message"":""Unexpected title: {{pageTitle}}""}}]}}," & vbCrLf &
        " {""id"":""llm"",""command"":""llm_analyze"",""params"":{""system"":""Summarize the title."",""user"":""Title: {{pageTitle}}"",""reject_if_empty"":true,""retry_on_invalid"":true,""require_key"":""summary""},""assign"":{""var"":""llmOut""}}," & vbCrLf &
        " {""id"":""report"",""command"":""render_report"",""params"":{""template"":""# Report\n\nTitle: {{pageTitle}}\n\nLLM: {{llmOut.summary}}"",""output_path"":""{{env.DESKTOP}}\\report.md""}} ]}" & vbCrLf &
        "" & vbCrLf &
        "19. COMMON PITFALLS & PREVENTION" & vbCrLf &
        "- Unresolved placeholders: Always ensure variables exist before referencing; logs show '[template] Unresolved placeholder'." & vbCrLf &
        "- Using extract_attribute without serialized nodes: Must have a variable containing a list of serialized node objects (normally produced by a custom step; not directly by extract_text/html)." & vbCrLf &
        "- Relying on AND logic: Use nested if or guard chains (no direct AND operator)." & vbCrLf &
        "- LLM JSON validation: Provide explicit require_key / require_array_key for deterministic retries." & vbCrLf &
        "- Guard skip logic: When guard condition false, step is skipped but considered successful; else_goto processed if provided." & vbCrLf &
        "- open_url vs http_request: open_url supports return_body flag; http_request always returns body. Both set DOM context." & vbCrLf &
        "- Binary file save/read: Must supply Base64 string when encoding=""binary""." & vbCrLf &
        "- foreach list missing: Loop silently returns {count:0,executed:0} (log warns) – ensure list variable exists." & vbCrLf &
        "" & vbCrLf &
        "20. RECOMMENDED LLM AUTHORING GUIDELINES" & vbCrLf &
        "- Always produce valid UTF-8 JSON with top-level object and 'steps' array." & vbCrLf &
        "- Include 'id' for every step (unique, short, no spaces)." & vbCrLf &
        "- For network steps requiring resilience, add retry:{max,delay_ms,backoff} AND (optionally) on_error to control flow." & vbCrLf &
        "- Use assign with path when extracting subset from complex results (e.g., assign.path:'data.items[0]')." & vbCrLf &
        "- For conditional branching prefer explicit if command rather than abusing guard for multi-step blocks." & vbCrLf &
        "- For loops, ensure the source list variable is known to exist (initialize via set_var or prior assign)." & vbCrLf &
        "- Provide require_key / require_array_key in llm_analyze to harden JSON outputs; combine with retry_on_invalid for reliability." & vbCrLf &
        "- Avoid inventing unsupported commands or fields (strict list above)." & vbCrLf &
        "- Keep URLs absolute or ensure base_url defined for relative forms." & vbCrLf &
        "- Use enable_dynamic only when needed (dynamic pages); avoids extra fetch overhead." & vbCrLf &
        "" & vbCrLf &
        "END OF SPEC" & vbCrLf

    Public Const WebAgentParameterSpec As String =
            "WEB AGENT PARAMETER PLACEHOLDER SPEC (Additive to Main Script Spec)" & vbCrLf &
            "Purpose: Enable runtime user parameterization inside a WebAgent script via numbered placeholders." & vbCrLf &
            "-------------------------------------------------------------------------------------------------" & vbCrLf &
            "OVERVIEW:" & vbCrLf &
            "- Parameter DEFINITIONS embed inline in the script: {parameterN=...definition...}" & vbCrLf &
            "- Parameter REFERENCES elsewhere: {parameterN}" & vbCrLf &
            "- Definitions are collected, a dialog prompts user for values (unless zero definitions)." & vbCrLf &
            "- After confirmation: definitions are removed and replaced by resolved (escaped) values; all matching {parameterN} references substituted with same value." & vbCrLf &
            "- If user cancels, processing returns False and caller should abort execution." & vbCrLf &
            "" & vbCrLf &
            "REGEX MATCHES:" & vbCrLf &
            "- Definition regex:  { parameter(\d+)= (.*?) }   (whitespace tolerant)" & vbCrLf &
            "- Reference regex:   { parameter(\d+) }" & vbCrLf &
            "" & vbCrLf &
            "UNIQUENESS & ORDER:" & vbCrLf &
            "- First definition for a given N wins; duplicates with same N are ignored." & vbCrLf &
            "- UI prompt order = ascending parameter number." & vbCrLf &
            "- Replacement inside script done from the end backward (reverse index) to preserve positions." & vbCrLf &
            "" & vbCrLf &
            "DEFINITION SYNTAX (semicolon-separated segments):" & vbCrLf &
            "  {parameterN=Description ; Type ; DefaultValue ; RangeOrOptions ; ExtraOptions }" & vbCrLf &
            "  Minimal form: {parameter1=Choose item} (Type defaults to string, default empty)" & vbCrLf &
            "" & vbCrLf &
            "SEGMENTS:" & vbCrLf &
            "  [0] Description (mandatory, shown in UI)" & vbCrLf &
            "  [1] Type (optional) → one of: string | integer | long | double | boolean (case-insensitive; default=string)" & vbCrLf &
            "  [2] Default value (optional; interpreted per type)" & vbCrLf &
            "  [3] EITHER a numeric range (only for integer/long/double) in form MIN-MAX (digits only) OR a comma list of options" & vbCrLf &
            "  [4] If [3] was a range and this segment exists → treated as additional options list (comma separated)" & vbCrLf &
            "" & vbCrLf &
            "OPTIONS SYNTAX:" & vbCrLf &
            "- Comma-separated: OptionA,OptionB,OptionC" & vbCrLf &
            "- Each option may embed a display/code pair:  Display Text <actual_code>" & vbCrLf &
            "  * If no <...> part present, display = code." & vbCrLf &
            "  * The UI shows display text; stored code value is inserted into script." & vbCrLf &
            "- Default value resolution: If DefaultValue matches a code entry, UI pre-selects corresponding display." & vbCrLf &
            "" & vbCrLf &
            "RANGE HANDLING (for numeric types):" & vbCrLf &
            "- Pattern: ^\\d+\\s*-\\s*\\d+$ (integers only). Extracted as inclusive [min,max]." & vbCrLf &
            "- User input (after mapping) clamped into range if parse succeeds." & vbCrLf &
            "- For integer/long: value rounded (Math.Round) then cast to integral." & vbCrLf &
            "- For double: stored as parsed (range still integer endpoints)." & vbCrLf &
            "" & vbCrLf &
            "TYPE PARSING & DEFAULTS:" & vbCrLf &
            "- boolean: Boolean.TryParse; output lowercased 'true'/'false'." & vbCrLf &
            "- integer/long: Integer/Long.TryParse; fallback 0 if invalid." & vbCrLf &
            "- double: Double.TryParse (culture invariant rules depend on environment; recommend dot decimal)." & vbCrLf &
            "- string: raw (after display→code mapping if options provided)." & vbCrLf &
            "" & vbCrLf &
            "SPECIAL / EMPTY SELECTION HANDLING:" & vbCrLf &
            "- If chosen value (case-insensitive) starts with '(keine auswahl)', '(no selection)', or '---' → final inserted value = empty string." & vbCrLf &
            "" & vbCrLf &
            "JSON ESCAPING:" & vbCrLf &
            "- Only backslash (\ → \\) and double quote ("" → \"") are escaped before insertion." & vbCrLf &
            "- No other JSON normalization (caller must ensure placement inside valid JSON string literal positions)." & vbCrLf &
            "" & vbCrLf &
            "REFERENCE REPLACEMENT:" & vbCrLf &
            "- Definitions are replaced in place (the entire {parameterN=...} token → final value)." & vbCrLf &
            "- Simple references {parameterN} replaced with same value if definition existed." & vbCrLf &
            "- References to undefined parameter numbers are left untouched." & vbCrLf &
            "" & vbCrLf &
            "BEHAVIOR WITH NO DEFINITIONS:" & vbCrLf &
            "- Function returns True immediately; no prompt; references remain unchanged (allowing static placeholders)." & vbCrLf &
            "" & vbCrLf &
            "VALID EXAMPLES:" & vbCrLf &
            "  {parameter1=API environment ; string ; prod ; prod<https://api.prod.example> , staging<https://api.staging.example> , dev<http://localhost:8080>} " & vbCrLf &
            "  {parameter2=Max retries ; integer ; 3 ; 0-10}" & vbCrLf &
            "  {parameter3=Confidence threshold ; double ; 0.65 ; 0-1 ; 0.25,0.5,0.65,0.75,0.9}" & vbCrLf &
            "  {parameter4=Enable verbose logging ; boolean ; true}" & vbCrLf &
            "  {parameter5=Output format ; string ; json ; json<application/json>,xml<application/xml>}" & vbCrLf &
            "" & vbCrLf &
            "UI PROMPT ORDER & CORRELATION:" & vbCrLf &
            "- User sees description; for enumerated selections a dropdown of display labels." & vbCrLf &
            "- After submit each selected/entered value is transformed (display→code, clamped, escaped) then substituted." & vbCrLf &
            "" & vbCrLf &
            "EDGE CASES / SAFEGUARDS:" & vbCrLf &
            "- Whitespace around semicolons ignored (segments trimmed)." & vbCrLf &
            "- Extra segments beyond defined semantics are ignored." & vbCrLf &
            "- If range present but user input not numeric, original string left (then possibly empty if sentinel)." & vbCrLf &
            "- Multiple identical options allowed but first matching code resolves default." & vbCrLf &
            "- Duplicate parameter definitions: only first retained; later ones inert text until replaced when earlier removal shifts indexes (safe due to reverse replacement)." & vbCrLf &
            "" & vbCrLf &
            "INTEGRATION INTO JSON SCRIPTS:" & vbCrLf &
            "- Place definitions where a literal value will finally appear (e.g. in place of a string literal's value body)." & vbCrLf &
            "- Example inside JSON (ensure resulting JSON remains valid *after* replacement):" & vbCrLf &
            "  ""url"": ""{parameter1=https://api.example.com ; string ; https://api.example.com}/v1/items""  (NOT RECOMMENDED because the definition will collapse into a raw URL → better put definition alone then follow with concatenation outside or define before and reference)" & vbCrLf &
            "- Recommended pattern: define at top (outside JSON structural tokens if pre-processed) or as value of a string field by itself: ""base_url"": ""{parameter1=Base URL ; string ; https://api.example.com}""" & vbCrLf &
            "" & vbCrLf &
            "RECOMMENDED AUTHORING RULES FOR LLM:" & vbCrLf &
            "- Use sequential numbering starting at 1; avoid gaps unless intentional." & vbCrLf &
            "- Use descriptive, concise descriptions (≤ 60 chars)." & vbCrLf &
            "- Provide defaults that yield a runnable script without user edits when possible." & vbCrLf &
            "- Use ranges only when numeric validation is concretely useful; otherwise supply enumerated options." & vbCrLf &
            "- Always ensure final substituted form keeps JSON valid (surround with quotes if expecting string)." & vbCrLf &
            "- Do NOT reference {parameterN} without defining it unless deliberately leaving literal marker." & vbCrLf &
            "" & vbCrLf &
            "NON-VALID / ANTI-PATTERNS:" & vbCrLf &
            "- {parameter1=}  (missing description)" & vbCrLf &
            "- {parameterX=...} where X not numeric" & vbCrLf &
            "- Embedding raw unescaped quotes in definition segments causing broken JSON context." & vbCrLf &
            "" & vbCrLf &
            "POST-PROCESSING OUTCOME SUMMARY:" & vbCrLf &
            "- Success → script mutated with raw values inserted, function True." & vbCrLf &
            "- Cancel → script unchanged from original input (except maybe partial test state), function False." & vbCrLf &
            "" & vbCrLf &
            "SECURITY CONSIDERATIONS:" & vbCrLf &
            "- Only minimal JSON escaping; if values may inject additional JSON structure, caller must sandbox or further sanitize." & vbCrLf &
            "- Angle-bracket code extraction is purely syntactic; no HTML interpretation." & vbCrLf &
            "" & vbCrLf &
            "QUICK TEMPLATE FOR LLM GENERATION:" & vbCrLf &
            "- Integer with range: {parameter1=Retry count ; integer ; 3 ; 0-10}" & vbCrLf &
            "- Double with options: {parameter2=Threshold ; double ; 0.75 ; 0.25,0.5,0.75,0.9}" & vbCrLf &
            "- Enum string: {parameter3=Mode ; string ; safe ; safe<safe>,fast<fast>,audit<audit>}" & vbCrLf &
            "- Boolean: {parameter4=Enable cache ; boolean ; false}" & vbCrLf &
            "" & vbCrLf &
            "END OF PARAMETER SPEC"

    ''' <summary>
    ''' LLM-assisted WebAgent script creation or amendment workflow.
    ''' Sequence: Directory validation → New/Amend mode selection → (if amend) Script editor for manual review
    ''' → User instruction prompt → Model selection (primary/secondary) → LLM generation with spec prompts
    ''' (WebAgentJSONInstruct + WebAgentParameterSpec) → JSON validation with optional auto-correction loop
    ''' → Save with .bak backup → Final editor display for manual parameter definition.
    ''' Supports overwrite confirmation and alternate model loading (INI_AlternateModelPath).
    ''' </summary>
    Public Async Sub CreateModifyWebAgentScript()

        If INILoadFail() Then Return

        ' 0) Preconditions: paths configured?
        Dim globalPath As String = ExpandEnvironmentVariables(INI_WebAgentPath)
        Dim localPath As String = ExpandEnvironmentVariables(INI_WebAgentPathLocal)

        If (String.IsNullOrWhiteSpace(globalPath) AndAlso String.IsNullOrWhiteSpace(localPath)) Then
            ShowCustomMessageBox("No WebAgent paths configured ('WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' Pick working directory (prefer local if available)
        Dim targetDir As String = If(Not String.IsNullOrWhiteSpace(localPath) AndAlso IO.Directory.Exists(localPath),
                                     localPath,
                                     If(Not String.IsNullOrWhiteSpace(globalPath) AndAlso IO.Directory.Exists(globalPath),
                                        globalPath,
                                        Nothing))
        If String.IsNullOrWhiteSpace(targetDir) Then
            ShowCustomMessageBox("Neither configured WebAgent directory exists.", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 1) New vs Amend
        Dim modeChoice = ShowCustomYesNoBox(
            "Do you want to create a new WebAgent script or amend an existing one?",
            "New",
            "Amend existing",
            $"{AN} WebAgent")
        If modeChoice = 0 Then Exit Sub

        Dim existingScriptPath As String = Nothing
        Dim originalScriptText As String = Nothing
        Dim isAmend As Boolean = (modeChoice = 2)

        If isAmend Then
            existingScriptPath = GetWebAgentScriptFile()
            If String.IsNullOrWhiteSpace(existingScriptPath) Then Exit Sub
            Try
                ' Read current script (initial content before possible user edits)
                originalScriptText = IO.File.ReadAllText(existingScriptPath, System.Text.Encoding.UTF8)

                ' Open in embedded text editor so user can review and edit
                ' Editor saves directly to the same path when user clicks Save
                SharedLibrary.SharedLibrary.SharedMethods.ShowTextFileEditor(existingScriptPath, $"{AN} WebAgent Script '{existingScriptPath}' (you can have it amended by the LLM once you close the editor):")

                ' Ask whether to proceed with automatic LLM amendment after editor closes
                Dim amendChoice As Integer = SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                    bodyText:="Do you want to automatically amend the script using the LLM now?",
                    button1Text:="Yes",
                    button2Text:="No",
                    header:=$"{AN} WebAgent"
                )

                ' ShowCustomYesNoBox returns 1 for first button ("Yes"), 2 for second ("No")
                If amendChoice <> 1 Then
                    Exit Sub
                End If

                ' Re-read file to capture any edits the user may have saved
                Try
                    originalScriptText = IO.File.ReadAllText(existingScriptPath, System.Text.Encoding.UTF8)
                Catch exReload As Exception
                    ShowCustomMessageBox("Failed to re-read edited script before amendment:" & vbCrLf & exReload.Message, $"{AN} WebAgent")
                    Exit Sub
                End Try

            Catch ex As Exception
                ShowCustomMessageBox("Unable to read existing script:" & vbCrLf & ex.Message, $"{AN} WebAgent")
                Exit Sub
            End Try
        End If

        ' 2) If new, gather filename with validation and overwrite confirmation
        Dim newScriptPath As String = existingScriptPath
        If Not isAmend Then
            Dim defaultName As String = $"{AN2}-ag-newscript.json"
            Dim rxName As New System.Text.RegularExpressions.Regex($"^{AN2}-ag-[a-z0-9._-]+\.json$", RegexOptions.IgnoreCase)
            Do
                Dim nameInput = ShowCustomInputBox($"Enter new script file name (pattern: {AN2}-ag-*.json)", $"{AN} WebAgent", True, defaultName)
                If String.IsNullOrWhiteSpace(nameInput) Then Exit Sub
                nameInput = nameInput.Trim()
                If Not rxName.IsMatch(nameInput) Then
                    ShowCustomMessageBox($"File name does not comply with pattern '{AN2}-ag-*.json'.", $"{AN} WebAgent")
                    Continue Do
                End If
                Dim fullCandidate = IO.Path.Combine(targetDir, nameInput)
                If IO.File.Exists(fullCandidate) Then
                    Dim ow = ShowCustomYesNoBox($"File '{nameInput}' already exists. Overwrite?", "Yes", "No", $"{AN} WebAgent")
                    If ow = 1 Then
                        newScriptPath = fullCandidate
                        Exit Do
                    ElseIf ow = 0 Then
                        Exit Sub
                    Else
                        Continue Do
                    End If
                Else
                    newScriptPath = fullCandidate
                    Exit Do
                End If
            Loop
        End If

        If String.IsNullOrWhiteSpace(newScriptPath) Then Exit Sub

        ' 3) User instructions
        Dim instruction = ShowCustomInputBox("You are about to have the LLM create or amend your script. For this purpose, describe what the WebAgent script shall do or how it shall be changed (note: user parameters need to be defined manually as per the user manual):", $"{AN} WebAgent", False, "")
        If String.IsNullOrWhiteSpace(instruction) Then Exit Sub
        OtherPrompt = instruction.Trim()

        If OtherPrompt = "" Then Return

        Dim useSecondAPI As Boolean = False

        Try

            ' 4) Primary vs Secondary model selection
            If INI_SecondAPI Then
                Dim modelChoice = ShowCustomYesNoBox("Select the model to generate or modify the script:", "Primary", "Secondary", $"{AN} WebAgent")
                If modelChoice = 0 Then Exit Sub
                If modelChoice = 2 Then
                    useSecondAPI = True

                    ' Attempt to autoload alternate model (best effort, non-fatal)
                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                        If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                            originalConfigLoaded = False
                            Return
                        End If
                    End If

                End If
            End If

            ' 5) Build user prompt (with optional existing script for amendment)
            Dim userPromptBuilder As New System.Text.StringBuilder()
            userPromptBuilder.Append("<USERPROMPT>").Append(OtherPrompt).Append("</USERPROMPT>")
            If isAmend AndAlso Not String.IsNullOrWhiteSpace(originalScriptText) Then
                userPromptBuilder.Append("<SCRIPT>").Append(originalScriptText).Append("</SCRIPT>")
            End If
            Dim userPromptFinal = userPromptBuilder.ToString()

            ' 6) Call LLM to generate or amend script
            Dim generated As String = Nothing
            Try
                generated = Await LLM(WebAgentJSONInstruct & " " & WebAgentParameterSpec, userPromptFinal, "", "", 0, useSecondAPI)
            Catch ex As Exception
                ShowCustomMessageBox("LLM call failed: " & ex.Message, $"{AN} WebAgent")
                Exit Sub
            End Try
            If String.IsNullOrWhiteSpace(generated) Then
                ShowCustomMessageBox("The LLM returned no content.", $"{AN} WebAgent")
                Exit Sub
            End If

            generated = WebAgentInterpreter.SanitizeLlmResult(generated)

            ' 7) Backup existing file before overwriting (only on first save, not during auto-correct loop)
            If IO.File.Exists(newScriptPath) AndAlso isAmend Then
                ' Amending existing file - make backup
                Try
                    SharedLibrary.SharedLibrary.SharedMethods.RenameFileToBak(newScriptPath)
                Catch
                End Try
            ElseIf IO.File.Exists(newScriptPath) AndAlso Not isAmend Then
                ' Overwriting a file in new mode - make backup
                Try
                    SharedLibrary.SharedLibrary.SharedMethods.RenameFileToBak(newScriptPath)
                Catch
                End Try
            End If

            ' 8) Save initial result
            Try
                IO.File.WriteAllText(newScriptPath, generated, System.Text.Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox("Failed to save generated script (script is in the clipboard instead):" & vbCrLf & ex.Message, $"{AN} WebAgent")
                PutInClipboard(generated)
                Exit Sub
            End Try

            ' 9) Validate JSON, with optional auto-correction loop
            Dim keepLoop As Boolean = True
            Dim firstTry As Boolean = True
            While keepLoop
                Dim parseOk As Boolean = True
                Dim parseError As String = Nothing
                Try
                    Newtonsoft.Json.Linq.JObject.Parse(IO.File.ReadAllText(newScriptPath, System.Text.Encoding.UTF8))
                Catch ex As Exception
                    parseOk = False
                    parseError = ex.Message
                End Try

                If parseOk Then
                    ' Validation successful
                    keepLoop = False
                    If firstTry Then
                        ShowCustomMessageBox("The JSON script created by the LLM has been tested for validity and was saved to:" & vbCrLf & vbCrLf & newScriptPath, $"{AN} WebAgent")
                    Else
                        ShowCustomMessageBox("The JSON script created by the LLM is now valid and was saved to:" & vbCrLf & vbCrLf & newScriptPath, $"{AN} WebAgent")
                    End If
                Else
                    ' Validation failed - offer auto-correction
                    Dim choice = ShowCustomYesNoBox($"The script JSON is invalid (it was saved anyway). Error:{vbCrLf}{parseError}{vbCrLf}{vbCrLf}Attempt auto-correction?", "Yes", "No", $"{AN} WebAgent")
                    If choice = 1 Then
                        ' Attempt auto-correction via LLM
                        Dim brokenJson = IO.File.ReadAllText(newScriptPath, System.Text.Encoding.UTF8)
                        Dim correctionSystem =
                        "You are a strict JSON repair assistant. Fix structural JSON errors ONLY. Preserve semantic content & keys. Output EXACTLY one corrected JSON object, no comments, no prose."
                        Dim correctionUser =
                        "<BROKEN_JSON>" & brokenJson & "</BROKEN_JSON><ERROR_MESSAGE>" & parseError & "</ERROR_MESSAGE>"
                        Dim corrected As String = Nothing
                        Try
                            corrected = Await LLM(correctionSystem, correctionUser, "", "", 0, useSecondAPI)
                        Catch ex As Exception
                            ShowCustomMessageBox("Auto-correction LLM call failed: " & ex.Message, $"{AN} WebAgent")
                            keepLoop = False
                            Exit While
                        End Try
                        If String.IsNullOrWhiteSpace(corrected) Then
                            ShowCustomMessageBox("Auto-correction returned empty content.", $"{AN} WebAgent")
                            keepLoop = False
                            Exit While
                        End If
                        Try
                            IO.File.WriteAllText(newScriptPath, corrected, System.Text.Encoding.UTF8)
                        Catch ex As Exception
                            ShowCustomMessageBox("Failed to save corrected script (script is in the clipboard instead):" & ex.Message, $"{AN} WebAgent")
                            PutInClipboard(corrected)
                            keepLoop = False
                            Exit While
                        End Try
                        ' Loop will re-validate
                    Else
                        keepLoop = False
                    End If
                    firstTry = False
                End If
            End While

        Catch
        Finally
            If useSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try

        ' 10) Show final file in editor for manual parameter definition
        Try
            ShowTextFileEditor(newScriptPath, $"{AN} WebAgent Script '{newScriptPath}'" & "(you can now add user parameters, e.g. '{{parameter1 = Date; String; 1.10.2025}}' - see user manual): ")
        Catch
        End Try

    End Sub


End Class
