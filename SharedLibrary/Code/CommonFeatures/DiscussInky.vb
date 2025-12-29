' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DiscussInky.vb
' Purpose: Implements a WinForms chat window ("Help me, AN8!") that lets the user
'          ask questions and displays LLM responses with Markdown rendering.
'
' Architecture:
'  - UI Layout: TableLayoutPanel hosts a `WebBrowser` chat view, a multi-line input
'    `TextBox`, and a `FlowLayoutPanel` with buttons and option checkboxes.
'  - Chat Rendering: Maintains HTML fragments in a queue until the `WebBrowser`
'    document is ready; appends fragments via small JS helpers.
'  - Persistence: Stores last chat as HTML (`My.Settings.LastHelpMeChatHtml`) and
'    as a capped plain transcript (`My.Settings.LastHelpMeChat`). Persists window
'    bounds and option flags.
'  - LLM Calls: Builds a system prompt from `_context.SP_HelpMe` and the current
'    conversation; optionally includes manual text loaded via a URL or local file.
'  - Manual Source: Loads a manual from URL/file path (supports HTML/plain/PDF/RTF/DOCX);
'    caches it per configured path for the lifetime of this form instance.
'  - Concurrency & Model Switching: Serializes LLM calls via `_modelSemaphore` and
'    optionally applies an alternate model configuration (using a second API key).
'  - Optional Config Inclusion: Adds sanitized INI/config content to the user prompt
'    when enabled; API keys are masked, comment lines are excluded.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Net
Imports System.Threading
Imports Markdig
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Net.Http
Imports System.Windows.Forms

Namespace SharedLibrary

    ''' <summary>
    ''' WinForms chat window that sends user questions to an LLM and renders responses as HTML/Markdown.
    ''' </summary>
    Public Class DiscussInky
        Inherits System.Windows.Forms.Form

        ''' <summary>
        ''' Title for the form window.
        ''' </summary>
        Private WindowTitle As String = $"Help me, {AN8}!"

        ''' <summary>
        ''' Display name used for assistant messages in the chat UI.
        ''' </summary>
        Private Const AssistantName As String = AN8

        ''' <summary>
        ''' Shared application context providing configuration and system prompts.
        ''' </summary>
        Private ReadOnly _context As ISharedContext

        ''' <summary>
        ''' Host application name (e.g., Word/Excel/Outlook) included in LLM context text.
        ''' </summary>
        Private ReadOnly _hostAppName As String

        ''' <summary>
        ''' Markdig pipeline used to render assistant Markdown into HTML.
        ''' </summary>
        Private ReadOnly _mdPipeline As Markdig.MarkdownPipeline

        ''' <summary>
        ''' WebBrowser used as a lightweight HTML chat transcript renderer.
        ''' </summary>
        Private ReadOnly _chat As WebBrowser = New WebBrowser() With {
        .Dock = DockStyle.Fill,
        .AllowWebBrowserDrop = False,
        .IsWebBrowserContextMenuEnabled = True,
        .WebBrowserShortcutsEnabled = True,
        .ScriptErrorsSuppressed = True
    }

        ''' <summary>
        ''' Multi-line textbox for user input.
        ''' </summary>
        Private ReadOnly _txtInput As TextBox = New TextBox() With {
        .Dock = DockStyle.Fill,
        .Multiline = True,
        .AcceptsReturn = True,
        .WordWrap = True
    }

        ''' <summary>
        ''' Clears history and persisted chat state.
        ''' </summary>
        Private ReadOnly _btnClear As Button = New Button() With {.Text = "Clear", .AutoSize = True}

        ''' <summary>
        ''' Closes the form.
        ''' </summary>
        Private ReadOnly _btnClose As Button = New Button() With {.Text = "Close", .AutoSize = True}

        ''' <summary>
        ''' Sends the current question to the assistant.
        ''' </summary>
        Private ReadOnly _btnSend As Button = New Button() With {.Text = $"Ask {AssistantName}", .AutoSize = True}

        ''' <summary>
        ''' When checked, disables TopMost behavior for the form.
        ''' </summary>
        Private ReadOnly _chkNoTopMost As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Do not stay on top", .AutoSize = True}

        ''' <summary>
        ''' When checked, includes configuration file content in the LLM prompt (with API keys masked).
        ''' </summary>
        Private ReadOnly _chkIncludeConfig As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Include configuration files", .AutoSize = True}

        ''' <summary>
        ''' Indicates whether the WebBrowser document is ready to accept appended chat fragments.
        ''' </summary>
        Private _htmlReady As Boolean = False

        ''' <summary>
        ''' Buffers HTML fragments while the WebBrowser document is loading.
        ''' </summary>
        Private ReadOnly _htmlQueue As New List(Of String)()

        ''' <summary>
        ''' Holds the DOM element id for the last "Thinking..." message, used for removal.
        ''' </summary>
        Private _lastThinkingId As String = Nothing

        ''' <summary>
        ''' In-memory conversation history (role/content) used both for transcript and for sending context to the LLM.
        ''' </summary>
        Private ReadOnly _history As New List(Of (Role As String, Content As String))()

        ''' <summary>
        ''' Cached manual text for the currently configured manual path/URL.
        ''' </summary>
        Private _manualCache As String = Nothing

        ''' <summary>
        ''' Manual path/URL used for the current cached manual text.
        ''' </summary>
        Private _manualCachePath As String = Nothing

        ''' <summary>
        ''' Flag used to prevent overlapping welcome generation; 0 = none, 1 = running.
        ''' </summary>
        Private _welcomeInProgress As Integer = 0 ' 0 = none, 1 = running

        ''' <summary>
        ''' Indicates whether alternate model availability has already been resolved for this instance.
        ''' </summary>
        Private _helpMeAltResolved As Boolean = False

        ''' <summary>
        ''' Indicates whether an alternate "HelpMe" model configuration is available and will be applied.
        ''' </summary>
        Private _helpMeAltAvailable As Boolean = False

        ''' <summary>
        ''' Serializes LLM calls to avoid concurrent model switching/restoring.
        ''' </summary>
        Private ReadOnly _modelSemaphore As New Threading.SemaphoreSlim(1, 1)


        ''' <summary>
        ''' Initializes the chat form UI, Markdown pipeline, and hooks event handlers.
        ''' </summary>
        ''' <param name="context">Shared context providing prompts and configuration.</param>
        ''' <param name="hostAppName">Optional host application name for system prompt enrichment.</param>
        Public Sub New(context As ISharedContext, Optional hostAppName As String = "")
            MyBase.New()
            _context = context
            _hostAppName = hostAppName

            Me.Text = WindowTitle
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.StartPosition = FormStartPosition.Manual
            Me.MinimumSize = New System.Drawing.Size(720, 420)
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
            Try
                Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
            Catch
            End Try

            Dim table As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(10)
        }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            _txtInput.Margin = New Padding(0, 10, 0, 6)
            Dim threeLines = (_txtInput.Font.Height * 3) + 10
            _txtInput.MinimumSize = New System.Drawing.Size(0, threeLines)
            _txtInput.Height = threeLines

            Dim pnlButtons As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(0, 0, 0, 4)
        }
            pnlButtons.Controls.Add(_btnSend)
            pnlButtons.Controls.Add(_btnClear)
            pnlButtons.Controls.Add(_btnClose)
            pnlButtons.Controls.Add(_chkNoTopMost)
            pnlButtons.Controls.Add(_chkIncludeConfig)

            table.Controls.Add(_chat, 0, 0)
            table.Controls.Add(_txtInput, 0, 1)
            table.Controls.Add(pnlButtons, 0, 2)
            Me.Controls.Add(table)

            _mdPipeline = New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseEmojiAndSmiley().
            UseSoftlineBreakAsHardlineBreak().
            Build()

            AddHandler Me.Load, AddressOf OnLoadForm
            AddHandler Me.FormClosing, AddressOf OnFormClosing
            AddHandler Me.Activated, AddressOf OnActivated ' Ensure TopMost behavior reapplied.
            AddHandler _btnSend.Click, AddressOf OnSend
            AddHandler _btnClear.Click, AddressOf OnClear
            AddHandler _btnClose.Click, AddressOf OnClose
            AddHandler _txtInput.KeyDown, AddressOf OnInputKeyDown
            AddHandler _chat.DocumentCompleted, AddressOf Chat_DocumentCompleted
            AddHandler _chat.Navigating, AddressOf Chat_Navigating
            AddHandler _chat.NewWindow, AddressOf Chat_NewWindow
            AddHandler _chkNoTopMost.CheckedChanged, AddressOf OnTopMostChanged
            AddHandler _chkIncludeConfig.CheckedChanged, AddressOf OnIncludeConfigChanged
        End Sub

        ''' <summary>
        ''' Writes diagnostic output to the debug stream.
        ''' </summary>
        Private Sub Dbg(msg As String)
            Debug.WriteLine($"[HelpMeInky {DateTime.Now:HH:mm:ss.fff}] {msg}")
        End Sub

        ''' <summary>
        ''' Runs an action on the UI thread (or directly if already on the UI thread).
        ''' </summary>
        Private Sub Ui(action As Action)
            If Me.IsDisposed Then Return
            If Me.InvokeRequired Then
                Try : Me.BeginInvoke(action) : Catch : End Try
            Else
                action()
            End If
        End Sub

        ''' <summary>
        ''' Shows the form (optionally owned), restores from minimized state, activates it,
        ''' applies TopMost behavior, and focuses the input box.
        ''' </summary>
        Public Sub ShowRaised(Optional owner As IWin32Window = Nothing)
            Dbg("ShowRaised")
            If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
            If Not Me.Visible Then
                If owner IsNot Nothing Then Me.Show(owner) Else Me.Show()
            End If
            Me.Activate()
            ApplyTopMostBehavior()
            _txtInput.Focus()
            _txtInput.SelectAll()
        End Sub

        ''' <summary>
        ''' Ensures the TopMost setting is applied when the form becomes active.
        ''' </summary>
        Private Sub OnActivated(sender As Object, e As EventArgs)
            ApplyTopMostBehavior()
        End Sub

        ''' <summary>
        ''' Persists the "Do not stay on top" option and reapplies TopMost.
        ''' </summary>
        Private Sub OnTopMostChanged(sender As Object, e As EventArgs)
            Try
                My.Settings.HelpMeNoTopMost = _chkNoTopMost.Checked
                My.Settings.Save()
            Catch
            End Try
            ApplyTopMostBehavior()
        End Sub

        ''' <summary>
        ''' Persists the "Include configuration files" option.
        ''' </summary>
        Private Sub OnIncludeConfigChanged(sender As Object, e As EventArgs)
            Try
                My.Settings.HelpMeIncludeConfig = _chkIncludeConfig.Checked
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Applies TopMost behavior according to the stored checkbox setting.
        ''' </summary>
        Private Sub ApplyTopMostBehavior()
            ' If unchecked => stay always on top.
            If _chkNoTopMost IsNot Nothing AndAlso _chkNoTopMost.Checked Then
                Me.TopMost = False
            Else
                Me.TopMost = True
            End If
        End Sub

        ''' <summary>
        ''' Loads persisted window state and options, initializes the chat HTML container,
        ''' restores the last chat if present, otherwise generates a welcome message.
        ''' </summary>
        Private Async Sub OnLoadForm(sender As Object, e As EventArgs)
            Dbg("OnLoadForm start")
            Try
                If My.Settings.HelpMeFormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.HelpMeFormSize <> System.Drawing.Size.Empty Then
                    Me.Location = My.Settings.HelpMeFormLocation
                    Me.Size = My.Settings.HelpMeFormSize
                Else
                    Dim area = Screen.PrimaryScreen.WorkingArea
                    Dim w = Math.Max(Me.MinimumSize.Width, 820)
                    Dim h = Math.Max(Me.MinimumSize.Height, 500)
                    Me.Location = New System.Drawing.Point(area.Left + (area.Width - w) \ 2, area.Top + (area.Height - h) \ 2)
                    Me.Size = New System.Drawing.Size(w, h)
                End If
            Catch ex As Exception
                Dbg("Restore bounds error: " & ex.Message)
            End Try

            ' Load persisted TopMost choice (default False => window stays on top).
            Try
                _chkNoTopMost.Checked = My.Settings.HelpMeNoTopMost
            Catch
                _chkNoTopMost.Checked = False
            End Try

            ' Load persisted IncludeConfig choice.
            Try
                _chkIncludeConfig.Checked = My.Settings.HelpMeIncludeConfig
            Catch
                _chkIncludeConfig.Checked = False
            End Try

            ApplyTopMostBehavior()

            InitializeChatHtml()

            If Not String.IsNullOrEmpty(My.Settings.LastHelpMeChatHtml) Then
                AppendHtml(My.Settings.LastHelpMeChatHtml)
            ElseIf Not String.IsNullOrEmpty(My.Settings.LastHelpMeChat) Then
                AppendTranscriptToHtml(My.Settings.LastHelpMeChat)
            Else
                Await SafeGenerateWelcomeAsync()
            End If
        End Sub

        ''' <summary>
        ''' Persists transcript, HTML, bounds, and option flags when the form is closing.
        ''' </summary>
        Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
            Dbg("OnFormClosing")
            Try
                PersistTranscriptLimited()
                PersistChatHtml()
                If Me.WindowState = FormWindowState.Normal Then
                    My.Settings.HelpMeFormLocation = Me.Location
                    My.Settings.HelpMeFormSize = Me.Size
                Else
                    My.Settings.HelpMeFormLocation = Me.RestoreBounds.Location
                    My.Settings.HelpMeFormSize = Me.RestoreBounds.Size
                End If
                My.Settings.HelpMeNoTopMost = _chkNoTopMost.Checked
                My.Settings.HelpMeIncludeConfig = _chkIncludeConfig.Checked
                My.Settings.Save()
            Catch ex As Exception
                Dbg("OnFormClosing error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Sends the current textbox content as a user message and starts an async LLM request.
        ''' </summary>
        Private Sub OnSend(sender As Object, e As EventArgs)
            Dim userText = _txtInput.Text.Trim()
            Dbg($"OnSend len={userText.Length}")
            If userText.Length = 0 Then Return

            AppendUserHtml(userText)
            _history.Add(("user", userText))
            _txtInput.Clear()
            ShowAssistantThinking()
            Dim __ = SendAsync(userText)
        End Sub

        ''' <summary>
        ''' Clears chat history and persisted state, then generates a new welcome message.
        ''' </summary>
        Private Async Sub OnClear(sender As Object, e As EventArgs)
            Dbg("OnClear start")
            Try
                _history.Clear()
                InitializeChatHtml()
                My.Settings.LastHelpMeChat = ""
                My.Settings.LastHelpMeChatHtml = ""
                My.Settings.Save()
                Dbg("State cleared & saved")
                Await SafeGenerateWelcomeAsync().ConfigureAwait(False)
            Catch ex As Exception
                Dbg("OnClear error: " & ex.Message)
            Finally
                If _txtInput.InvokeRequired Then
                    _txtInput.BeginInvoke(Sub() _txtInput.Focus())
                Else
                    _txtInput.Focus()
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Closes the form.
        ''' </summary>
        Private Sub OnClose(sender As Object, e As EventArgs)
            Dbg("OnClose")
            Me.Close()
        End Sub

        ''' <summary>
        ''' Handles Enter to send and Escape to close.
        ''' </summary>
        Private Sub OnInputKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                OnSend(Me, EventArgs.Empty)
            ElseIf e.KeyCode = Keys.Escape Then
                Me.Close()
            End If
        End Sub

        ''' <summary>
        ''' Generates a welcome message if not already in progress.
        ''' </summary>
        Private Async Function SafeGenerateWelcomeAsync() As Task
            If Interlocked.CompareExchange(_welcomeInProgress, 1, 0) <> 0 Then
                Dbg("SafeGenerateWelcomeAsync skipped: already running")
                Exit Function
            End If
            Try
                Await GenerateWelcomeAsync()
            Catch ex As Exception
                Dbg("SafeGenerateWelcomeAsync fatal: " & ex.Message)
                RemoveAssistantThinking()
                AppendAssistantMarkdown("*(Welcome failed: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
            Finally
                Interlocked.Exchange(_welcomeInProgress, 0)
            End Try
        End Function

        ''' <summary>
        ''' Builds a welcome prompt (including day-part and manual availability) and appends the assistant response.
        ''' </summary>
        Private Async Function GenerateWelcomeAsync() As Task
            Dbg("GenerateWelcomeAsync start")
            Dim langName As String = System.Globalization.CultureInfo.CurrentUICulture.DisplayName
            Dim partOfDay As String = GetPartOfDay()
            Dim manualText As String = Await GetManualOnceAsync()
            Dim systemPrompt As String

            manualText = manualText.Trim()
            If manualText.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) OrElse manualText = "" Then
                systemPrompt = $"Generate a brief, friendly {langName} welcome that naturally references it is {partOfDay} now, but tell the user that you can't work because you have no access to the manual (which needs to be configured and is retrieved either via an URL or file path; most likely, the path/URL is wrong or not working). Advise that the configured source should be checked or configured as per the manual."
            Else
                systemPrompt = $"Generate a brief, friendly {langName} welcome that naturally references it is {partOfDay} now and asks what you can do. Do NOT state your name. One short short sentence, not talkative."
            End If
            Dim userPrompt As String = ""
            Dim answer As String = ""
            Try
                Dim sw = Stopwatch.StartNew()
                answer = Await CallHelpMeLlmAsync(systemPrompt, userPrompt)
                sw.Stop()
                Dbg($"Welcome LLM ms={sw.ElapsedMilliseconds} rawLen={If(answer, "").Length}")
            Catch ex As Exception
                Dbg("Welcome LLM error: " & ex.Message)
                answer = "Hello! How can I help?"
            End Try

            answer = If(answer, "").Trim()
            AppendAssistantMarkdown(answer)
            _history.Add(("assistant", answer))
            PersistChatHtml()
            Dbg("GenerateWelcomeAsync done")
        End Function

        ''' <summary>
        ''' Builds the LLM prompt (user question + manual + conversation + optional config files)
        ''' and appends the assistant response to the chat.
        ''' </summary>
        Private Async Function SendAsync(userText As String) As Task
            Dbg("SendAsync start")
            Try
                Dim hostInfo As String = If(String.IsNullOrEmpty(_hostAppName), "", $" (Host application (and version of {AN} add-in): Microsoft {_hostAppName})")
                Dim systemPrompt As String = _context.SP_HelpMe & hostInfo
                Dim manualText As String = Await GetManualOnceAsync()
                Dim convo As String = BuildConversationForLlm()

                manualText = manualText.Trim()
                If manualText.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) Or manualText = "" Then
                    manualText = "No manual"
                End If
                Dim sb As New StringBuilder()
                sb.AppendLine("User question:")
                sb.AppendLine(userText)
                sb.AppendLine()
                sb.AppendLine("Manual:")
                sb.AppendLine(manualText)
                sb.AppendLine()
                sb.AppendLine("Conversation so far:")
                sb.AppendLine(convo)

                ' Include configuration files if the checkbox is checked.
                If _chkIncludeConfig.Checked Then
                    Dim configContent = GetConfigurationContent()
                    If Not String.IsNullOrEmpty(configContent) Then
                        sb.AppendLine()
                        sb.AppendLine(configContent)
                    End If
                End If

                Dim sw = Stopwatch.StartNew()
                Dim answer As String = Await CallHelpMeLlmAsync(systemPrompt, sb.ToString())
                sw.Stop()

                answer = If(answer, "").Trim()
                Dbg($"SendAsync ms={sw.ElapsedMilliseconds} ansLen={answer.Length}")

                RemoveAssistantThinking()
                AppendAssistantMarkdown(answer)
                _history.Add(("assistant", answer))
                PersistChatHtml()
            Catch ex As Exception
                Dbg("SendAsync error: " & ex.Message)
                RemoveAssistantThinking()
                AppendAssistantMarkdown("*(Error: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
            End Try
        End Function

        ''' <summary>
        ''' Reads relevant configuration files and returns a tagged string for inclusion in the LLM prompt.
        ''' API keys are masked by <see cref="SanitizeConfigContent"/>.
        ''' </summary>
        Private Function GetConfigurationContent() As String
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine("<Configuration Files>")

                ' Get main config file.
                Dim mainPath As String = Nothing
                Try
                    mainPath = GetActiveConfigFilePath(_context)
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(mainPath) AndAlso File.Exists(mainPath) Then
                    sb.AppendLine($"<Main Configuration ({AN2}.ini)>")
                    sb.AppendLine($"Path: {mainPath}")
                    Try
                        Dim content = File.ReadAllText(mainPath)
                        sb.AppendLine(SanitizeConfigContent(content))
                    Catch ex As Exception
                        sb.AppendLine($"Error reading file: {ex.Message}")
                    End Try
                    sb.AppendLine($"</Main Configuration>")
                End If

                ' Get default INI paths if available.
                Try
                    Dim defaultPaths = SharedMethods.DefaultINIPaths
                    For Each kvp In defaultPaths
                        Dim p = Environment.ExpandEnvironmentVariables(kvp.Value)
                        If File.Exists(p) AndAlso Not String.Equals(p, mainPath, StringComparison.OrdinalIgnoreCase) Then
                            sb.AppendLine($"<{kvp.Key} Configuration>")
                            sb.AppendLine($"Path: {p}")
                            Try
                                Dim content = File.ReadAllText(p)
                                sb.AppendLine(SanitizeConfigContent(content))
                            Catch ex As Exception
                                sb.AppendLine($"Error reading file: {ex.Message}")
                            End Try
                            sb.AppendLine($"</{kvp.Key} Configuration>")
                        End If
                    Next
                Catch
                    ' DefaultINIPaths might not be accessible.
                End Try

                ' Alternate model path.
                If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                    Dim alt = Environment.ExpandEnvironmentVariables(_context.INI_AlternateModelPath)
                    If File.Exists(alt) AndAlso Not String.Equals(alt, mainPath, StringComparison.OrdinalIgnoreCase) Then
                        sb.AppendLine("<Alternate Model Configuration>")
                        sb.AppendLine($"Path: {alt}")
                        Try
                            Dim content = File.ReadAllText(alt)
                            sb.AppendLine(SanitizeConfigContent(content))
                        Catch ex As Exception
                            sb.AppendLine($"Error reading file: {ex.Message}")
                        End Try
                        sb.AppendLine("</Alternate Model Configuration>")
                    End If
                End If

                ' Special service path.
                If Not String.IsNullOrWhiteSpace(_context.INI_SpecialServicePath) Then
                    Dim sp = Environment.ExpandEnvironmentVariables(_context.INI_SpecialServicePath)
                    If File.Exists(sp) AndAlso Not String.Equals(sp, mainPath, StringComparison.OrdinalIgnoreCase) Then
                        sb.AppendLine("<Special Service Configuration>")
                        sb.AppendLine($"Path: {sp}")
                        Try
                            Dim content = File.ReadAllText(sp)
                            sb.AppendLine(SanitizeConfigContent(content))
                        Catch ex As Exception
                            sb.AppendLine($"Error reading file: {ex.Message}")
                        End Try
                        sb.AppendLine("</Special Service Configuration>")
                    End If
                End If

                sb.AppendLine("</Configuration Files>")
                Return sb.ToString()
            Catch ex As Exception
                Dbg($"GetConfigurationContent error: {ex.Message}")
                Return $"<Configuration Files>Error retrieving configuration: {ex.Message}</Configuration Files>"
            End Try
        End Function

        ''' <summary>
        ''' Removes comment lines and masks API key values in configuration text prior to sending to the LLM.
        ''' </summary>
        Private Function SanitizeConfigContent(content As String) As String
            If String.IsNullOrEmpty(content) Then Return content

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim result As New StringBuilder()

            For Each line In lines
                Dim trimmedLine = line.TrimStart()

                ' Skip comment lines (starting with ";").
                If trimmedLine.StartsWith(";") Then
                    Continue For
                End If

                ' Mask values for "APIKey" and "APIKey_2" entries.
                Dim apiKeyMatch = Regex.Match(line, "^(\s*APIKey(?:_2)?)\s*=\s*(.*)$", RegexOptions.IgnoreCase)
                If apiKeyMatch.Success Then
                    Dim key = apiKeyMatch.Groups(1).Value
                    Dim value = apiKeyMatch.Groups(2).Value.Trim()
                    If String.IsNullOrWhiteSpace(value) Then
                        result.AppendLine(line)
                    Else
                        result.AppendLine($"{key}=[for security reasons, you are not provided with the real API key contained in this file]")
                    End If
                Else
                    result.AppendLine(line)
                End If
            Next

            Return result.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' Calls the LLM using the provided system and user prompts. Applies an alternate model configuration
        ''' when available, serialized by a semaphore to allow safe restore.
        ''' </summary>
        Private Async Function CallHelpMeLlmAsync(systemPrompt As String, userPrompt As String) As Task(Of String)
            If _context Is Nothing Then Return ""
            If Not String.IsNullOrEmpty(_hostAppName) AndAlso systemPrompt.IndexOf(_hostAppName, StringComparison.OrdinalIgnoreCase) < 0 Then
                systemPrompt &= $" (This chat runs inside Microsoft {_hostAppName}.)"
            End If

            If Not _helpMeAltResolved Then
                _helpMeAltAvailable = False
                If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                    If SharedMethods.GetSpecialTaskModel(_context, _context.INI_AlternateModelPath, "HelpMe") Then
                        _helpMeAltAvailable = True
                    End If
                End If
                _helpMeAltResolved = True
                Dbg($"Alternate HelpMe model available={_helpMeAltAvailable}")
            End If

            Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
            Dim backupConfig As ModelConfig = Nothing
            Dim appliedAlternate As Boolean = False
            Dim useSecondApi As Boolean = False
            Dim timeout As Long = 0

            Try
                If _helpMeAltAvailable Then
                    backupConfig = SharedMethods.GetCurrentConfig(_context)
                    useSecondApi = True
                    appliedAlternate = True
                    timeout = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, _context.INI_Timeout)
                Else
                    timeout = _context.INI_Timeout
                End If

                Return Await SharedMethods.LLM(_context,
                                           systemPrompt,
                                           userPrompt,
                                           "",
                                           "",
                                           timeout,
                                           useSecondApi,
                                           True).ConfigureAwait(False)
            Finally
                If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(_context, backupConfig)
                End If
                _modelSemaphore.Release()
            End Try
        End Function

        ''' <summary>
        ''' Loads the manual text once per configured manual path/URL and caches the result.
        ''' </summary>
        Private Async Function GetManualOnceAsync() As Task(Of String)
            Dim path = If(_context IsNot Nothing, _context.INI_HelpMeInkyPath, "")
            If String.IsNullOrWhiteSpace(path) Then Return ""
            If _manualCache IsNot Nothing AndAlso String.Equals(_manualCachePath, path, StringComparison.OrdinalIgnoreCase) Then
                Return _manualCache
            End If
            Dbg("Loading manual fresh: " & path)
            Dim loaded = Await GetManualTextFreshAsync(path, _context)
            If Not String.IsNullOrEmpty(loaded) Then
                _manualCache = loaded
                _manualCachePath = path
            End If
            Return If(_manualCache, "")
        End Function

        ''' <summary>
        ''' Loads manual text from a URL or local file path. Supports HTML/plain text, PDF (via temp file),
        ''' RTF (via SharedMethods), and DOCX (via Word interop).
        ''' </summary>
        Private Shared Async Function GetManualTextFreshAsync(pathOrUrl As String, Optional context As ISharedContext = Nothing) As Task(Of String)
            If String.IsNullOrWhiteSpace(pathOrUrl) Then Return ""
            Dim s As String = pathOrUrl.Trim()

            ' Ensure modern TLS (harmless if already enabled).
            Try
                '#If NETFRAMEWORK Then
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or CType(&HC00, SecurityProtocolType) ' Include TLS 1.3 if supported.
                '#End If
            Catch
            End Try

            ' Remote URL.
            If s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse s.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim handler As New HttpClientHandler()
                    handler.AllowAutoRedirect = True
                    handler.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate

                    Using client As New HttpClient(handler)
                        client.Timeout = TimeSpan.FromSeconds(30)
                        ' Some servers reject requests without a user-agent.
                        Try
                            client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "RedInk/1.0 (+https://redink.ai)")
                            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/pdf, text/*, */*")
                        Catch
                        End Try

                        Using resp As HttpResponseMessage = Await client.GetAsync(s, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                            If Not resp.IsSuccessStatusCode Then Return ""

                            Dim data As Byte() = Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)

                            ' Extract media type if provided.
                            Dim mediaType As String = ""
                            If resp.Content IsNot Nothing AndAlso resp.Content.Headers IsNot Nothing AndAlso resp.Content.Headers.ContentType IsNot Nothing Then
                                If Not String.IsNullOrEmpty(resp.Content.Headers.ContentType.MediaType) Then
                                    mediaType = resp.Content.Headers.ContentType.MediaType.ToLowerInvariant()
                                End If
                            End If

                            ' PDF detection.
                            Dim isPdf As Boolean = False

                            ' 1) Declared content-type.
                            If Not String.IsNullOrEmpty(mediaType) AndAlso mediaType.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                isPdf = True
                            End If

                            ' 2) URL contains ".pdf" anywhere (also handles querystring).
                            If Not isPdf Then
                                If s.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    isPdf = True
                                End If
                            End If

                            ' 3) Magic header scan for "%PDF" within first KB (after possible BOM or garbage).
                            If Not isPdf AndAlso data IsNot Nothing AndAlso data.Length >= 4 Then
                                Dim scanMax As Integer = Math.Min(data.Length - 4, 1024)
                                Dim i As Integer = 0
                                While i <= scanMax
                                    If data(i) = AscW("%"c) AndAlso data(i + 1) = AscW("P"c) AndAlso data(i + 2) = AscW("D"c) AndAlso data(i + 3) = AscW("F"c) Then
                                        isPdf = True
                                        Exit While
                                    End If
                                    i += 1
                                End While
                            End If

                            If isPdf Then
                                Try
                                    Dim tmpPath As String = Path.Combine(Path.GetTempPath(), "manual_" & Guid.NewGuid().ToString("N") & ".pdf")
                                    File.WriteAllBytes(tmpPath, data)
                                    Return Await SharedMethods.ReadPdfAsText(tmpPath, True, False, False, context).ConfigureAwait(False)
                                Catch
                                    Return ""
                                End Try
                            End If

                            ' Fallback: decode as text.
                            Dim enc As Encoding = Encoding.UTF8
                            Dim charset As String = ""
                            If resp.Content IsNot Nothing AndAlso resp.Content.Headers IsNot Nothing AndAlso resp.Content.Headers.ContentType IsNot Nothing Then
                                charset = resp.Content.Headers.ContentType.CharSet
                            End If
                            If Not String.IsNullOrEmpty(charset) Then
                                Try
                                    enc = Encoding.GetEncoding(charset)
                                Catch
                                    enc = Encoding.UTF8
                                End Try
                            End If

                            Dim text As String = enc.GetString(data)

                            ' HTML -> plain.
                            If Not String.IsNullOrEmpty(mediaType) AndAlso mediaType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                If LooksLikeHtml(text) Then
                                    Return HtmlToPlain(text)
                                Else
                                    Return text
                                End If
                            End If

                            ' Generic octet-stream sometimes still is HTML.
                            If LooksLikeHtml(text) Then
                                Return HtmlToPlain(text)
                            End If

                            Return text
                        End Using
                    End Using
                Catch
                    Return ""
                End Try
            End If

            ' Local file path.
            Try
                If Not File.Exists(s) Then Return ""
                Select Case Path.GetExtension(s).ToLowerInvariant()
                    Case ".txt", ".md", ".log"
                        Return File.ReadAllText(s, Encoding.UTF8)
                    Case ".docx"
                        Return ReadDocxWithWordInterop(s)
                    Case ".rtf"
                        Try
                            Return SharedMethods.ReadRtfAsText(s)
                        Catch
                            Return ""
                        End Try
                    Case ".pdf"
                        Try
                            Return Await SharedMethods.ReadPdfAsText(s, True, False, False, context).ConfigureAwait(False)
                        Catch
                            Return ""
                        End Try
                    Case Else
                        Return File.ReadAllText(s, Encoding.UTF8)
                End Select
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Reads a DOCX file using Word interop and returns its plain text content.
        ''' </summary>
        ''' <remarks>
        ''' If this method creates a new Word instance, it will be closed before returning.
        ''' </remarks>
        Private Shared Function ReadDocxWithWordInterop(path As String) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim createdApp As Boolean = False

            Try
                Try
                    app = CType(Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                    createdApp = False
                Catch
                    app = New Microsoft.Office.Interop.Word.Application() With {.Visible = False}
                    createdApp = True
                End Try

                Dim fileName As Object = path
                doc = app.Documents.Open(fileName, [ReadOnly]:=True, Visible:=False)

                Dim txt = If(doc.Content?.Text, "")

                doc.Close(SaveChanges:=False)
                Return txt
            Catch
                Try
                    If doc IsNot Nothing Then doc.Close(SaveChanges:=False)
                Catch
                End Try
                Return ""
            Finally
                ' Release COM objects in the correct order.
                Try
                    If doc IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(doc)
                Catch
                Finally
                    doc = Nothing
                End Try

                ' Only quit if we created the Word instance; never quit the user's existing Word session.
                If createdApp AndAlso app IsNot Nothing Then
                    Try
                        app.Quit(SaveChanges:=False)
                    Catch
                    End Try
                End If

                Try
                    If app IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(app)
                Catch
                Finally
                    app = Nothing
                End Try

                ' Encourage COM cleanup; harmless if nothing to collect.
                Try
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                Catch
                End Try
            End Try
        End Function

        ''' <summary>
        ''' Quick heuristic to detect whether a string is likely HTML.
        ''' </summary>
        Private Shared Function LooksLikeHtml(s As String) As Boolean
            If String.IsNullOrEmpty(s) Then Return False
            Dim t = s.TrimStart()
            Return t.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("<html", StringComparison.OrdinalIgnoreCase) _
            OrElse t.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        ''' <summary>
        ''' Converts HTML content to plain text using HtmlAgilityPack.
        ''' </summary>
        Private Shared Function HtmlToPlain(html As String) As String
            Try
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(html)
                Return doc.DocumentNode.InnerText
            Catch
                Return html
            End Try
        End Function

        ''' <summary>
        ''' Initializes the WebBrowser document with base HTML/CSS/JS used for appending chat messages.
        ''' </summary>
        Private Sub InitializeChatHtml()
            Ui(Sub()
                   _htmlQueue.Clear()
                   _htmlReady = False
                   Dbg("InitializeChatHtml")
                   Dim baseSize = If(Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
                   Dim fontPt = Math.Max(CSng(baseSize + 1.0F), 10.0F)
                   Dim css =
                    $"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
                        body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
                        #chat{{padding:8px;}}
                        .msg{{margin:6px 0;word-wrap:break-word;}}
                        .msg .who{{font-weight:600;margin-right:4px;}}
                        .msg.user .who{{color:#333;}}
                        .msg.assistant .who{{color:#003366;}}
                        .msg.thinking .content{{opacity:.75;font-style:italic;}}
                        a{{color:#0068c9;text-decoration:underline;cursor:pointer;}}
                        pre{{white-space:pre-wrap;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:4px;padding:6px;}}"
                   Dim html =
                    $"<!DOCTYPE html>
                        <html>
                        <head>
                        <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
                        <meta charset=""utf-8"">
                        <style>{css}</style>
                        <script>
                        function appendMessage(html) {{
                          var c=document.getElementById('chat'); if(!c) return;
                          var temp=document.createElement('div'); temp.innerHTML=html;
                          while(temp.firstChild){{c.appendChild(temp.firstChild);}}
                          window.scrollTo(0, document.body.scrollHeight);
                        }}
                        function removeById(id) {{
                          var el=document.getElementById(id); if(!el||!el.parentNode) return;
                          el.parentNode.removeChild(el);
                        }}
                        </script>
                        </head>
                        <body><div id=""chat""></div></body>
                        </html>"
                   _chat.DocumentText = html
               End Sub)
        End Sub

        ''' <summary>
        ''' Marks the WebBrowser document as ready and flushes any queued HTML fragments.
        ''' </summary>
        Private Sub Chat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
            _htmlReady = True
            Dbg("DocumentCompleted flushQueue=" & _htmlQueue.Count)
            If _htmlQueue.Count > 0 Then
                Try
                    For Each frag In _htmlQueue
                        _chat.Document.InvokeScript("appendMessage", New Object() {frag})
                    Next
                Catch ex As Exception
                    Dbg("Flush error: " & ex.Message)
                Finally
                    _htmlQueue.Clear()
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Intercepts navigation clicks for common schemes and opens them via the OS shell.
        ''' </summary>
        Private Sub Chat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
            Try
                Dim scheme = e.Url?.Scheme?.ToLowerInvariant()
                If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                    e.Cancel = True
                    Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
                End If
            Catch ex As Exception
                Dbg("Navigating error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Cancels popup windows from the embedded browser.
        ''' </summary>
        Private Sub Chat_NewWindow(sender As Object, e As CancelEventArgs)
            e.Cancel = True
        End Sub

        ''' <summary>
        ''' Appends an HTML fragment to the browser or queues it if the document is not ready.
        ''' </summary>
        Private Sub AppendHtml(fragment As String)
            If String.IsNullOrEmpty(fragment) Then Return
            Ui(Sub()
                   If Not _htmlReady OrElse _chat.Document Is Nothing Then
                       _htmlQueue.Add(fragment)
                       Return
                   End If
                   Try
                       _chat.Document.InvokeScript("appendMessage", New Object() {fragment})
                   Catch
                       _htmlQueue.Add(fragment)
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' Adds a user message to the chat transcript as HTML-encoded text.
        ''' </summary>
        Private Sub AppendUserHtml(text As String)
            Dim encoded = WebUtility.HtmlEncode(text).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>").Replace(vbCr, "<br>")
            AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
            PersistChatHtml()
        End Sub

        ''' <summary>
        ''' Appends a "Thinking..." assistant message and stores its element id for later removal.
        ''' </summary>
        Private Sub ShowAssistantThinking()
            _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
            AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><span class='content'>Thinking...</span></div>")
        End Sub

        ''' <summary>
        ''' Removes the most recent "Thinking..." assistant message from the chat UI.
        ''' </summary>
        Private Sub RemoveAssistantThinking()
            If String.IsNullOrEmpty(_lastThinkingId) Then Return
            Ui(Sub()
                   Try
                       If _chat.Document IsNot Nothing Then
                           _chat.Document.InvokeScript("removeById", New Object() {_lastThinkingId})
                       End If
                   Catch
                   Finally
                       _lastThinkingId = Nothing
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' Renders assistant Markdown to HTML and appends it to the chat transcript.
        ''' </summary>
        Private Sub AppendAssistantMarkdown(md As String)
            md = If(md, "")
            Dim body = Markdig.Markdown.ToHtml(md, _mdPipeline)
            Dim t = body.Trim()
            Dim isSingle = Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                       Not Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

            If isSingle Then
                ' Single paragraph: keep fully inline.
                Dim inlineHtml = Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
                AppendHtml($"<div class='msg assistant'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><span class='content'>{inlineHtml}</span></div>")
            Else
                ' Multi-block: inline only the first paragraph; render the rest below.
                Dim m = Regex.Match(t, "^\s*<p>([\s\S]*?)</p>\s*", RegexOptions.IgnoreCase)
                If m.Success Then
                    Dim firstInline = m.Groups(1).Value
                    Dim rest = t.Substring(m.Index + m.Length).Trim()
                    Dim sb As New StringBuilder()
                    sb.Append("<div class='msg assistant'>")
                    sb.Append("<span class='who'>").Append(WebUtility.HtmlEncode(AssistantName)).Append(":</span>")
                    sb.Append("<span class='content'>").Append(firstInline).Append("</span>")
                    If rest.Length > 0 Then
                        sb.Append("<div class='content'>").Append(rest).Append("</div>")
                    End If
                    sb.Append("</div>")
                    AppendHtml(sb.ToString())
                Else
                    ' Fallback: previous behavior.
                    AppendHtml($"<div class='msg assistant'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><div class='content'>{t}</div></div>")
                End If
            End If
        End Sub

        ''' <summary>
        ''' Persists the current chat HTML content to user settings.
        ''' </summary>
        Private Sub PersistChatHtml()
            Ui(Sub()
                   Try
                       If _chat.Document Is Nothing Then Return
                       Dim root = _chat.Document.GetElementById("chat")
                       If root Is Nothing Then Return
                       My.Settings.LastHelpMeChatHtml = root.InnerHtml
                       My.Settings.Save()
                   Catch ex As Exception
                       Dbg("PersistChatHtml error: " & ex.Message)
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' Converts a persisted plain transcript back into HTML chat messages.
        ''' </summary>
        Private Sub AppendTranscriptToHtml(transcript As String)
            If String.IsNullOrEmpty(transcript) Then Return
            Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split({vbLf}, StringSplitOptions.None)
            Dim currentRole As String = Nothing
            Dim content As New StringBuilder()

            Dim flush =
            Sub()
                If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                    content.Clear() : currentRole = Nothing : Return
                End If
                If currentRole = "user" Then
                    Dim enc = WebUtility.HtmlEncode(content.ToString()).Replace(vbLf, "<br>")
                    AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{enc}</span></div>")
                Else
                    AppendAssistantMarkdown(content.ToString())
                End If
                content.Clear()
                currentRole = Nothing
            End Sub

            For Each ln In lines
                If ln.StartsWith("You:", StringComparison.OrdinalIgnoreCase) Then
                    flush() : currentRole = "user" : content.Append(ln.Substring(4).TrimStart())
                ElseIf ln.StartsWith(AssistantName & ":", StringComparison.OrdinalIgnoreCase) Then
                    flush() : currentRole = "assistant" : content.Append(ln.Substring((AssistantName & ":").Length).TrimStart())
                Else
                    If content.Length > 0 Then content.AppendLine()
                    content.Append(ln)
                End If
            Next
            flush()
            PersistChatHtml()
        End Sub

        ''' <summary>
        ''' Persists a capped plain-text transcript of the current conversation to settings.
        ''' </summary>
        Private Sub PersistTranscriptLimited()
            Dim transcript = BuildTranscriptPlain()
            Dim cap As Integer = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
            If transcript.Length > cap Then
                transcript = transcript.Substring(transcript.Length - cap)
            End If
            My.Settings.LastHelpMeChat = transcript
        End Sub

        ''' <summary>
        ''' Builds a plain-text transcript from the in-memory history list.
        ''' </summary>
        Private Function BuildTranscriptPlain() As String
            Dim sb As New StringBuilder()
            For Each m In _history
                If m.Role = "user" Then
                    sb.AppendLine("You: " & m.Content)
                Else
                    sb.AppendLine(AssistantName & ": " & m.Content)
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Builds a capped conversation string for the LLM prompt by walking history from newest to oldest.
        ''' </summary>
        Private Function BuildConversationForLlm() As String
            Dim sb As New StringBuilder()
            Dim cap As Integer = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
            Dim acc = 0
            For i = _history.Count - 1 To 0 Step -1
                Dim line = If(_history(i).Role = "user", "User: ", AssistantName & ": ") & _history(i).Content & Environment.NewLine
                If acc + line.Length > cap Then
                    Dim remain = cap - acc
                    If remain > 0 Then sb.Insert(0, line.Substring(line.Length - remain))
                    Exit For
                Else
                    sb.Insert(0, line)
                    acc += line.Length
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Removes Markdown formatting elements for transcript use.
        ''' </summary>
        Private Shared Function StripMarkdownForTranscript(md As String) As String
            If String.IsNullOrEmpty(md) Then Return ""
            Dim s = Regex.Replace(md, "```.*?```", "", RegexOptions.Singleline)
            s = s.Replace("**", "").Replace("__", "").Replace("*", "").Replace("_", "").Replace("`", "")
            Return s
        End Function

        ''' <summary>
        ''' Returns a coarse part-of-day label based on the current local hour.
        ''' </summary>
        Private Shared Function GetPartOfDay() As String
            Dim h = DateTime.Now.Hour
            If h < 12 Then Return "Morning"
            If h < 18 Then Return "Afternoon"
            Return "Evening"
        End Function
    End Class

End Namespace