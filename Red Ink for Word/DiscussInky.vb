' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DiscussInky.vb
' Purpose: Hosts the "Discuss Inky" multi-persona chat surface inside Word,
'          wiring persona selection, knowledge loading, transcript persistence,
'          and LLM invocation with optional alternate models.
'
' Architecture:
'  - UI Composition: Builds a WinForms surface composed of WebBrowser transcript,
'    multiline input box, and action buttons (Send, Persona, Knowledge, etc.).
'  - Session State: Persists persona choice, chat transcript, window geometry,
'    active-document flag, and knowledge file references via My.Settings plus
'    process-level caches.
'  - Personas & Knowledge: Loads persona prompts from local/global libraries,
'    opens arbitrary knowledge files (TXT/RTF/DOC/PDF/…), and caches their text.
'  - LLM Pipeline: Constructs prompts with persona instructions, knowledge text,
'    optional active-document excerpts, and prior conversation; routes calls
'    through SharedLibrary LLM helpers with optional alternate/secondary models.
'  - HTML Transcript: Renders chat history via Markdig HTML, keeps "thinking"
'    placeholders, restores transcripts on startup, and persists DOM fragments.
'  - External Dependencies: Relies on SharedLibrary.SharedMethods for model
'    management, file dialogs, message boxes, PDF parsing, and selection UI.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' WinForms surface for persona-driven LLM discussions tied to knowledge files.
''' </summary>
Public Class DiscussInky
    Inherits System.Windows.Forms.Form

#Region "Constants and Fields"

    Private Const AssistantName As String = Globals.ThisAddIn.AN6
    Private _currentPersonaName As String = AssistantName
    Private _currentPersonaPrompt As String = ""

    Private ReadOnly _context As ISharedContext
    Private ReadOnly _mdPipeline As Markdig.MarkdownPipeline

    ' Runtime knowledge cache (persists while Word is running, not in My.Settings)
    Private Shared _cachedKnowledgeContent As String = Nothing
    Private Shared _cachedKnowledgeFilePath As String = Nothing

    ' Random words for response variety
    Private Shared ReadOnly _randomModifiers As String() = {
        "thoughtfully", "carefully", "precisely", "clearly", "concisely",
        "helpfully", "insightfully", "thoroughly", "directly", "naturally"
    }
    Private Shared ReadOnly _rng As New Random()

    ' UI Controls
    Private ReadOnly _chat As WebBrowser = New WebBrowser() With {
        .Dock = DockStyle.Fill,
        .AllowWebBrowserDrop = False,
        .IsWebBrowserContextMenuEnabled = True,
        .WebBrowserShortcutsEnabled = True,
        .ScriptErrorsSuppressed = True
    }
    Private ReadOnly _txtInput As TextBox = New TextBox() With {
        .Dock = DockStyle.Fill,
        .Multiline = True,
        .AcceptsReturn = True,
        .WordWrap = True
    }
    Private ReadOnly _btnClear As Button = New Button() With {.Text = "Clear", .AutoSize = True}
    Private ReadOnly _btnClose As Button = New Button() With {.Text = "Close", .AutoSize = True}
    Private ReadOnly _btnSend As Button = New Button() With {.Text = $"Send", .AutoSize = True}
    Private ReadOnly _btnPersona As Button = New Button() With {.Text = "Persona", .AutoSize = True}
    Private ReadOnly _btnEditPersona As Button = New Button() With {.Text = "Edit Local", .AutoSize = True}
    Private ReadOnly _btnKnowledge As Button = New Button() With {.Text = "Knowledge", .AutoSize = True}
    Private ReadOnly _btnAlternateModel As Button = New Button() With {.Text = "Alternate Model", .AutoSize = True}
    Private ReadOnly _chkIncludeActiveDoc As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Include active document", .AutoSize = True}

    ' State
    Private _htmlReady As Boolean = False
    Private ReadOnly _htmlQueue As New List(Of String)()
    Private _lastThinkingId As String = Nothing
    Private ReadOnly _history As New List(Of (Role As String, Content As String))()
    Private _knowledgeContent As String = Nothing
    Private _knowledgeFilePath As String = Nothing
    Private _welcomeInProgress As Integer = 0
    Private _personaSelectedThisSession As Boolean = False

    ' Alternate model support (new implementation matching Form1.vb pattern)
    Private _alternateModelSelected As Boolean = False
    Private _alternateModelConfig As ModelConfig = Nothing
    Private _alternateModelDisplayName As String = Nothing
    Private ReadOnly _modelSemaphore As New Threading.SemaphoreSlim(1, 1)

    ''' <summary>
    ''' Holds a persona definition loaded from a file, including its prompt and display metadata.
    ''' </summary>
    Private Structure PersonaEntry
        Public Name As String
        Public Prompt As String
        Public IsLocal As Boolean
        Public DisplayName As String
    End Structure
    Private _personas As New List(Of PersonaEntry)()

#End Region

#Region "Constructor"

    ''' <summary>
    ''' Initializes UI, loads configuration references, and wires event handlers.
    ''' </summary>
    ''' <param name="context">Shared configuration context providing INI settings and model configuration.</param>
    Public Sub New(context As ISharedContext)
        MyBase.New()
        _context = context

        Me.Text = $"Discuss this, {AssistantName}"
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.Manual
        Me.MinimumSize = New System.Drawing.Size(780, 480)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Try
            Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
        Catch
        End Try

        ' Layout
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
        Dim fiveLines = (_txtInput.Font.Height * 5) + 10
        _txtInput.MinimumSize = New System.Drawing.Size(0, fiveLines)
        _txtInput.Height = fiveLines

        Dim pnlButtons As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(0, 0, 0, 4)
        }
        pnlButtons.Controls.Add(_btnSend)
        pnlButtons.Controls.Add(_btnPersona)
        pnlButtons.Controls.Add(_btnEditPersona)
        pnlButtons.Controls.Add(_btnKnowledge)

        ' Show alternate model button if either second API is configured or an alternate INI exists
        If _context.INI_SecondAPI OrElse Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            UpdateAlternateModelButtonText()
            pnlButtons.Controls.Add(_btnAlternateModel)
        End If

        pnlButtons.Controls.Add(_btnClear)
        pnlButtons.Controls.Add(_btnClose)
        pnlButtons.Controls.Add(_chkIncludeActiveDoc)

        table.Controls.Add(_chat, 0, 0)
        table.Controls.Add(_txtInput, 0, 1)
        table.Controls.Add(pnlButtons, 0, 2)
        Me.Controls.Add(table)

        _mdPipeline = New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseEmojiAndSmiley().
            UseSoftlineBreakAsHardlineBreak().
            Build()

        ' Event handlers
        AddHandler Me.Load, AddressOf OnLoadForm
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler Me.Activated, AddressOf OnActivated
        AddHandler _btnSend.Click, AddressOf OnSend
        AddHandler _btnClear.Click, AddressOf OnClear
        AddHandler _btnClose.Click, AddressOf OnClose
        AddHandler _btnPersona.Click, AddressOf OnSelectPersona
        AddHandler _btnEditPersona.Click, AddressOf OnEditLocalPersona
        AddHandler _btnKnowledge.Click, AddressOf OnLoadKnowledge
        AddHandler _btnAlternateModel.Click, AddressOf OnAlternateModelClick
        AddHandler _txtInput.KeyDown, AddressOf OnInputKeyDown
        AddHandler _chat.DocumentCompleted, AddressOf Chat_DocumentCompleted
        AddHandler _chat.Navigating, AddressOf Chat_Navigating
        AddHandler _chat.NewWindow, AddressOf Chat_NewWindow
        AddHandler _chkIncludeActiveDoc.CheckedChanged, AddressOf OnIncludeActiveDocChanged
    End Sub

#End Region

#Region "Utility Methods"

    ''' <summary>
    ''' Executes an action on the UI thread, marshaling via BeginInvoke when required.
    ''' </summary>
    ''' <param name="action">Action to execute on the UI thread.</param>
    Private Sub Ui(action As System.Action)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Try : Me.BeginInvoke(action) : Catch : End Try
        Else
            action.Invoke()
        End If
    End Sub

    ''' <summary>
    ''' Builds the window caption to reflect persona, knowledge file, and model state.
    ''' </summary>
    Private Sub UpdateWindowTitle()
        Dim title = $"Discuss this, {_currentPersonaName}"
        If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
            title &= $" - {Path.GetFileName(_knowledgeFilePath)}"
        End If

        ' Show current model in title if alternate is selected
        If _alternateModelSelected AndAlso Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
            title &= $" (using {_alternateModelDisplayName})"
        End If

        Ui(Sub() Me.Text = title)
    End Sub

    ''' <summary>
    ''' Refreshes the Send button label with the current persona name.
    ''' </summary>
    Private Sub UpdateSendButtonText()
        Ui(Sub() _btnSend.Text = $"Send to {_currentPersonaName}")
    End Sub

    ''' <summary>
    ''' Returns a random adverb used to vary assistant tone.
    ''' </summary>
    ''' <returns>Randomly selected adverb string.</returns>
    Private Function GetRandomModifier() As String
        Return _randomModifiers(_rng.Next(_randomModifiers.Length))
    End Function

    ''' <summary>
    ''' Formats the current date for inclusion in LLM prompts.
    ''' </summary>
    ''' <returns>Formatted date string.</returns>
    Private Function GetDateContext() As String
        Dim now = DateTime.Now
        Return $"Today is {now:dd-MMM-yyyy}."
    End Function

#End Region

#Region "Form Events"

    ''' <summary>
    ''' Shows (or brings forward) the form and focuses the input box.
    ''' </summary>
    ''' <param name="owner">Optional owner window.</param>
    Public Sub ShowRaised(Optional owner As IWin32Window = Nothing)
        If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
        If Not Me.Visible Then
            If owner IsNot Nothing Then Me.Show(owner) Else Me.Show()
        End If
        Me.Activate()
        _txtInput.Focus()
        _txtInput.SelectAll()
    End Sub

    ''' <summary>
    ''' Handles form activation; TopMost behavior is disabled.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnActivated(sender As Object, e As EventArgs)
        ' No longer applying TopMost behavior
    End Sub

    ''' <summary>
    ''' Persists the 'include active document' checkbox state when changed.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnIncludeActiveDocChanged(sender As Object, e As EventArgs)
        Try
            My.Settings.DiscussIncludeActiveDoc = _chkIncludeActiveDoc.Checked
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Restores persisted settings, persona, knowledge cache, transcript, and optionally triggers a welcome.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Async Sub OnLoadForm(sender As Object, e As EventArgs)
        ' Restore window position/size
        Try
            If My.Settings.DiscussFormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.DiscussFormSize <> System.Drawing.Size.Empty Then
                Me.Location = My.Settings.DiscussFormLocation
                Me.Size = My.Settings.DiscussFormSize
            Else
                Dim area = Screen.PrimaryScreen.WorkingArea
                Dim w = Math.Max(Me.MinimumSize.Width, 860)
                Dim h = Math.Max(Me.MinimumSize.Height, 540)
                Me.Location = New System.Drawing.Point(area.Left + (area.Width - w) \ 2, area.Top + (area.Height - h) \ 2)
                Me.Size = New System.Drawing.Size(w, h)
            End If
        Catch
        End Try

        ' Load persisted settings
        Try : _chkIncludeActiveDoc.Checked = My.Settings.DiscussIncludeActiveDoc : Catch : _chkIncludeActiveDoc.Checked = False : End Try

        ' Load personas
        LoadPersonas()

        ' Check if persona was previously saved - if not, force selection
        Dim savedPersona = ""
        Try
            savedPersona = My.Settings.DiscussSelectedPersona
        Catch
        End Try

        Dim personaRestoredFromSettings = False
        If Not String.IsNullOrEmpty(savedPersona) Then
            Dim found = _personas.FirstOrDefault(Function(p) p.Name.Equals(savedPersona, StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrEmpty(found.Name) Then
                _currentPersonaName = found.Name
                _currentPersonaPrompt = found.Prompt
                personaRestoredFromSettings = True
            End If
        End If

        UpdateWindowTitle()
        UpdateSendButtonText()

        InitializeChatHtml()

        ' Restore chat or load knowledge
        Dim hasChat = False
        Dim restoredHtmlHadAlternateModel = False
        Try
            ' First, restore _history from plain transcript (this ensures LLM sees the conversation)
            Dim savedTranscript = My.Settings.DiscussLastChat
            If Not String.IsNullOrEmpty(savedTranscript) Then
                RestoreHistoryFromTranscript(savedTranscript)
            End If

            ' Then restore the HTML display
            Dim savedHtml = My.Settings.DiscussLastChatHtml
            If Not String.IsNullOrEmpty(savedHtml) Then
                ' Check if the restored HTML contains an alternate/secondary model switch message
                restoredHtmlHadAlternateModel = ChatHtmlIndicatesAlternateModel(savedHtml)
                AppendHtml(savedHtml)
                hasChat = True
            ElseIf Not String.IsNullOrEmpty(savedTranscript) Then
                AppendTranscriptToHtml(savedTranscript)
                hasChat = True
            End If
        Catch
        End Try

        ' If restored chat indicated an alternate model was active, notify user we're back on primary
        If hasChat AndAlso restoredHtmlHadAlternateModel Then
            ' Ensure alternate model state is reset (it should be by default, but be explicit)
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing
            UpdateAlternateModelButtonText()

            ' Notify user in chat that we're back on primary
            AppendSystemMessage($"Session restored. Now using primary model ({_context.INI_Model}).")
        End If

        ' Restore knowledge from runtime cache first, then from settings
        If Not String.IsNullOrEmpty(_cachedKnowledgeContent) AndAlso Not String.IsNullOrEmpty(_cachedKnowledgeFilePath) Then
            _knowledgeContent = _cachedKnowledgeContent
            _knowledgeFilePath = _cachedKnowledgeFilePath
            UpdateWindowTitle()
        Else
            Try
                Dim savedPath = My.Settings.DiscussKnowledgePath
                If Not String.IsNullOrEmpty(savedPath) AndAlso File.Exists(savedPath) Then
                    _knowledgeFilePath = savedPath
                    _knowledgeContent = Await LoadKnowledgeFileAsync(savedPath)
                    ' Cache for runtime
                    _cachedKnowledgeContent = _knowledgeContent
                    _cachedKnowledgeFilePath = _knowledgeFilePath
                    UpdateWindowTitle()
                End If
            Catch
            End Try
        End If

        ' Force persona selection on first run after Word starts (if not restored from settings)
        If Not personaRestoredFromSettings AndAlso _personas.Count > 0 AndAlso Not _personaSelectedThisSession Then
            OnSelectPersona(Nothing, EventArgs.Empty)
            _personaSelectedThisSession = True
        End If

        ' Prompt for knowledge if not available
        If String.IsNullOrEmpty(_knowledgeContent) AndAlso Not hasChat Then
            Await PromptForKnowledgeFileAsync()
        End If

        If Not hasChat Then
            Await SafeGenerateWelcomeAsync()
        End If
    End Sub

    ''' <summary>
    ''' Persists geometry, transcript, persona, knowledge path, and checkbox state on close.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments containing form-closing details.</param>
    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        Try
            PersistTranscriptLimited()
            PersistChatHtml()
            If Me.WindowState = FormWindowState.Normal Then
                My.Settings.DiscussFormLocation = Me.Location
                My.Settings.DiscussFormSize = Me.Size
            Else
                My.Settings.DiscussFormLocation = Me.RestoreBounds.Location
                My.Settings.DiscussFormSize = Me.RestoreBounds.Size
            End If
            My.Settings.DiscussIncludeActiveDoc = _chkIncludeActiveDoc.Checked
            My.Settings.DiscussSelectedPersona = _currentPersonaName
            My.Settings.DiscussKnowledgePath = If(_knowledgeFilePath, "")
            My.Settings.Save()
        Catch
        End Try
    End Sub

#End Region

#Region "Alternate Model Handling"

    ''' <summary>
    ''' Sets the alternate-model button caption according to availability and selection state.
    ''' </summary>
    Private Sub UpdateAlternateModelButtonText()
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            _btnAlternateModel.Text = If(_alternateModelSelected, "Primary Model", "Alternate Model")
        Else
            _btnAlternateModel.Text = "Switch Model"
        End If
    End Sub

    ''' <summary>
    ''' Handles alternate model toggling or selection, mirroring Form1 pattern.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnAlternateModelClick(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            ' If an alternate is already active -> switch back to primary without dialog
            If _alternateModelSelected Then
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
                UpdateAlternateModelButtonText()
                UpdateWindowTitle()
                AppendSystemMessage($"Switched back to primary model ({_context.INI_Model}).")
                Return
            End If

            ' Pre-check: verify the alternate model file exists and has content
            Dim altPath = ExpandEnvironmentVariables(_context.INI_AlternateModelPath)
            If String.IsNullOrWhiteSpace(altPath) OrElse Not File.Exists(altPath) Then
                AppendSystemMessage("Alternate model configuration file not found.")
                Return
            End If

            ' Selecting an alternate
            SharedMethods.LastAlternateModel = "" ' sentinel
            Dim ok As Boolean = SharedMethods.ShowModelSelection(
                _context,
                _context.INI_AlternateModelPath,
                "Alternate Model",
                "Select the alternate model you want to use:",
                "",
                2
            )
            If Not ok Then
                ' User cancelled
                Return
            End If

            ' The selector applies the chosen model to the context at this point.
            ' Snapshot it, then restore the original immediately so globals remain clean.
            Dim justApplied As ModelConfig = SharedMethods.GetCurrentConfig(_context)

            If SharedMethods.originalConfigLoaded Then
                SharedMethods.RestoreDefaults(_context, SharedMethods.originalConfig)
            End If
            SharedMethods.originalConfigLoaded = False

            Dim userChoseAlternate As Boolean = Not String.IsNullOrWhiteSpace(SharedMethods.LastAlternateModel)

            If userChoseAlternate Then
                _alternateModelSelected = True
                _alternateModelConfig = justApplied
                _alternateModelDisplayName = SharedMethods.LastAlternateModel
                AppendSystemMessage($"Switched to alternate model: {_alternateModelDisplayName}")
            Else
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
            End If

            UpdateAlternateModelButtonText()
            UpdateWindowTitle()
        Else
            ' Legacy behavior: simple toggle to secondary model (if configured)
            If _context.INI_SecondAPI Then
                ' Toggle between primary and secondary
                If _alternateModelSelected Then
                    _alternateModelSelected = False
                    _alternateModelConfig = Nothing
                    _alternateModelDisplayName = Nothing
                    AppendSystemMessage($"Switched back to primary model ({_context.INI_Model}).")
                Else
                    _alternateModelSelected = True
                    _alternateModelDisplayName = _context.INI_Model_2
                    AppendSystemMessage($"Switched to secondary model: {_alternateModelDisplayName}")
                End If
                UpdateAlternateModelButtonText()
                UpdateWindowTitle()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Runs an LLM request while temporarily applying any selected alternate model, restoring afterward.
    ''' </summary>
    ''' <param name="systemPrompt">System instruction for the LLM.</param>
    ''' <param name="userPrompt">User message payload.</param>
    ''' <returns>LLM-generated response text.</returns>
    Private Async Function CallLlmWithSelectedModelAsync(systemPrompt As String, userPrompt As String) As Task(Of String)
        Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
        Dim backupConfig As ModelConfig = Nothing
        Dim appliedAlternate As Boolean = False
        Dim useSecondApi As Boolean = False

        Try
            ' If the user selected an alternate model, apply it to the context as the "second model" just for this call.
            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                ' Back up current config (the "original state at rest")
                backupConfig = SharedMethods.GetCurrentConfig(_context)

                ' Apply the selected alternate config
                SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                appliedAlternate = True

                ' Enforce second API usage for alternate models
                useSecondApi = True
            ElseIf _alternateModelSelected AndAlso _alternateModelConfig Is Nothing AndAlso _context.INI_SecondAPI Then
                ' Legacy toggle: use second API without config swap
                useSecondApi = True
            End If

            ' Execute the LLM call
            Return Await LLM(_context,
                             systemPrompt,
                             userPrompt,
                             "",
                             "",
                             0,
                             useSecondApi,
                             True).ConfigureAwait(False)

        Finally
            ' Always restore the original config after the call so the rest of the add-in sees the original state.
            If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                SharedMethods.RestoreDefaults(_context, backupConfig)
            End If
            _modelSemaphore.Release()
        End Try
    End Function

#End Region

#Region "Persona Management"

    ''' <summary>
    ''' Loads persona definitions from configured local and global files into memory.
    ''' </summary>
    Private Sub LoadPersonas()
        _personas.Clear()

        Dim localPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))
        Dim globalPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPath, ""))

        Dim localLoaded = False
        Dim globalLoaded = False

        ' Load local personas first (marked with (local))
        If Not String.IsNullOrWhiteSpace(localPath) Then
            localLoaded = LoadPersonasFromFile(localPath, isLocal:=True)
        End If

        ' Load global personas
        If Not String.IsNullOrWhiteSpace(globalPath) Then
            globalLoaded = LoadPersonasFromFile(globalPath, isLocal:=False)
        End If

        ' Show error only if both paths are configured but neither loaded any personas
        If _personas.Count = 0 Then
            If Not String.IsNullOrWhiteSpace(localPath) OrElse Not String.IsNullOrWhiteSpace(globalPath) Then
                AppendSystemMessage("No personas could be loaded. Please check your persona configuration files.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Parses a persona file, appending entries and marking whether they are local.
    ''' </summary>
    ''' <param name="filePath">Path to the persona file.</param>
    ''' <param name="isLocal">Whether the file is local (vs. global).</param>
    ''' <returns>True if any persona entries were loaded.</returns>
    Private Function LoadPersonasFromFile(filePath As String, isLocal As Boolean) As Boolean
        ' Must be a file, not a directory
        If String.IsNullOrWhiteSpace(filePath) Then
            Return False
        End If

        If Directory.Exists(filePath) Then
            AppendSystemMessage($"Persona path must be a file, not a directory: {filePath}")
            Return False
        End If

        If Not File.Exists(filePath) Then
            ' Only show error if this is likely intentional (path is configured but file missing)
            ' Don't show error for empty/unconfigured paths
            Return False
        End If

        Dim loadedAny = False
        Try
            For Each rawLine In File.ReadAllLines(filePath, Encoding.UTF8)
                Dim line = If(rawLine, "").Trim()

                ' Skip empty lines and comments
                If line.Length = 0 OrElse line.StartsWith(";", StringComparison.Ordinal) Then
                    Continue For
                End If

                ' Parse Name|Prompt format
                Dim pipeIdx = line.IndexOf("|"c)
                If pipeIdx < 1 Then Continue For

                Dim name = line.Substring(0, pipeIdx).Trim()
                Dim prompt = line.Substring(pipeIdx + 1).Trim()

                If name.Length = 0 OrElse prompt.Length = 0 Then Continue For

                ' Create unique display name
                Dim displayName = name & If(isLocal, " (local)", "")
                displayName = MakeUniqueDisplay(displayName, _personas.Select(Function(p) p.DisplayName).ToList())

                _personas.Add(New PersonaEntry With {
                    .Name = name,
                    .Prompt = prompt,
                    .IsLocal = isLocal,
                    .DisplayName = displayName
                })
                loadedAny = True
            Next
        Catch ex As Exception
            AppendSystemMessage($"Error loading persona file: {ex.Message}")
            Return False
        End Try

        Return loadedAny
    End Function

    ''' <summary>
    ''' Ensures persona display names are unique by appending numeric suffixes.
    ''' </summary>
    ''' <param name="baseText">Initial display name.</param>
    ''' <param name="existing">Collection of existing display names.</param>
    ''' <returns>Unique display name.</returns>
    Private Function MakeUniqueDisplay(baseText As String, existing As ICollection(Of String)) As String
        If Not existing.Contains(baseText) Then Return baseText
        Dim n = 2
        While True
            Dim candidate = baseText & " [" & n.ToString() & "]"
            If Not existing.Contains(candidate) Then Return candidate
            n += 1
        End While
    End Function

    ''' <summary>
    ''' Shows persona picker and applies the chosen persona prompt.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnSelectPersona(sender As Object, e As EventArgs)
        If _personas.Count = 0 Then
            ShowCustomMessageBox("No personas configured. Please configure INI_DiscussInkyPath or INI_DiscussInkyPathLocal in your settings.",
                                 extraButtonText:="Edit Local Personas",
                                 extraButtonAction:=Sub() OnEditLocalPersona(Nothing, EventArgs.Empty))
            Return
        End If

        ' Build selection items
        Dim items As New List(Of SelectionItem)()
        For i = 0 To _personas.Count - 1
            items.Add(New SelectionItem(_personas(i).DisplayName, i + 1))
        Next

        ' Find current selection
        Dim defaultVal = 1
        For i = 0 To _personas.Count - 1
            If _personas(i).Name.Equals(_currentPersonaName, StringComparison.OrdinalIgnoreCase) Then
                defaultVal = i + 1
                Exit For
            End If
        Next

        Dim result = SelectValue(items, defaultVal, "Select the persona discussing:", AN & " - Select Persona")

        If result > 0 AndAlso result <= _personas.Count Then
            Dim selected = _personas(result - 1)
            _currentPersonaName = selected.Name
            _currentPersonaPrompt = selected.Prompt
            _personaSelectedThisSession = True
            UpdateWindowTitle()
            UpdateSendButtonText()

            Try
                My.Settings.DiscussSelectedPersona = _currentPersonaName
                My.Settings.Save()
            Catch
            End Try

            AppendSystemMessage($"Persona changed to: {_currentPersonaName}")
        End If
    End Sub

    ''' <summary>
    ''' Ensures the local persona file exists and opens it in the shared text editor.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnEditLocalPersona(sender As Object, e As EventArgs)
        Dim localPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))

        If String.IsNullOrWhiteSpace(localPath) Then
            ShowCustomMessageBox("INI_DiscussInkyPathLocal is not configured in your settings.")
            Return
        End If

        ' Create directory if needed
        Dim dir = Path.GetDirectoryName(localPath)
        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
            Try
                Directory.CreateDirectory(dir)
            Catch ex As Exception
                ShowCustomMessageBox($"Cannot create directory: {ex.Message}")
                Return
            End Try
        End If

        ' Create file with sample content if it doesn't exist
        If Not File.Exists(localPath) Then
            Try
                File.WriteAllText(localPath,
                    "; DiscussInky Local Personas" & vbCrLf &
                    "; Format: Name|System Prompt" & vbCrLf &
                    "; Lines starting with ; are comments" & vbCrLf &
                    vbCrLf &
                    "Teacher|You are a teacher and will do an exam with the user based on the knowledge you will be provided. Check the responses and provide feedback." & vbCrLf &
                    "Summarizer|Summarize the knowledge document for the user in a clear and concise way. Answer follow-up questions about the content." & vbCrLf,
                    Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox($"Cannot create file: {ex.Message}")
                Return
            End Try
        End If

        ShowTextFileEditor(localPath, $"{AN} - Edit Local Personas (changes active after restart):", False, _context)
    End Sub

#End Region

#Region "Knowledge File Management"

    ''' <summary>
    ''' Button handler that launches the knowledge file picker.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Async Sub OnLoadKnowledge(sender As Object, e As EventArgs)
        Await PromptForKnowledgeFileAsync()
    End Sub

    ''' <summary>
    ''' Prompts the user for a knowledge file, loads its text, caches it, and updates state.
    ''' </summary>
    Private Async Function PromptForKnowledgeFileAsync() As Task
        Try
            Globals.ThisAddIn.DragDropFormLabel = "Drag & drop a knowledge file (PDF, Word, TXT, etc.) or click Browse"
            Globals.ThisAddIn.DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.md;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                                                   "Text Files (*.txt;*.md)|*.txt;*.md|" &
                                                   "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                                   "PDF Files (*.pdf)|*.pdf|" &
                                                   "All Files (*.*)|*.*"

            Dim filePath = Globals.ThisAddIn.GetFileName()

            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""

            If String.IsNullOrWhiteSpace(filePath) Then
                Return
            End If

            ShowAssistantThinking()
            _knowledgeContent = Await LoadKnowledgeFileAsync(filePath)
            RemoveAssistantThinking()

            If String.IsNullOrWhiteSpace(_knowledgeContent) Then
                AppendSystemMessage("Failed to load knowledge file or file is empty.")
                Return
            End If

            _knowledgeFilePath = filePath

            ' Update runtime cache
            _cachedKnowledgeContent = _knowledgeContent
            _cachedKnowledgeFilePath = _knowledgeFilePath

            UpdateWindowTitle()

            Try
                My.Settings.DiscussKnowledgePath = filePath
                My.Settings.Save()
            Catch
            End Try

            AppendSystemMessage($"Knowledge loaded: {Path.GetFileName(filePath)} ({_knowledgeContent.Length:N0} characters)")

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendSystemMessage($"Error loading knowledge file: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Reads supported file formats into plain text, delegating to specialized readers when needed.
    ''' </summary>
    ''' <param name="filePath">Path to the knowledge file.</param>
    ''' <returns>Plain text content of the file.</returns>
    Private Async Function LoadKnowledgeFileAsync(filePath As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Return ""
        End If

        Try
            Dim ext = Path.GetExtension(filePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".md", ".log", ".ini", ".csv", ".json", ".xml", ".html", ".htm"
                    Return File.ReadAllText(filePath, Encoding.UTF8)

                Case ".rtf"
                    Return ReadRtfAsText(filePath)

                Case ".doc", ".docx"
                    Return ReadWordDocument(filePath)

                Case ".pdf"
                    Return Await ReadPdfAsText(filePath, True, False, False, _context)

                Case Else
                    Return File.ReadAllText(filePath, Encoding.UTF8)
            End Select

        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

#Region "Chat Actions"

    ''' <summary>
    ''' Captures the user's message, adds it to history, and starts asynchronous LLM processing.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnSend(sender As Object, e As EventArgs)
        Dim userText = _txtInput.Text.Trim()
        If userText.Length = 0 Then Return

        AppendUserHtml(userText)
        _history.Add(("user", userText))
        _txtInput.Clear()
        ShowAssistantThinking()
        Dim __ = SendAsync(userText)
    End Sub

    ''' <summary>
    ''' Clears transcript and history, then regenerates the welcome sequence.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Async Sub OnClear(sender As Object, e As EventArgs)
        Try
            _history.Clear()
            InitializeChatHtml()
            My.Settings.DiscussLastChat = ""
            My.Settings.DiscussLastChatHtml = ""
            My.Settings.Save()
            Await SafeGenerateWelcomeAsync().ConfigureAwait(False)
        Catch
        Finally
            Ui(Sub() _txtInput.Focus())
        End Try
    End Sub

    ''' <summary>
    ''' Closes the DiscussInky form.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnClose(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ''' <summary>
    ''' Handles Enter/Escape shortcuts for sending and closing.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Key event arguments.</param>
    Private Sub OnInputKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True
            OnSend(Me, EventArgs.Empty)
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

#End Region

#Region "Welcome Message"

    ''' <summary>
    ''' Serializes welcome generation and surfaces any failures in the chat.
    ''' </summary>
    Private Async Function SafeGenerateWelcomeAsync() As Task
        If Interlocked.CompareExchange(_welcomeInProgress, 1, 0) <> 0 Then
            Return
        End If
        Try
            ' Show current session info before welcome
            ShowSessionInfo()
            Await GenerateWelcomeAsync()
        Catch ex As Exception
            RemoveAssistantThinking()
            AppendAssistantMarkdown("*(Welcome failed: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
        Finally
            Interlocked.Exchange(_welcomeInProgress, 0)
        End Try
    End Function

    ''' <summary>
    ''' Posts a system message summarizing the active persona and knowledge file.
    ''' </summary>
    Private Sub ShowSessionInfo()
        Dim sb As New StringBuilder()

        ' Persona info
        sb.Append($"Persona: {_currentPersonaName}")
        sb.Append(" (change with 'Persona' button)")

        ' Knowledge document info
        If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
            sb.Append($" | Knowledge: {Path.GetFileName(_knowledgeFilePath)}")
        Else
            sb.Append(" | Knowledge: None loaded")
        End If
        sb.Append(" (change with 'Knowledge' button)")

        AppendSystemMessage(sb.ToString())
    End Sub

    ''' <summary>
    ''' Requests a short persona-aware welcome message from the LLM.
    ''' </summary>
    Private Async Function GenerateWelcomeAsync() As Task
        Dim langName = System.Globalization.CultureInfo.CurrentUICulture.DisplayName
        Dim partOfDay = GetPartOfDay()
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()

        Dim systemPrompt As String

        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            systemPrompt = $"{dateContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           "Tell the user they should load a knowledge document using the 'Knowledge' button to start a discussion. " &
                           "You are ready to discuss any knowledge they provide. One short sentence, not talkative."
        Else
            ' Use persona prompt to shape the welcome message
            Dim personaContext = ""
            If Not String.IsNullOrEmpty(_currentPersonaPrompt) Then
                personaContext = $" Your persona and role is defined as: '{_currentPersonaPrompt}'. " &
                                 "Generate a welcome that fits this persona - for example, a Teacher might greet students and mention they're ready to teach, " &
                                 "an Examiner might announce they're ready to test knowledge, a Summarizer might offer to explain the content."
            End If

            systemPrompt = $"{dateContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           $"A knowledge base has been loaded (it may contain multiple documents or sections).{personaContext} " &
                           "Ask what the user would like to discuss about the knowledge. One or two short sentences, stay in character."
        End If

        Dim answer = ""
        Try
            Dim sw = Stopwatch.StartNew()
            answer = Await CallLlmWithSelectedModelAsync(systemPrompt, "")
            sw.Stop()
        Catch ex As Exception
            answer = $"Good {partOfDay.ToLower()}! How can I help you today?"
        End Try

        answer = If(answer, "").Trim()
        AppendAssistantMarkdown(answer)
        ' Include welcome in history - it's part of the conversation
        _history.Add(("assistant", answer))
        PersistChatHtml()
    End Function

#End Region

#Region "Send Message"

    ''' <summary>
    ''' Builds the full prompt (persona, knowledge, history, document) and sends it to the LLM.
    ''' </summary>
    ''' <param name="userText">User's message text.</param>
    Private Async Function SendAsync(userText As String) As Task
        Try
            ' Build system prompt from persona or default
            Dim dateContext = GetDateContext()
            Dim randomWord = GetRandomModifier()

            Dim basePrompt = If(Not String.IsNullOrEmpty(_currentPersonaPrompt),
                                _currentPersonaPrompt,
                                $"You are {_currentPersonaName}, a helpful assistant. Discuss the provided knowledge with the user.")

            Dim systemPrompt = $"{basePrompt}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                               "The knowledge provided may consist of multiple documents or sections combined into one. " &
                               $"Refer to it as 'the knowledge' or 'the materials' rather than 'the document' when appropriate. {dateContext}"

            ' Build user prompt with knowledge and context
            Dim sb As New StringBuilder()

            sb.AppendLine("User message:")
            sb.AppendLine(userText)
            sb.AppendLine()

            ' Include full knowledge document without truncation for smaller docs
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                sb.AppendLine("<Knowledge Base>")
                Dim knowledgeText = _knowledgeContent
                sb.AppendLine(knowledgeText)
                sb.AppendLine("</Knowledge Base>")
                sb.AppendLine()
            End If

            ' Include active document if checkbox checked
            If _chkIncludeActiveDoc.Checked Then
                Dim activeDocContent = GetActiveDocumentContent()
                If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                    sb.AppendLine("<User's Active Document>")
                    sb.AppendLine(TruncateForPrompt(activeDocContent, _context.INI_ChatCap))
                    sb.AppendLine("</User's Active Document>")
                    sb.AppendLine()
                End If
            End If

            ' Include conversation history (excluding welcome messages)
            Dim convo = BuildConversationForLlm()
            If Not String.IsNullOrWhiteSpace(convo) Then
                sb.AppendLine("Conversation so far:")
                sb.AppendLine(convo)
            End If

            Dim sw = Stopwatch.StartNew()
            Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
            sw.Stop()

            answer = If(answer, "").Trim()

            RemoveAssistantThinking()
            AppendAssistantMarkdown(answer)
            _history.Add(("assistant", answer))
            PersistChatHtml()

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendAssistantMarkdown("*(Error: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
        End Try
    End Function

    ''' <summary>
    ''' Extracts the current Word document and selection details for prompt inclusion.
    ''' </summary>
    ''' <returns>Formatted string with document name, selected paragraph, and full text.</returns>
    Private Function GetActiveDocumentContent() As String
        Try
            Dim app = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents.Count = 0 Then Return ""

            Dim doc = app.ActiveDocument
            If doc Is Nothing Then Return ""

            Dim fullText = doc.Content.Text
            Dim selectedPara = ""

            ' Get current selection's paragraph
            Try
                Dim sel = app.Selection
                If sel IsNot Nothing AndAlso sel.Paragraphs.Count > 0 Then
                    selectedPara = sel.Paragraphs(1).Range.Text.Trim()
                End If
            Catch
            End Try

            Dim sb As New StringBuilder()
            sb.AppendLine($"Document: {doc.Name}")

            If Not String.IsNullOrWhiteSpace(selectedPara) Then
                sb.AppendLine()
                sb.AppendLine("Currently selected paragraph:")
                sb.AppendLine(selectedPara)
            End If

            sb.AppendLine()
            sb.AppendLine("Full document text:")
            sb.AppendLine(fullText)

            Return sb.ToString()

        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Limits long strings to the configured cap and annotates truncation.
    ''' </summary>
    ''' <param name="text">Input text.</param>
    ''' <param name="maxLen">Maximum length.</param>
    ''' <returns>Truncated or original text with annotation if truncated.</returns>
    Private Function TruncateForPrompt(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & vbCrLf & $"...[TRUNCATED - showing {maxLen:N0} of {text.Length:N0} characters]"
    End Function

#End Region

#Region "HTML Chat Display"

    ''' <summary>
    ''' Creates the base HTML document and CSS used by the WebBrowser control.
    ''' </summary>
    Private Sub InitializeChatHtml()
        Ui(Sub()
               _htmlQueue.Clear()
               _htmlReady = False
               Dim baseSize = If(Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
               Dim fontPt = Math.Max(CSng(baseSize + 1.0F), 10.0F)
               Dim css =
                   $"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
                    body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
                    #chat{{padding:8px;}}
                    .msg{{margin:8px 0;word-wrap:break-word;}}
                    .msg .who{{font-weight:600;margin-right:4px;}}
                    .msg.user{{background:#e8f4fc;border-left:3px solid #0078d4;padding:8px 10px;border-radius:4px;margin-right:40px;}}
                    .msg.user .who{{color:#0078d4;}}
                    .msg.assistant{{padding:8px 0;margin-left:0;}}
                    .msg.assistant .who{{color:#003366;}}
                    .msg.system{{color:#666;font-style:italic;background:#f9f9f9;padding:4px 8px;border-radius:4px;}}
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
    ''' Flushes queued HTML fragments once the browser document is ready.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub Chat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        _htmlReady = True
        If _htmlQueue.Count > 0 Then
            Try
                For Each frag In _htmlQueue
                    _chat.Document.InvokeScript("appendMessage", New Object() {frag})
                Next
            Catch
            Finally
                _htmlQueue.Clear()
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Intercepts navigation to open http/https/mailto links externally.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Navigation event arguments.</param>
    Private Sub Chat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            Dim scheme = e.Url?.Scheme?.ToLowerInvariant()
            If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                e.Cancel = True
                Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Prevents the WebBrowser control from spawning new windows.
    ''' </summary>
    ''' <param name="sender">Event source.</param>
    ''' <param name="e">Cancel event arguments.</param>
    Private Sub Chat_NewWindow(sender As Object, e As CancelEventArgs)
        e.Cancel = True
    End Sub

    ''' <summary>
    ''' Appends HTML to the chat DOM, queuing if the document is not ready.
    ''' </summary>
    ''' <param name="fragment">HTML fragment to append.</param>
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
    ''' Adds a user message block to the transcript and persists HTML.
    ''' </summary>
    ''' <param name="text">User message text.</param>
    Private Sub AppendUserHtml(text As String)
        Dim encoded = WebUtility.HtmlEncode(text).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>").Replace(vbCr, "<br>")
        AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Adds a system message block and persists HTML.
    ''' </summary>
    ''' <param name="text">System message text.</param>
    Private Sub AppendSystemMessage(text As String)
        Dim encoded = WebUtility.HtmlEncode(text)
        AppendHtml($"<div class='msg system'>{encoded}</div>")
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Inserts a temporary 'thinking' placeholder for the assistant.
    ''' </summary>
    Private Sub ShowAssistantThinking()
        _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
        AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{WebUtility.HtmlEncode(_currentPersonaName)}:</span><span class='content'>Thinking...</span></div>")
    End Sub

    ''' <summary>
    ''' Removes the current thinking placeholder if present.
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
    ''' Converts assistant markdown to HTML and appends it to the transcript.
    ''' </summary>
    ''' <param name="md">Markdown text from assistant.</param>
    Private Sub AppendAssistantMarkdown(md As String)
        md = If(md, "")
        Dim body = Markdig.Markdown.ToHtml(md, _mdPipeline)
        Dim t = body.Trim()
        Dim isSingle = Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                   Not Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

        Dim whoHtml = WebUtility.HtmlEncode(_currentPersonaName)

        If isSingle Then
            Dim inlineHtml = Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
            AppendHtml($"<div class='msg assistant'><span class='who'>{whoHtml}:</span><span class='content'>{inlineHtml}</span></div>")
        Else
            Dim m = Regex.Match(t, "^\s*<p>([\s\S]*?)</p>\s*", RegexOptions.IgnoreCase)
            If m.Success Then
                Dim firstInline = m.Groups(1).Value
                Dim rest = t.Substring(m.Index + m.Length).Trim()
                Dim sb As New StringBuilder()
                sb.Append("<div class='msg assistant'>")
                sb.Append("<span class='who'>").Append(whoHtml).Append(":</span>")
                sb.Append("<span class='content'>").Append(firstInline).Append("</span>")
                If rest.Length > 0 Then
                    sb.Append("<div class='content'>").Append(rest).Append("</div>")
                End If
                sb.Append("</div>")
                AppendHtml(sb.ToString())
            Else
                AppendHtml($"<div class='msg assistant'><span class='who'>{whoHtml}:</span><div class='content'>{t}</div></div>")
            End If
        End If
    End Sub

#End Region

#Region "Persistence"

    ''' <summary>
    ''' Saves the current chat DOM fragment to settings for restoration.
    ''' </summary>
    Private Sub PersistChatHtml()
        Ui(Sub()
               Try
                   If _chat.Document Is Nothing Then Return
                   Dim root = _chat.Document.GetElementById("chat")
                   If root Is Nothing Then Return
                   My.Settings.DiscussLastChatHtml = root.InnerHtml
                   My.Settings.Save()
               Catch
               End Try
           End Sub)
    End Sub

    ''' <summary>
    ''' Rebuilds the history list from the plain-text transcript copy.
    ''' </summary>
    ''' <param name="transcript">Plain text transcript to parse.</param>
    Private Sub RestoreHistoryFromTranscript(transcript As String)
        _history.Clear()
        If String.IsNullOrEmpty(transcript) Then Return

        Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split({vbLf}, StringSplitOptions.None)
        Dim currentRole As String = Nothing
        Dim content As New StringBuilder()

        Dim flush =
            Sub()
                If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                    content.Clear() : currentRole = Nothing : Return
                End If
                _history.Add((currentRole, content.ToString().Trim()))
                content.Clear()
                currentRole = Nothing
            End Sub

        For Each ln In lines
            If ln.StartsWith("You:", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "user"
                content.Append(ln.Substring(4).TrimStart())
            ElseIf ln.StartsWith(_currentPersonaName & ":", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "assistant"
                content.Append(ln.Substring((_currentPersonaName & ":").Length).TrimStart())
            ElseIf ln.StartsWith(AssistantName & ":", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "assistant"
                content.Append(ln.Substring((AssistantName & ":").Length).TrimStart())
            Else
                If content.Length > 0 Then content.AppendLine()
                content.Append(ln)
            End If
        Next
        flush()
    End Sub

    ''' <summary>
    ''' Recreates chat HTML from the stored transcript text.
    ''' </summary>
    ''' <param name="transcript">Plain text transcript to convert to HTML.</param>
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
            ElseIf ln.StartsWith(_currentPersonaName & ":", StringComparison.OrdinalIgnoreCase) Then
                flush() : currentRole = "assistant" : content.Append(ln.Substring((_currentPersonaName & ":").Length).TrimStart())
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
    ''' Truncates and saves the plain transcript respecting the configured cap.
    ''' </summary>
    Private Sub PersistTranscriptLimited()
        Dim transcript = BuildTranscriptPlain()
        Dim cap = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
        If transcript.Length > cap Then
            transcript = transcript.Substring(transcript.Length - cap)
        End If
        My.Settings.DiscussLastChat = transcript
    End Sub

    ''' <summary>
    ''' Returns the current chat history in 'You:/Persona:' text format.
    ''' </summary>
    ''' <returns>Plain text transcript of all messages.</returns>
    Private Function BuildTranscriptPlain() As String
        Dim sb As New StringBuilder()
        For Each m In _history
            If m.Role = "user" Then
                sb.AppendLine("You: " & m.Content)
            Else
                sb.AppendLine(_currentPersonaName & ": " & m.Content)
            End If
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Builds a capped, reversed conversation snippet for prompt injection.
    ''' </summary>
    ''' <returns>Truncated conversation history for LLM context.</returns>
    Private Function BuildConversationForLlm() As String
        Dim sb As New StringBuilder()
        Dim cap = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
        Dim acc = 0
        For i = _history.Count - 1 To 0 Step -1
            Dim line = If(_history(i).Role = "user", "User: ", _currentPersonaName & ": ") & _history(i).Content & Environment.NewLine
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

#End Region

#Region "Helpers"

    ''' <summary>
    ''' Determines 'Morning/Afternoon/Evening' from the current hour.
    ''' </summary>
    ''' <returns>Part of day string.</returns>
    Private Shared Function GetPartOfDay() As String
        Dim h = DateTime.Now.Hour
        If h < 12 Then Return "Morning"
        If h < 18 Then Return "Afternoon"
        Return "Evening"
    End Function

    ''' <summary>
    ''' Detects whether the restored HTML ended on an alternate-model state by checking for model switch messages.
    ''' </summary>
    ''' <param name="html">Saved HTML content from chat transcript.</param>
    ''' <returns>True if an alternate model was active when the chat was saved.</returns>
    Private Function ChatHtmlIndicatesAlternateModel(html As String) As Boolean
        If String.IsNullOrEmpty(html) Then Return False

        Try
            ' Look for the last occurrence of model switch messages
            Dim switchedToAlternateIdx = html.LastIndexOf("Switched to alternate model", StringComparison.OrdinalIgnoreCase)
            Dim switchedToSecondaryIdx = html.LastIndexOf("Switched to secondary model", StringComparison.OrdinalIgnoreCase)
            Dim switchedBackIdx = html.LastIndexOf("Switched back to primary model", StringComparison.OrdinalIgnoreCase)

            ' Find the latest "switched to" message
            Dim lastSwitchToIdx = Math.Max(switchedToAlternateIdx, switchedToSecondaryIdx)

            ' If there's no switch-to message, no alternate was active
            If lastSwitchToIdx < 0 Then Return False

            ' If there's no switch-back, or the switch-back is before the last switch-to, alternate was active
            If switchedBackIdx < 0 OrElse switchedBackIdx < lastSwitchToIdx Then
                Return True
            End If

            Return False
        Catch
            Return False
        End Try
    End Function

#End Region

End Class