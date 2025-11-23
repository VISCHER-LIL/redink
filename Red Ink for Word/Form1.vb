' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

' =============================================================================
' Form1.vb – Word Chat UI (Inky)  SHORT OVERVIEW
' =============================================================================
' Core purpose
'   Interactive chat assistant inside Word:
'     • Builds System + User prompts (can include active document, selection, and NOW other open docs)
'     • Sends to LLM via SharedLibrary; shows Markdown-rendered replies
'     • (Optionally) executes trusted bot commands on the active document
'     • Persists chat + window state; supports model switching (primary / second / alternate)
'
' Main UI elements
'   Text: txtChatHistory (hidden when HTML view active), txtUserInput
'   HTML chat: WebBrowser wbChat (+ Markdown → HTML via Markdig)
'   Buttons: btnSend, btnCopy, btnCopyLastAnswer, btnClear, btnExit, btnSwitchModel
'   Checkboxes:
'       chkIncludeDocText        – include full active document (final view, no deleted text)
'       chkIncludeselection      – include current selection OR cursor context
'       chkIncludeOtherDocs      – NEW: silently include all other open Word documents (excludes active)
'       chkPermitCommands        – allow write commands (requires doc or selection)
'       chkStayOnTop             – toggles topmost
'       chkConvertMarkdown       – format inserted/replied text and comments as Markdown
'
' Prompt construction (Send)
'   SystemPrompt adds capabilities lines conditionally:
'       active doc / selection / other open docs
'       command permission block (from INI templates)
'   User prompt body includes:
'       Active document text (if chosen)
'       Selection or cursor context
'       Other open docs (GatherSelectedDocuments Silent, IncludeName:=True, ExceptCurrent:=True)
'       Recent conversation (trimmed to INI_ChatCap)
'
' Other open documents (NEW)
'   When chkIncludeOtherDocs = True:
'       Calls: GatherSelectedDocuments(True, False, True, True)
'       Injects numbered <DOCUMENTn> blocks with document names into the prompt
'       Adds explanatory line to SystemPrompt so the model may reference them
'
' Model handling
'   btnSwitchModel: toggles or selects alternate model (snapshot + restore pattern)
'   UpdateDocumentCheckboxesState: disables all document/selection/other-doc checkboxes when second/alternate API active
'
' Command execution (optional)
'   Parses patterns: [#verb: @@arg1@@ §§arg2§§ #]
'   Supported verbs: find, replace, insert, insertafter, insertbefore, addcomment, replycomment
'   Uses chunked search (FindLongTextInChunks) for reliability with large text
'
' Markdown & HTML
'   Assistant replies → Markdown to HTML (links instrumented to open externally)
'   User input HTML-encoded only
'   “Thinking…” placeholder removed on response
'
' Persistence
'   Plain chat (LastChatHistory), HTML chat (LastChatHistoryHtml), window bounds, user checkbox prefs
'
' Safety / COM notes
'   Word view temporarily adjusted (final revisions view) for text extraction
'   Selection restored after operations; commands guarded by Esc abort key polling
'
' Extension points
'   • Add new verbs: extend ParseCommands + ExecuteAnyCommands Select Case
'   • Add new prompt flags: adjust SystemPrompt composition + UI checkbox
'
' =============================================================================

Public Class frmAIChat

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    Const AN As String = "Red Ink"
    Const AN5 As String = "Inky"   ' for Chatbox
    Const AN6 As String = "RI"

    Const MarkerChar As String = ChrW(&HE000)
    Const CursorPositionCount As Integer = 25

    Private PreceedingNewline As String = ""
    Private OldChat As String = ""
    Private UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
    Private SystemPrompt As String = ""

    ' Tracks a user-chosen alternate model for temporary use per-call.
    Private _alternateModelSelected As Boolean = False
    Private _alternateModelConfig As ModelConfig = Nothing
    Private _alternateModelDisplayName As String = Nothing

    Private WithEvents btnCopy As New Button() With {.Text = "Copy All", .AutoSize = True}
    Private WithEvents btnCopyLastAnswer As New Button() With {.Text = "Copy Last Answer", .AutoSize = True}
    Private WithEvents btnClear As New Button() With {.Text = "Clear", .AutoSize = True}
    Private WithEvents btnExit As New Button() With {.Text = "Close", .AutoSize = True}
    Private WithEvents btnSend As New Button() With {.Text = "Send", .AutoSize = True}
    Private WithEvents btnSwitchModel As New Button() With {.Text = "Switch Model", .AutoSize = True}
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {.Text = "Include document", .AutoSize = True, .Checked = My.Settings.IncludeDocument}
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {.Text = "Include selection", .AutoSize = True, .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)}
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {.Text = "Grant write access", .AutoSize = True, .Checked = My.Settings.DoCommands}
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {.Text = "Not always on top", .AutoSize = True, .Checked = My.Settings.NotAlwaysOnTop}
    Private WithEvents chkConvertMarkdown As New System.Windows.Forms.CheckBox() With {.Text = "Do format", .AutoSize = True, .Checked = My.Settings.ConvertMarkdownInChat}
    Private WithEvents chkIncludeOtherDocs As New System.Windows.Forms.CheckBox() With {.Text = "Include all other open Word docs", .AutoSize = True, .Checked = False}


    Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    Dim pnlCheckboxes As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    Private _context As ISharedContext = New SharedContext()

    ' Tracks whether we are using the second model/API.
    Private _useSecondApi As Boolean = False

    ' We keep the entire conversation in a List of (role, content).
    Private _chatHistory As New List(Of (Role As String, Content As String))



    Public Sub New(context As ISharedContext)
        ' This call is required by the designer.
        InitializeComponent()

        Me.AutoSize = False

        txtChatHistory.Multiline = True
        txtUserInput.Multiline = True

        ' 1) TableLayoutPanel anlegen
        Dim mainLayout As New TableLayoutPanel() With {
        .ColumnCount = 1,
        .RowCount = 5,
        .Dock = DockStyle.Fill,
        .AutoSize = False,
        .Padding = New Padding(10)   ' wird gleich überschrieben
    }

        ' 2) Spalten‑Breite auf 100 % setzen
        mainLayout.ColumnStyles.Clear()
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        ' 3) Rechts 20 px Innenabstand
        mainLayout.Padding = New Padding(left:=10, top:=10, right:=20, bottom:=10)

        ' 4) Zeilen definieren
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' 5) Controls konfigurieren
        lblInstructions.AutoSize = True
        lblInstructions.Dock = DockStyle.Top
        txtChatHistory.Dock = DockStyle.Fill
        txtUserInput.Dock = DockStyle.Fill

        ' 6) Controls in die Tabelle packen
        mainLayout.Controls.Add(lblInstructions, 0, 0)
        mainLayout.Controls.Add(txtChatHistory, 0, 1)
        mainLayout.Controls.Add(txtUserInput, 0, 2)
        mainLayout.Controls.Add(pnlCheckboxes, 0, 3)
        mainLayout.Controls.Add(pnlButtons, 0, 4)

        InitChatHtmlUI(mainLayout)

        ' 7) Form neu befüllen
        Me.Controls.Clear()
        Me.Controls.Add(mainLayout)

        _context = context
    End Sub



    ' Runs once when form loads.
    Private Async Sub frmAIChat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.StartPosition = FormStartPosition.Manual
        Me.KeyPreview = True

        ' Restore saved chat text from My.Settings
        Dim previousChat As String = My.Settings.LastChatHistory
        If Not String.IsNullOrEmpty(previousChat) Then
            txtChatHistory.Text = previousChat
            OldChat = previousChat
            PreceedingNewline = Environment.NewLine
        End If

        ' Initialize the HTML chat view and (optionally) show previous transcript as preformatted text
        InitializeChatHtml()

        ' Prefer restoring from stored HTML; otherwise fallback to plain transcript.
        Dim previousChatHtml As String = My.Settings.LastChatHistoryHtml
        Dim hasExistingChat As Boolean = False

        If Not String.IsNullOrEmpty(previousChatHtml) Then
            ' Append as one fragment so wireLinks runs and events attach
            AppendHtml(previousChatHtml)
            hasExistingChat = True
        ElseIf Not String.IsNullOrEmpty(previousChat) Then
            AppendTranscriptToHtml(previousChat)
            hasExistingChat = True
        End If

        ' Set basic form props
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
        Me.TopMost = True
        Me.MinimumSize = New Size(830, 521)

        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If

        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown

        lblInstructions.Text = $"Enter your question and Enter (or 'Send'). You can allow the chatbot to do actions on your document (search, replace, delete, insert text and add or reply to comments). It does not see deletions, markups as such or formatting."
        lblInstructions.AutoSize = True
        lblInstructions.Height = 50
        lblInstructions.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblInstructions.TextAlign = ContentAlignment.MiddleLeft

        ' FlowLayoutPanel for buttons
        pnlButtons.Padding = New Padding(0, 2, 8, 12)
        pnlButtons.Controls.Add(btnSend)
        pnlButtons.Controls.Add(btnCopyLastAnswer)
        pnlButtons.Controls.Add(btnCopy)
        pnlButtons.Controls.Add(btnClear)

        ' Show the model button if either second API is configured or an alternate INI exists
        If _context.INI_SecondAPI OrElse Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            UpdateModelButtonText()
            pnlButtons.Controls.Add(btnSwitchModel)
        End If

        pnlButtons.Controls.Add(btnExit)

        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkStayOnTop)
        pnlCheckboxes.Controls.Add(chkConvertMarkdown)
        pnlCheckboxes.Controls.Add(chkIncludeOtherDocs)


        AddHandler btnCopy.Click, AddressOf btnCopy_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click
        AddHandler btnSend.Click, AddressOf btnSend_Click
        AddHandler btnCopyLastAnswer.Click, AddressOf btnCopyLastAnswer_Click
        AddHandler btnSwitchModel.Click, AddressOf btnSwitchModel_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click
        AddHandler chkIncludeselection.Click, AddressOf chkIncludeselection_Click
        AddHandler chkIncludeDocText.Click, AddressOf chkIncludeDocText_Click
        AddHandler chkPermitCommands.Click, AddressOf chkPermitCommands_Click
        AddHandler chkStayOnTop.Click, AddressOf chkStayontop_Click
        AddHandler chkConvertMarkdown.Click, AddressOf chkConvertMarkdown_Click

        ' Set the window title after all fields are ready
        UpdateTitle()

        If hasExistingChat Then
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()
        Else
            Dim result = Await WelcomeMessage()
        End If

        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
    End Sub

    ' Centralized title updater: shows primary/second/alternate model name.
    Private Sub UpdateTitle()
        Dim titleModel As String
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) AndAlso _alternateModelSelected AndAlso Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
            titleModel = _alternateModelDisplayName
        Else
            titleModel = If(_useSecondApi, _context.INI_Model_2, _context.INI_Model)
        End If
        Me.Text = $"Chat (using {titleModel})"
    End Sub

    ' Executes an LLM call while temporarily applying any selected alternate model.
    ' Always restores the original config afterwards.
    Private Async Function CallLlmWithSelectedModelAsync(systemPrompt As String, fullPrompt As String) As Task(Of String)
        Dim backupConfig As ModelConfig = Nothing
        Dim appliedAlternate As Boolean = False

        Try
            ' If the user selected an alternate model, apply it to the context as the "second model" just for this call.
            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                ' Back up current config (the "original state at rest")
                backupConfig = SharedMethods.GetCurrentConfig(_context)

                ' Apply the selected alternate config
                SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                appliedAlternate = True

                ' Enforce second API usage for alternate models
                _useSecondApi = True
            End If

            ' Execute the LLM call
            Return Await SharedMethods.LLM(_context, systemPrompt, fullPrompt, "", "", 0, _useSecondApi, True)

        Finally
            ' Always restore the original config after the call so the rest of the add-in sees the original state.
            If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                SharedMethods.RestoreDefaults(_context, backupConfig)
            End If
        End Try
    End Function

    ' When the user clicks Send, we call the LLM with context.
    ' Then append the AI response to the conversation.

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Dim errorOccurred As Boolean = False
        Dim errorMessage As String = ""

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland." & If(chkIncludeDocText.Checked, "\nYou have access to the user's document. \n", "") & If(chkIncludeselection.Checked, "\nYou have access to a selection of user's document. \n ", "") & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked), _context.SP_Add_ChatWord_Commands, _context.SP_Add_Chat_NoCommands)

            SystemPrompt = _context.SP_ChatWord().
                        Replace("{UserLanguage}", UserLanguage) &
                        $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}." &
                        If(chkIncludeDocText.Checked, vbLf & "You have access to the user's active document." & vbLf, "") &
                        If(chkIncludeselection.Checked, vbLf & "You have access to a selection of the active document." & vbLf, "") &
                        If(chkIncludeOtherDocs.Checked, vbLf & "You also have access to all other open Word documents (the user's request may refer to them)." & vbLf, "") &
                        If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked),
                           _context.SP_Add_ChatWord_Commands,
                           _context.SP_Add_Chat_NoCommands)

            Dim conversationSoFar As String = BuildConversationString(_chatHistory)
            If Not String.IsNullOrWhiteSpace(OldChat) Then
                conversationSoFar += "\n" & OldChat
                OldChat = ""
            End If

            Dim appGuard As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If (chkIncludeDocText.Checked Or chkIncludeselection.Checked) AndAlso
           (appGuard Is Nothing _
            OrElse appGuard.Documents Is Nothing _
            OrElse appGuard.Documents.Count = 0 _
            OrElse appGuard.ActiveDocument Is Nothing _
            OrElse appGuard.ActiveWindow Is Nothing) Then

                ShowCustomMessageBox("There is no active Word document. Please open or activate a document, then try again.")
                Return
            End If

            ' Optionally include Word document text or selection
            Dim docText As String = If(chkIncludeDocText.Checked, GetActiveDocumentText(), "")
            Dim selectionText As String = If(chkIncludeselection.Checked Or chkIncludeDocText.Checked, GetCurrentSelectionText(), "")

            ' If full document is included but no selection, get cursor context
            Dim sel As Microsoft.Office.Interop.Word.Selection = Globals.ThisAddIn.Application.Selection
            If sel IsNot Nothing AndAlso sel.Start = sel.End Then
                selectionText = GetCursorContext(CursorPositionCount)
            End If

            ' Gather other open documents silently if toggle selected
            Dim otherDocs As String = ""
            If chkIncludeOtherDocs.Checked Then
                otherDocs = Globals.ThisAddIn.GatherSelectedDocuments(
                    IncludeName:=True,
                    IncludeNone:=False,
                    ExceptCurrent:=True,
                    SilentAndGetAll:=True)
            End If

            ' Construct the full prompt
            Dim fullPrompt As New StringBuilder()

            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine("The user's document has the name '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' and has the following content: '" & docText & "'")
            End If
            If Not String.IsNullOrEmpty(selectionText) Then
                If chkIncludeDocText.Checked AndAlso sel.Start = sel.End Then
                    fullPrompt.AppendLine("In the user's document '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' the cursor is currently positioned in the following context: '" & selectionText & "'")
                Else
                    fullPrompt.AppendLine("In the user's document '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' the user has selected the following text: '" & selectionText & "'")
                End If
            End If

            If chkIncludeOtherDocs.Checked AndAlso
               Not String.IsNullOrEmpty(otherDocs) AndAlso
               Not otherDocs.Equals("NONE", StringComparison.OrdinalIgnoreCase) AndAlso
               Not otherDocs.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then

                fullPrompt.AppendLine("The following are the other open Word documents (each enclosed in <DOCUMENTn> tags, including their name so you can refer to them):")
                fullPrompt.AppendLine(otherDocs)
            End If

            fullPrompt.AppendLine("User: " & userPrompt)
            fullPrompt.AppendLine("The conversation so far (not including any previously added text document):\n" & conversationSoFar)

            Debug.WriteLine("Document=" & Globals.ThisAddIn.Application.ActiveDocument.Name)
            Debug.WriteLine(fullPrompt.ToString())

            ' Update UI on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(PreceedingNewline & "You: " & userPrompt.TrimEnd() & Environment.NewLine & Environment.NewLine)
                                    txtUserInput.Clear()
                                    PreceedingNewline = Environment.NewLine
                                End Sub)

            Await UpdateUIAsync(Sub()
                                    AppendUserHtml(userPrompt.TrimEnd())
                                End Sub)

            _chatHistory.Add(("user", userPrompt.TrimEnd()))

            ' Add a placeholder for AI response while waiting
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory($"{AN5}: Thinking...")
                                End Sub)

            Await UpdateUIAsync(Sub()
                                    ShowAssistantThinking()
                                End Sub)

            ' Call the LLM function asynchronously
            Dim aiResponseOriginal As String = Await CallLlmWithSelectedModelAsync(SystemPrompt, fullPrompt.ToString())

            ' Keep original Markdown for HTML rendering
            Dim aiResponseMd As String = (If(aiResponseOriginal, "")).TrimEnd()

            ' Maintain your existing plain-text pipeline for persistence/commands
            Dim aiResponsePlain As String = aiResponseMd
            aiResponsePlain = aiResponsePlain.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ").Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ").Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
            aiResponsePlain = aiResponsePlain.Replace($"  *  ", "  " & ChrW(8226) & "  ")
            aiResponsePlain = RemoveMarkdownFormatting(aiResponsePlain)

            Dim CommandsString As String = ""
            If My.Settings.DoCommands And (chkIncludeselection.Checked Or chkIncludeDocText.Checked) Then
                CommandsString = aiResponsePlain
                aiResponsePlain = RemoveCommands(aiResponsePlain)
                aiResponsePlain = Regex.Replace(aiResponsePlain, "[\r\n\s]+$", "")
            End If

            ' Remove commands from the Markdown we display to the user (HTML)
            Dim aiResponseMdDisplay As String = RemoveCommands(aiResponseMd)
            aiResponseMdDisplay = Regex.Replace(aiResponseMdDisplay, "[\r\n\s]+$", "")

            Debug.WriteLine("AI response: " & CommandsString)

            Await UpdateUIAsync(Sub()
                                    ' Remove "Thinking..." in text and HTML
                                    RemoveLastLineFromChatHistory()
                                    RemoveAssistantThinking()

                                    ' Append assistant answer to text transcript (plain)
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponsePlain.TrimStart().TrimEnd().Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine) & Environment.NewLine)

                                    ' Append assistant answer as Markdown-rendered HTML (commands filtered)
                                    AppendAssistantMarkdown(aiResponseMdDisplay.TrimStart())

                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        Try
                                            ExecuteAnyCommands(CommandsString, chkIncludeselection.Checked)
                                        Catch cmdEx As Exception
                                            ' Report command execution error to chat
                                            ReportCommandExecutionError(cmdEx.Message)
                                        End Try
                                    End If
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
                                End Sub)

            _chatHistory.Add(("assistant", aiResponsePlain.TrimEnd()))

        Catch ex As System.Exception
            ' Just capture the error, don't do async work here
            errorOccurred = True
            errorMessage = $"Error processing request: {ex.Message}"
        End Try

        ' Handle the error outside the Try-Catch if it occurred
        If errorOccurred Then
            Await UpdateUIAsync(Sub()
                                    ReportCommandExecutionError(errorMessage)
                                    txtUserInput.Text = userPrompt ' Restore user input so they can try again
                                End Sub)
        End If

    End Sub

    Private Sub ReportCommandExecutionError(errorMessage As String)
        If String.IsNullOrWhiteSpace(errorMessage) Then Return

        Dim errorText As String = $"⚠ Error: {errorMessage}"

        ' Add to plain text chat history
        AppendToChatHistory(Environment.NewLine & "─────────────────────────────────────" & Environment.NewLine)
        AppendToChatHistory(errorText & Environment.NewLine)
        AppendToChatHistory("─────────────────────────────────────" & Environment.NewLine)

        ' Add to HTML chat with formatting
        Dim htmlError As String = $"<div class='msg assistant error' style='border-left: 3px solid #ff9800; padding-left: 10px; margin: 10px 0; background-color: #fff3e0;'>
            <span class='who' style='color: #ff9800;'>System:</span>
            <div class='content'>
                <hr style='border: none; border-top: 1px solid #ff9800; margin: 8px 0;' />
                <strong>⚠ {HtmlEncode(errorMessage)}</strong>
                <hr style='border: none; border-top: 1px solid #ff9800; margin: 8px 0;' />
            </div>
        </div>"

        AppendHtml(htmlError)
        PersistChatHtml()

        ' Add to chat history so AI can see the error
        _chatHistory.Add(("assistant", $"System Error: {errorMessage}"))
    End Sub

    ' Gets context around the cursor position (characters before and after) with cursor position marker
    Private Function GetCursorContext(charCount As Integer) As String
        Try
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            ' Check if there's an actual selection (not just cursor position)
            If Not String.IsNullOrEmpty(sel.Text) AndAlso sel.Start <> sel.End Then
                Return ""
            End If

            Dim cursorPos As Integer = sel.Start
            Dim docStart As Integer = activeDoc.Content.Start
            Dim docEnd As Integer = activeDoc.Content.End

            ' Calculate the range boundaries
            Dim contextStart As Integer = Math.Max(docStart, cursorPos - charCount)
            Dim contextEnd As Integer = Math.Min(docEnd, cursorPos + charCount)

            ' Get text before cursor
            Dim beforeRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(contextStart, cursorPos)
            Dim textBefore As String = beforeRange.Text

            ' Get text after cursor
            Dim afterRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(cursorPos, contextEnd)
            Dim textAfter As String = afterRange.Text

            ' Combine with cursor marker
            Dim contextText As String = textBefore & "[cursor is here]" & textAfter

            ' Try to extract bubbles/comments if available for the entire context range
            Dim bubbles As String = ""
            Try
                Dim fullContextRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(contextStart, contextEnd)
                bubbles = ThisAddIn.BubblesExtract(fullContextRange, True) ' Silent=True
            Catch
                ' ignore and keep contextText
            End Try

            If Not String.IsNullOrEmpty(bubbles) Then
                Return contextText & " " & bubbles
            End If

            Return contextText

        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Async Function WelcomeMessage() As Task(Of String)

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}."
            txtUserInput.Text = ""

            ' Call the LLM function asynchronously
            Dim aiResponseRaw As String = Await CallLlmWithSelectedModelAsync(SystemPrompt, $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage} , such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.")

            ' Keep Markdown for HTML display (filter bot-commands if any)
            Dim aiDisplayMd As String = RemoveCommands(If(aiResponseRaw, ""))

            ' Maintain your existing plain text behavior for the transcript
            Dim aiResponseTxt As String = If(aiResponseRaw, "")
            aiResponseTxt = aiResponseTxt.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf
            aiResponseTxt = aiResponseTxt.Replace("**", "").Replace("_", "").Replace("`", "")

            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponseTxt.Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine))
                                    ' Also show the formatted version in the HTML chat
                                    AppendAssistantMarkdown(aiDisplayMd)
                                End Sub)

            _chatHistory.Add(("assistant", aiResponseTxt))

            PreceedingNewline = Environment.NewLine

            Return ""

        Catch ex As System.Exception
            Return ""
        End Try
    End Function

    Private Function ConvertHtmlToPlainText(html As String) As String
        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)
        Return doc.DocumentNode.InnerText
    End Function

    ' Helper method to ensure UI updates occur on the correct thread
    Private Async Function UpdateUIAsync(action As System.Action) As System.Threading.Tasks.Task
        If InvokeRequired Then
            Await System.Threading.Tasks.Task.Run(Sub() Me.Invoke(action))
        Else
            action()
        End If
    End Function


    Private Sub AppendToChatHistory(text As String)
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() txtChatHistory.AppendText(text))
        Else
            txtChatHistory.AppendText(text)
        End If
    End Sub

    Private Sub RemoveLastLineFromChatHistory()
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() RemoveLastLineFromChatHistory())
        Else
            Dim lines As String() = txtChatHistory.Lines
            If lines.Length > 0 Then
                txtChatHistory.Lines = lines.Take(lines.Length - 1).ToArray()
            End If
        End If
    End Sub

    Private Sub chkStayontop_Click(sender As Object, e As EventArgs)
        Me.TopMost = Not Me.TopMost
        My.Settings.NotAlwaysOnTop = Me.TopMost
        My.Settings.Save()
    End Sub

    Private Sub chkConvertMarkdown_Click(sender As Object, e As EventArgs)
        My.Settings.ConvertMarkdownInChat = chkConvertMarkdown.Checked
        My.Settings.Save()
    End Sub


    Private Sub chkPermitCommands_Click(sender As Object, e As EventArgs)
        My.Settings.DoCommands = Not My.Settings.DoCommands

        If My.Settings.DoCommands And Not chkIncludeselection.Checked Then
            chkIncludeDocText.Checked = True
            My.Settings.IncludeDocument = chkIncludeDocText.Checked
        End If

        My.Settings.Save()
    End Sub


    Private Sub chkIncludeselection_Click(sender As Object, e As EventArgs)
        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Get the selection from the active window
        Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

        If String.IsNullOrWhiteSpace(sel.Text) Then
            chkIncludeselection.Checked = False
        ElseIf chkIncludeDocText.Checked Then
            chkIncludeDocText.Checked = False
        End If
        My.Settings.IncludeSelection = chkIncludeselection.Checked

        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub

    Private Sub chkIncludeDocText_Click(sender As Object, e As EventArgs)
        If chkIncludeselection.Checked Then
            chkIncludeselection.Checked = False
        End If
        My.Settings.IncludeDocument = chkIncludeDocText.Checked

        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub


    ' Copies the entire conversation to the clipboard.

    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        My.Computer.Clipboard.SetText(txtChatHistory.Text)
    End Sub


    ' Copies only the last AI answer to the clipboard.

    Private Sub btnCopyLastAnswer_Click(sender As Object, e As EventArgs)
        Dim lastAssistantMsg = _chatHistory.Where(Function(x) x.Role = "assistant").LastOrDefault()
        If lastAssistantMsg.Content IsNot Nothing Then
            My.Computer.Clipboard.SetText(lastAssistantMsg.Content)
        Else
            SharedMethods.ShowCustomMessageBox("No last AI answer available.")
        End If
    End Sub


    ' Switches the model from model1 to model2 and vice versa.

    ' Select/toggle model. When Alternate INI exists, capture the alternate config and
    ' immediately restore the original config to keep globals pristine between calls.
    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            ' If an alternate is already active -> switch back to primary without dialog
            If _alternateModelSelected Then
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
                _useSecondApi = False
                UpdateModelButtonText()
                UpdateTitle()
                UpdateDocumentCheckboxesState()
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
                _useSecondApi = True
            Else
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
                _useSecondApi = False
            End If

            UpdateModelButtonText()
            UpdateTitle()
            UpdateDocumentCheckboxesState()
        Else
            ' Legacy behavior: simple toggle between primary and configured second model
            _useSecondApi = Not _useSecondApi
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing
            UpdateModelButtonText()
            UpdateTitle()
            UpdateDocumentCheckboxesState()
        End If
    End Sub

    ' Sets the model button text depending on the current state and availability of alternates.
    Private Sub UpdateModelButtonText()
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            btnSwitchModel.Text = If(_alternateModelSelected, "Primary model", "Alternate Model")
        Else
            btnSwitchModel.Text = "Switch Model"
        End If
    End Sub


    ' Disables document/selection checkboxes when using second API models
    Private Sub UpdateDocumentCheckboxesState()
        If _useSecondApi Then
            ' Uncheck document-related checkboxes

            chkIncludeDocText.Checked = False
            chkIncludeselection.Checked = False
            chkPermitCommands.Checked = False
            chkIncludeOtherDocs.Checked = False

            ' Update settings
            My.Settings.IncludeDocument = False
            My.Settings.IncludeSelection = False
            My.Settings.DoCommands = False
            My.Settings.Save()
        Else
            ' Re-enable checkboxes when switching back to primary model
            chkIncludeDocText.Enabled = True
            chkIncludeselection.Enabled = True
            chkPermitCommands.Enabled = True

            ' Optionally restore previous settings (or leave unchecked)
            ' chkIncludeDocText.Checked = My.Settings.IncludeDocument
            ' chkIncludeselection.Checked = My.Settings.IncludeSelection
            ' chkPermitCommands.Checked = My.Settings.DoCommands
        End If
    End Sub

    ' Clears the conversation from both the UI and saved settings.
    Private Async Sub btnClear_Click(sender As Object, e As EventArgs)

        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.LastChatHistoryHtml = ""
        My.Settings.Save()

        ClearChatHtml()

        Await WelcomeMessage()
    End Sub

    ' Press Escape to close. Also button-based exit.

    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If
            My.Settings.LastChatHistory = conversation
            PersistChatHtml()
            My.Settings.Save()
            Close()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation
        PersistChatHtml()
        My.Settings.Save()
        Close()
    End Sub

    Private Sub frmAIChat_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Save the chat history before the form closes
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation

        ' Save the form's location and size to My.Settings
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.FormLocation = Me.Location
            My.Settings.FormSize = Me.Size
        Else
            ' If the form is minimized or maximized, save the restored bounds
            My.Settings.FormLocation = Me.RestoreBounds.Location
            My.Settings.FormSize = Me.RestoreBounds.Size
        End If
        PersistChatHtml()
        My.Settings.Save()

    End Sub


    ' Trigger the Send button on Ctrl+Enter in the user input textbox.

    Private Sub oldUserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.Enter Then
            btnSend.PerformClick()
            e.Handled = True
        End If
    End Sub

    ' Trigger the Send button on Enter, allow Shift+Enter for new line
    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If e.Shift Then
                ' Allow Shift+Enter to insert a new line (default behavior)
                Return
            Else
                ' Enter alone sends the message
                e.SuppressKeyPress = True
                btnSend.PerformClick()
                e.Handled = True
            End If
        End If
    End Sub

    ' Reads the entire document's text.

    Private Function GetActiveDocumentText() As String
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Save current view settings
            Dim originalRevisionsView As Word.WdRevisionsView = wordApp.ActiveWindow.View.RevisionsView
            Dim originalShowRevisions As Boolean = wordApp.ActiveWindow.View.ShowRevisionsAndComments

            Try
                ' Temporarily show only final text (no tracked changes)
                With wordApp.ActiveWindow.View
                    .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                    .ShowRevisionsAndComments = False
                End With

                ' Now get the text (will exclude deleted content)
                Dim baseText As String = doc.Content.Text

                Dim bubbles As String = ""
                Try
                    bubbles = ThisAddIn.BubblesExtract(doc.Content, True) ' Silent=True
                Catch
                    ' ignore and keep baseText
                End Try

                If Not String.IsNullOrEmpty(bubbles) Then
                    Return baseText & vbCr & vbCr & bubbles
                End If

                Return baseText

            Finally
                ' Restore original view settings
                With wordApp.ActiveWindow.View
                    .RevisionsView = originalRevisionsView
                    .ShowRevisionsAndComments = originalShowRevisions
                End With
            End Try

        Catch ex As Exception
            Return ""
        End Try
    End Function

    ' Reads the current selection's text.
    Private Function GetCurrentSelectionText() As String
        Try
            ' Get the active document
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Get the selection from the active window
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            If String.IsNullOrEmpty(sel.Text) Then
                chkIncludeselection.Checked = False
                Return ""
            Else
                ' Save current view settings
                Dim originalRevisionsView As Word.WdRevisionsView = wordApp.ActiveWindow.View.RevisionsView
                Dim originalShowRevisions As Boolean = wordApp.ActiveWindow.View.ShowRevisionsAndComments

                Try
                    ' Temporarily show only final text (no tracked changes)
                    With wordApp.ActiveWindow.View
                        .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = False
                    End With

                    ' Now get the text (will exclude deleted content)
                    Dim baseText As String = sel.Text

                    Dim bubbles As String = ""
                    Try
                        bubbles = ThisAddIn.BubblesExtract(sel.Range, True) ' Silent=True
                    Catch
                        ' ignore and keep baseText
                    End Try

                    If Not String.IsNullOrEmpty(bubbles) Then
                        Return baseText & " " & bubbles
                    End If

                    Return baseText

                Finally
                    ' Restore original view settings
                    With wordApp.ActiveWindow.View
                        .RevisionsView = originalRevisionsView
                        .ShowRevisionsAndComments = originalShowRevisions
                    End With
                End Try
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ' Builds the conversation history as a single string.

    Private Function BuildConversationString(history As List(Of (Role As String, Content As String))) As String
        Dim sb As New StringBuilder()
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = _context.INI_ChatCap

        ' Iterate through the history in reverse order (most recent messages first)
        For Each msg In history.AsEnumerable().Reverse()
            Dim message As String
            If msg.Role = "user" Then
                message = $"User: {msg.Content}{Environment.NewLine}"
            Else
                message = $"{AN5}: {msg.Content}{Environment.NewLine}"
            End If

            ' Check if adding this message will exceed the limit
            If totalLength + message.Length > maxLength Then
                ' If so, truncate the message to fit within the limit
                Dim remainingLength As Integer = maxLength - totalLength
                If remainingLength > 0 Then
                    sb.Insert(0, message.Substring(0, remainingLength))
                End If
                Exit For
            Else
                ' Otherwise, append the full message
                sb.Insert(0, message)
                totalLength += message.Length
            End If
        Next

        Return sb.ToString()
    End Function


    Private Sub pnlCheckboxes_Paint(sender As Object, e As PaintEventArgs)

    End Sub


    Private Function DecodeParagraphMarks(raw As String) As String
        If String.IsNullOrEmpty(raw) Then Return ""

        ' 1. Unify actual control characters first
        raw = raw.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)

        ' 2. Word Find tokens → paragraph
        raw = Regex.Replace(raw, "\^p", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "\^0*13", vbCr, RegexOptions.IgnoreCase)

        ' 3. Convert literal (escaped) sequences coming from LLM output:
        '    - \r\n  → single paragraph break (treat as one)
        '    - \n    → paragraph
        '    - \r    → paragraph
        '    Only when NOT double-escaped (i.e. ignore \\r, \\n).
        raw = Regex.Replace(raw, "(?<!\\)\\r\\n", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "(?<!\\)\\r", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "(?<!\\)\\n", vbCr, RegexOptions.IgnoreCase)

        ' 4. (Optional) Collapse any accidental multiple consecutive paragraphs caused by mixed encodings
        '    Comment out if you intentionally need empties:
        ' raw = Regex.Replace(raw, vbCr & "{2,}", vbCr & vbCr)

        Return raw
    End Function

    Private Function EnsureParagraphs(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return DecodeParagraphMarks(text)
    End Function

    Private Function CleanArgument(arg As String) As String
        If arg Is Nothing Then Return ""
        arg = DecodeParagraphMarks(arg)
        ' Trim but keep leading/trailing paragraph marks if they were intentional:
        ' Only trim spaces/tabs.
        Return Regex.Replace(arg, "^[ \t]+|[ \t]+$", "")
    End Function

    Public Class ParsedCommand
        Public Property Command As String
        Public Property Argument1 As String
        Public Property Argument2 As String
    End Class

    ' Parses the input string for embedded commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' Returns a List of ParsedCommand objects.
    ' argument2 is optional; if not present, it defaults to "".
    Private Function ParseCommands(input As String) As List(Of ParsedCommand)
        Dim results As New List(Of ParsedCommand)
        Try
            ' ------------------------------------------------------------------------------
            ' REGEX PATTERN to parse blocks shaped like
            '     [#cmd:@@arg1@@ §§arg2§§#]
            '
            '   \[#(?<cmd>[^:]+):\s*@@(?<arg1>(?:[^@]|@(?!@))*?)@@\s*
            '     (?:§§(?<arg2>(?:[^§]|§(?!§))*?)§§)?\s*#\]
            '
            ' EXPLANATION (left-to-right)
            ' ------------------------------------------------------------------------------
            ' \[#                       – literal “[#” (opens the block)
            '
            ' (?<cmd>[^:]+)             – named group  cmd
            '                              • one or more characters, anything except “:”
            '                              • therefore ends exactly at the first colon
            '
            ' :\s*@@                    – literal “:” plus optional whitespace,
            '                              followed by **exactly two** @ (start delimiter
            '                              for arg1)
            '
            ' (?<arg1>(?:[^@]|@(?!@))*?) – named group  arg1
            '                              • any character sequence
            '                              • a single @ is allowed
            '                              • **stops only** at a double @@
            '                                (tempered-greedy token  @(?!@) )
            '
            ' @@\s*                     – end delimiter for arg1 (double @) plus
            '                              optional whitespace
            '
            ' (?:                       – ── optional arg2 block ──
            '     §§
            '     (?<arg2>(?:[^§]|§(?!§))*?)
            '                            – named group  arg2
            '                              • any character sequence
            '                              • a single § is allowed
            '                              • **stops only** at a double §§
            '     §§
            ' )?                        – end of optional arg2 block
            '
            ' \s*#\]                    – optional whitespace, literal “#]”
            '                              (closes the entire block)
            ' ------------------------------------------------------------------------------
            ' Notes:
            ' • Single @ or § inside the arguments are allowed; only **double** @@ or §§
            '   terminate the corresponding argument.
            ' • You can change the delimiters if needed—just keep the same “tempered
            '   greedy token” logic so the inner data remains safe.
            ' ------------------------------------------------------------------------------
            Dim pattern As String = "\[#(?<cmd>[^:]+):\s*@@(?<arg1>(?:[^@]|@(?!@))*?)@@\s*(?:§§(?<arg2>(?:[^§]|§(?!§))*?)§§)?\s*#\]"
            Dim regex As New Regex(pattern, RegexOptions.Singleline)

            For Each m As Match In regex.Matches(input)
                Dim pc As New ParsedCommand()
                pc.Command = m.Groups("cmd").Value.Trim()

                Dim raw1 As String = m.Groups("arg1").Value
                Dim raw2 As String = If(m.Groups("arg2") IsNot Nothing, m.Groups("arg2").Value, "")

                pc.Argument1 = CleanArgument(raw1)
                pc.Argument2 = CleanArgument(raw2)

                ' If REPLACE (any case) and no Argument2 -> treat as delete (keep arg2 empty)
                ' (No extra transformation needed now.)
                If Not results.Any(Function(x) x.Command.Equals(pc.Command, StringComparison.OrdinalIgnoreCase) _
                                        AndAlso x.Argument1 = pc.Argument1 AndAlso x.Argument2 = pc.Argument2) Then
                    results.Add(pc)
                End If
            Next
        Catch ex As Exception
            MsgBox("Error in ParseCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try
        Return results
    End Function


    ' Removes all commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' from the input string.
    Public Function RemoveCommands(input As String) As String
        Dim output As String = input
        Try
            ' Remove the commands along with immediate surrounding whitespace and line breaks
            Dim commandPattern As String = "\s*[\r\n]*\s*\[#[^:]+:\s*@@[^@]+@@\s*(?:§§[^§]*§§)?\s*#\]\s*[\r\n]*\s*"
            Dim regex As New Regex(commandPattern)
            output = regex.Replace(input, "")

            ' Collapse multiple consecutive line breaks into a single line break
            Dim whitespacePattern As String = "[\r\n]{3,}"
            Dim collapseRegex As New Regex(whitespacePattern)
            output = collapseRegex.Replace(output, Environment.NewLine)

        Catch ex As System.Exception
            MsgBox("Error in RemoveCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return output
    End Function


    Private CommandsList As String = ""
    Private FailedCommandsList As New List(Of String)() ' Add this to track failed commands

    Public Sub ExecuteAnyCommands(teststring As String, OnlySelection As Boolean)

        Dim commands = ParseCommands(teststring)
        Dim topmost As Boolean = Me.TopMost

        Me.TopMost = False

        CommandsList = ""
        FailedCommandsList.Clear() ' Clear previous failed commands
        Dim LastCommandsList As String = ""

        Dim wordApp As Microsoft.Office.Interop.Word.Application
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' ============= ENSURE WE'RE IN MAIN STORY WITHOUT CHANGING SELECTION =============
        Try
            wordApp = Globals.ThisAddIn.Application

            If wordApp IsNot Nothing AndAlso wordApp.ActiveDocument IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                Dim currentDoc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument
                Dim currentSel As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                Dim currentStory As Word.WdStoryType = currentSel.StoryType

                ' Only act if we're NOT already in the main text story
                If currentStory <> Word.WdStoryType.wdMainTextStory Then
                    ' Force view back to print view to get out of special editing modes
                    wordApp.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

                    ' Move to start of main document story without selecting anything
                    Dim mainStoryRange As Word.Range = currentDoc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                    mainStoryRange.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                    mainStoryRange.Select()

                    ' Collapse to insertion point (no selection)
                    currentSel.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                End If
            End If
        Catch ex As Exception
            ' Best-effort; continue even if this fails
            Debug.WriteLine($"Warning: Could not reset to main story: {ex.Message}")
        End Try
        ' ================================================================================



        If commands.Count() > 0 Then

            Globals.ThisAddIn.Application.Activate()
            'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):")
            System.Threading.Thread.Sleep(200)

            wordApp = Globals.ThisAddIn.Application
            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = False
            End With

        End If

        For Each pc In commands
            Debug.WriteLine($"Command: '{pc.Command}' wit '{pc.Argument1}' '{pc.Argument2}'")
            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If

            Dim commandSuccess As Boolean = True ' Track success of each command
            Dim commandDescription As String = ""

            Select Case pc.Command.ToLower()
                Case "find"
                    commandDescription = $"Finding '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteFindCommand(pc.Argument1, OnlySelection)

                Case "addcomment"
                    commandDescription = $"Adding comment '{pc.Argument2}' to the text '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteAddComment(pc.Argument1, pc.Argument2, OnlySelection)

                Case "replycomment"
                    commandDescription = $"Replying to comment '{pc.Argument1}' with '{pc.Argument2}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteReplyToCommentByIdToken(pc.Argument1, pc.Argument2)

                Case "replace"
                    If String.IsNullOrEmpty(pc.Argument2) Then
                        commandDescription = $"Deleting '{pc.Argument1}'"
                    Else
                        commandDescription = $"Replacing '{pc.Argument1}' with '{pc.Argument2}'"
                    End If
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteReplaceCommand(pc.Argument1, pc.Argument2, OnlySelection, MarkerChar)

                Case "insertafter"
                    commandDescription = $"Inserting '{pc.Argument2}' after '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, False)

                Case "insertbefore"
                    commandDescription = $"Inserting '{pc.Argument2}' before '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, True)

                Case "insert"
                    commandDescription = $"Inserting '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    Debug.WriteLine("ExecuteInsert")
                    commandSuccess = ExecuteInsertCommand(pc.Argument1)

                Case Else
                    ' Unknown command or default
                    commandDescription = $"Unknown command: '{pc.Command}'"
                    commandSuccess = False
            End Select

            ' Track failed commands
            If Not commandSuccess AndAlso Not String.IsNullOrWhiteSpace(commandDescription) Then
                FailedCommandsList.Add($"Failed: {commandDescription}")
            End If

            If LastCommandsList <> CommandsList Then
                'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                System.Threading.Thread.Sleep(500)
            End If
        Next

        If commands.Count() > 0 Then

            'InfoBox.ShowInfoBox("Cleaning up ... almost done.")
            'System.Threading.Thread.Sleep(300)

            ' Remove marker
            ReplaceSpecialCharacter(OnlySelection)

            InfoBox.ShowInfoBox("")

            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = True
            End With

        End If

        ' COM-Objekt sauber freigeben
        If wordApp IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
            wordApp = Nothing
        End If

        Me.TopMost = topmost
        Me.Focus()

        ' Report failed commands to chat if any
        If FailedCommandsList.Count > 0 Then
            ReportFailedCommands()
        End If

    End Sub


    ' Returns the end position of the TOC that contains (or overlaps) the given range.
    ' 0 means: not inside/overlapping any TOC.
    Private Function TocEndIfInside(foundRange As Word.Range, doc As Word.Document) As Integer
        If foundRange Is Nothing OrElse doc Is Nothing Then Return 0
        For Each toc As Word.TableOfContents In doc.TablesOfContents
            Dim tr As Word.Range = toc.Range
            ' Treat any overlap with the TOC as "inside" for skipping purposes
            If foundRange.Start < tr.End AndAlso foundRange.End > tr.Start Then
                Return tr.End
            End If
        Next
        Return 0
    End Function

    ' Keep the old boolean helper for backwards compatibility (now uses TocEndIfInside)
    Private Function IsInsideToc(range As Word.Range, doc As Word.Document) As Boolean
        Return TocEndIfInside(range, doc) > 0
    End Function



    Private Sub ReportFailedCommands()
        If FailedCommandsList Is Nothing OrElse FailedCommandsList.Count = 0 Then Return

        Dim errorMessage As New System.Text.StringBuilder()
        errorMessage.AppendLine()
        errorMessage.AppendLine("─────────────────────────────────────")
        errorMessage.AppendLine("⚠ Some commands could not be executed:")
        errorMessage.AppendLine()

        For Each failedCmd In FailedCommandsList
            errorMessage.AppendLine($"  • {failedCmd}")
        Next

        errorMessage.AppendLine()
        errorMessage.AppendLine("─────────────────────────────────────")

        ' Add to plain text chat history
        AppendToChatHistory(errorMessage.ToString())

        ' Add to HTML chat with formatting
        Dim htmlError As String = $"<div class='msg assistant error' style='border-left: 3px solid #d93025; padding-left: 10px; margin: 10px 0; background-color: #fef1f0;'>
            <span class='who' style='color: #d93025;'>System:</span>
            <div class='content'>
                <hr style='border: none; border-top: 1px solid #d93025; margin: 8px 0;' />
                <strong>⚠ Some commands could not be executed:</strong><br/>
                <ul style='margin: 8px 0;'>"

        For Each failedCmd In FailedCommandsList
            htmlError += $"<li>{HtmlEncode(failedCmd)}</li>"
        Next

        htmlError += "</ul><hr style='border: none; border-top: 1px solid #d93025; margin: 8px 0;' /></div></div>"

        AppendHtml(htmlError)
        PersistChatHtml()

        ' Add to chat history so AI can see the failures
        _chatHistory.Add(("assistant", $"System: Some commands failed - {String.Join("; ", FailedCommandsList)}"))
    End Sub

    Private Sub ReplaceSpecialCharacter(Optional OnlySelection As Boolean = False)

        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled = doc.TrackRevisions

        Try
            doc.TrackRevisions = True
            Dim rng As Word.Range =
            If(OnlySelection AndAlso Not String.IsNullOrEmpty(doc.Application.Selection.Text),
               doc.Application.Selection.Range.Duplicate,
               doc.Content.Duplicate)

            With rng.Find
                .ClearFormatting()
                .Text = MarkerChar
                .Replacement.ClearFormatting()
                .Replacement.Text = ""
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
                Do While .Execute(Replace:=Word.WdReplace.wdReplaceOne)
                    ' keep looping until none left
                Loop
            End With
        Catch ex As Exception
            MsgBox("Error in ReplaceSpecialCharacter: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Sub


    ' Adds a threaded reply to an existing Word comment identified by a single LLM-friendly token.
    ' Token formats accepted (order of precedence):
    '  - "1234|abcdef..."              (id|hash)
    '  - "id=1234;hash=abcdef..."      (labels; separators ; , | or whitespace)
    '  - "wid:1234 ph:abcdef..."       (labels with ':' and whitespace)
    '  - "1234"                        (id only)
    '  - "abcdef..."                   (hash only)

    Private Function ExecuteReplyToCommentByIdToken(
    ByVal idToken As String,
    ByVal replyText As String
) As Boolean

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim hadSel As Boolean = False
        Dim origStart As Integer = -1
        Dim origEnd As Integer = -1

        Try
            app = Globals.ThisAddIn.Application
            If app IsNot Nothing AndAlso app.Documents IsNot Nothing AndAlso app.Documents.Count > 0 Then
                doc = app.ActiveDocument
                If app.Selection IsNot Nothing Then
                    origStart = app.Selection.Start
                    origEnd = app.Selection.End
                    hadSel = True
                End If
            End If

            ' ——— validation logic ———
            If String.IsNullOrWhiteSpace(idToken) Then
                Debug.WriteLine("Add-Reply: Missing ID token.")
                Return False
            End If
            If String.IsNullOrWhiteSpace(replyText) Then
                Debug.WriteLine("Add-Reply: Reply text is empty.")
                Return False
            End If

            Dim wordId As System.Nullable(Of Integer) = Nothing
            Dim pseudoHash As String = Nothing

            If Not TryParseCommentIdToken(idToken, wordId, pseudoHash) Then
                Debug.WriteLine("Add-Reply: Could not parse ID token (expected formats like '1234|abcdef' or 'id=1234;hash=abcdef').")
                Return False
            End If

            ' Add detailed logging to debug the issue
            Debug.WriteLine($"Add-Reply: Parsed token '{idToken}' -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")

            Dim formatted As Boolean = chkConvertMarkdown.Checked
            Dim ok As Boolean = ThisAddIn.ReplyToWordComment(wordId, pseudoHash, AN6 & ": " & replyText, formatted)

            If ok Then
                Debug.WriteLine($"Add-Reply: Successfully added reply to comment {If(wordId.HasValue, wordId.Value.ToString(), pseudoHash)}")
            Else
                Debug.WriteLine($"Add-Reply: Failed to add reply to comment {If(wordId.HasValue, wordId.Value.ToString(), pseudoHash)} (target not found).")
            End If

            Return ok

        Catch ex As Exception
            ' Log the error but don't throw - let the calling code handle it
            Debug.WriteLine($"Add-Reply Error: {ex.Message}")
            Return False
        Finally
            ' Restore focus and selection to main text story; avoid leaving caret in a comment
            Try
                If app IsNot Nothing AndAlso doc IsNot Nothing AndAlso hadSel Then
                    ' Ensure we're back in the main story before restoring selection
                    app.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                    Dim s As Integer = Math.Max(doc.Content.Start, Math.Min(origStart, doc.Content.End))
                    Dim e As Integer = Math.Max(doc.Content.Start, Math.Min(origEnd, doc.Content.End))
                    doc.Range(s, e).Select() ' use a doc Range to force wdMainTextStory
                End If
            Catch
                ' best-effort restore; ignore failures
            End Try
        End Try
    End Function



    ' Parses a combined ID token into Word comment Index (WordID) and/or PseudoHash.
    ' Returns True if at least one identifier could be extracted.
    Private Function TryParseCommentIdToken(
    ByVal raw As String,
    ByRef wordId As System.Nullable(Of Integer),
    ByRef pseudoHash As String
) As Boolean
        wordId = Nothing
        pseudoHash = Nothing
        If String.IsNullOrWhiteSpace(raw) Then Return False

        Dim s As String = raw.Trim()

        ' Log what we're trying to parse
        Debug.WriteLine($"TryParseCommentIdToken: Parsing '{s}'")

        ' 1) Fast path: split "id|hash"
        Dim pipeParts = s.Split(New Char() {"|"c}, 2, StringSplitOptions.None)
        If pipeParts.Length = 2 Then
            Dim left = pipeParts(0).Trim()
            Dim right = pipeParts(1).Trim()
            Dim idVal As Integer
            If Integer.TryParse(left, idVal) Then wordId = idVal
            If Not String.IsNullOrWhiteSpace(right) Then pseudoHash = right
            Debug.WriteLine($"TryParseCommentIdToken: Pipe format -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")
            Return (wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash))
        End If

        ' 2) Labeled forms: allow separators ; , | or whitespace; allow labels wid/id and ph/hash/pseudohash
        ' Examples: "id=1234;hash=abcdef", "wid:1234 ph:abcdef", "id:3"
        Dim idMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bwid|\bid|\bwordid)\s*[:=]\s*(?<id>-?\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If idMatch.Success Then
            Dim idVal As Integer
            If Integer.TryParse(idMatch.Groups("id").Value, idVal) Then
                wordId = idVal
                Debug.WriteLine($"TryParseCommentIdToken: Found WordId={wordId.Value} from labeled format")
            End If
        End If

        Dim hashMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bph|\bhash|\bpseudohash)\s*[:=]\s*(?<hash>[A-Za-z0-9_-]{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If hashMatch.Success Then
            pseudoHash = hashMatch.Groups("hash").Value.Trim()
            Debug.WriteLine($"TryParseCommentIdToken: Found Hash={pseudoHash} from labeled format")
        End If

        If wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash) Then
            Debug.WriteLine($"TryParseCommentIdToken: Labeled format -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")
            Return True
        End If

        ' 3) Single token fallback:
        '    - all digits => id only
        '    - otherwise   => treat as hash
        Dim onlyDigits As Boolean = s.All(Function(ch) Char.IsDigit(ch))
        If onlyDigits Then
            Dim idVal As Integer
            If Integer.TryParse(s, idVal) Then
                wordId = idVal
                Debug.WriteLine($"TryParseCommentIdToken: Plain number -> WordId={wordId.Value}")
                Return True
            End If
        Else
            ' Accept as hash if it looks non-empty
            If s.Length >= 6 Then
                pseudoHash = s
                Debug.WriteLine($"TryParseCommentIdToken: Plain text -> Hash={pseudoHash}")
                Return True
            End If
        End If

        Debug.WriteLine("TryParseCommentIdToken: Failed to parse")
        Return False
    End Function

    Private Function ExecuteAddComment(
    ByVal searchTerm As String,
    ByVal commentText As String,
    Optional ByVal onlySelection As Boolean = False
) As Boolean

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

        ' Validate inputs
        If String.IsNullOrWhiteSpace(searchTerm) Then
            Debug.WriteLine("AddComments: Search term is empty.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(commentText) Then
            Debug.WriteLine("AddComments: Comment text is empty.")
            Return False
        End If

        ' Get Word application and active document
        Try
            Try
                app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch
                app = Globals.ThisAddIn.Application
            End Try
        Catch ex As System.Exception
            Debug.WriteLine("AddComments: Unable to access Word Application instance.")
            Return False
        End Try

        Try
            doc = app.ActiveDocument
        Catch
            Debug.WriteLine("AddComments: No active document found.")
            Return False
        End Try
        If doc Is Nothing Then
            Debug.WriteLine("AddComments: No active document found.")
            Return False
        End If

        Dim sel As Microsoft.Office.Interop.Word.Selection = doc.Application.Selection
        Dim originalSelStart As Integer = sel.Start
        Dim originalSelEnd As Integer = sel.End

        ' Determine working range
        Dim workRange As Microsoft.Office.Interop.Word.Range
        If onlySelection AndAlso sel IsNot Nothing AndAlso Not String.IsNullOrEmpty(sel.Text) Then
            workRange = sel.Range.Duplicate
        Else
            workRange = doc.Content.Duplicate
        End If

        ' Initialize selection to the working range bounds
        sel.SetRange(workRange.Start, workRange.End)
        Dim limitEnd As Integer = workRange.End

        Dim added As Integer = 0

        Try
            ' Iterate all matches using the robust chunk finder already available
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchTerm, sel) = True
                If sel Is Nothing Then Exit Do

                Try
                    ' Anchor the comment to the found range
                    Dim anchor As Microsoft.Office.Interop.Word.Range = sel.Range.Duplicate
                    Dim newComment As Microsoft.Office.Interop.Word.Comment = Nothing

                    ' Create empty comment, then fill body
                    newComment = doc.Comments.Add(anchor, String.Empty)

                    If chkConvertMarkdown.Checked Then
                        ThisAddIn.InsertMarkdownToComment(newComment.Range, AN6 & ": " & commentText)
                    Else
                        newComment.Range.Text = AN6 & ": " & commentText
                    End If

                    added += 1
                Catch
                    ' Ignore and continue with next occurrence
                End Try

                ' Advance selection beyond current match to continue searching
                sel.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                ' Safety: if we reached the end of our working region, stop
                If sel.Start >= limitEnd Then Exit Do

                sel.SetRange(sel.Start, limitEnd)
            Loop
        Catch ex As System.Exception
            Debug.WriteLine($"AddComments failed: {ex.Message}")
        Finally
            ' Restore original selection and ensure we're back in the main document story
            Try
                Dim s As Integer = Math.Max(doc.Content.Start, Math.Min(originalSelStart, doc.Content.End))
                Dim e As Integer = Math.Max(doc.Content.Start, Math.Min(originalSelEnd, doc.Content.End))
                doc.Range(s, e).Select()
            Catch
            End Try
        End Try
        Debug.WriteLine($"AddComments: Added {added} comments for term '{searchTerm}'.")
        Return added > 0
    End Function

    Private Function ExecuteFindCommand(searchTerm As String, Optional OnlySelection As Boolean = False) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName
        Dim selectionStart As Integer = doc.Application.Selection.Start
        Dim selectionEnd As Integer = doc.Application.Selection.End
        Dim found As Boolean = False ' Track if anything was found

        Try
            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            searchTerm = DecodeParagraphMarks(searchTerm)
            If String.IsNullOrWhiteSpace(searchTerm) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Return False ' Return false for empty search
            End If

            ' Define the starting selection based on OnlySelection
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
                End If
            Else
                doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
            End If

            Dim lastSelectionStart As Integer = -1 ' Track last selection position
            Dim stuckCounter As Integer = 0        ' Counter for repeated positions
            Dim maxStuckLimit As Integer = 2        ' Maximum allowed stuck occurrences

            ' Loop through the content to find and mark all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchTerm, doc.Application.Selection, True) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                found = True

                ' Highlight the found text
                doc.Application.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow

                ' Check if we are stuck at the same selection position
                If doc.Application.Selection.Start = lastSelectionStart Then
                    stuckCounter += 1
                    If stuckCounter >= maxStuckLimit Then
                        ' Force exit if stuck too many times
                        Exit Do
                    End If
                Else
                    stuckCounter = 0 ' Reset counter if we moved forward
                End If
                lastSelectionStart = doc.Application.Selection.Start ' Update tracking

                ' Collapse the selection to the end of the current match
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' Check if the selection is inside a table and at the end of a cell
                If doc.Application.Selection.Range.Tables.Count > 0 Then
                    Try
                        Dim currentCell As Word.Cell = doc.Application.Selection.Cells(1) ' Get current cell

                        ' Ensure that we are at the end of the current cell
                        If doc.Application.Selection.End >= currentCell.Range.End - 1 Then
                            ' Move to the next cell or out of the table
                            doc.Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCell, Count:=1, Extend:=Word.WdMovementType.wdMove)
                        End If

                    Catch ex As System.Exception
                        ' If an error occurs, it means the selection is not inside a valid cell - ignore and continue
                    End Try
                End If

                ' Ensure we don't get stuck inside an empty cell
                If doc.Application.Selection.Range.Text = vbCr Or doc.Application.Selection.Range.Text = "" Then
                    doc.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                End If

                ' Check if the collapsed selection has reached the end of the document or the selection
                If OnlySelection Then
                    If doc.Application.Selection.Start >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, selectionEnd)
                Else
                    If doc.Application.Selection.Start >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, doc.Content.End)
                End If
            Loop

            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

            Return found ' Return success status

        Catch ex As System.Exception
            MsgBox("Error in ExecuteFindCommand: " & ex.Message)
            Return False ' Return false on error

        Finally
            ' Restore original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor

            ' Restore original selection
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()
        End Try
    End Function


    Private Function ExecuteReplaceCommand(oldText As String, newText As String, OnlySelection As Boolean, Marker As String) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        ' Ensure markup is visible while we operate; restore at the end.
        Dim view = doc.Application.ActiveWindow.View
        Dim originalRevisionsView As Word.WdRevisionsView = view.RevisionsView
        Dim originalShowRevisions As Boolean = view.ShowRevisionsAndComments

        Try
            oldText = DecodeParagraphMarks(oldText)
            newText = DecodeParagraphMarks(newText)

            ' Normalize inputs
            oldText = If(oldText, String.Empty)
            newText = If(newText, String.Empty)

            If String.IsNullOrWhiteSpace(oldText) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            ' Show "Final Showing Markup" while replacing
            view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
            view.ShowRevisionsAndComments = True

            Dim workRange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workRange = doc.Content.Duplicate
                Else
                    workRange = doc.Application.Selection.Range.Duplicate
                End If
            Else
                workRange = doc.Content.Duplicate
            End If

            Debug.WriteLine($"Replacing '{oldText}' with '{newText}'")

            Dim newTextWithMarker As String
            If newText.Length > 2 Then
                newTextWithMarker =
                    newText.Substring(0, newText.Length - 2) &
                    Marker &
                    newText.Substring(newText.Length - 2)
            Else
                newTextWithMarker = newText
            End If

            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End
            doc.Application.Selection.SetRange(workRange.Start, workRange.End)

            Dim found As Boolean = False

            ' Loop through the content to find and replace all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, doc.Application.Selection, True) = True
                If doc.Application.Selection Is Nothing Then Exit Do
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit Do

                found = True

                Dim currentEnd As Integer

                ' Remember where the match starts; Selection will move after assignment
                Dim matchStart As Integer = doc.Application.Selection.Start

                    ' Replace the text
                    doc.Application.Selection.Text = newTextWithMarker

                    ' Select exactly the inserted span for conversion, if any
                    If chkConvertMarkdown.Checked AndAlso newTextWithMarker.Length > 0 Then
                        Try
                            Dim insStart As Integer = matchStart
                            Dim insEnd As Integer = matchStart + Len(newTextWithMarker)
                            doc.Range(insStart, insEnd).Select()
                            Globals.ThisAddIn.ConvertMarkdownToWord()
                        Catch
                            ' Best effort
                        End Try
                    End If

                    ' Maintain selection bounds when constrained
                    If OnlySelection Then
                        selectionEnd = selectionEnd - Len(oldText) + Len(newTextWithMarker)
                    End If

                    ' Continue after the inserted text
                    currentEnd = matchStart + Len(newTextWithMarker)


                ' Continue searching within the allowed range
                If OnlySelection Then
                    If currentEnd >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(currentEnd, selectionEnd)
                Else
                    If currentEnd >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(currentEnd, doc.Content.End)
                End If
            Loop

            If Not found Then
                CommandsList = $"Note: The search term was not found (Chunk Search)." & Environment.NewLine & CommandsList
            End If

            ' Restore original selection
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()

            Return found

        Catch ex As System.Exception
#If DEBUG Then
            Debug.WriteLine("Error: " & ex.Message)
            Debug.WriteLine("Stacktrace: " & ex.StackTrace)
            System.Diagnostics.Debugger.Break()
#End If
            MsgBox("Error in ExecuteReplaceCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False

        Finally
            ' Restore original settings
            view.RevisionsView = originalRevisionsView
            view.ShowRevisionsAndComments = originalShowRevisions

            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Function



    Private Function ExecuteInsertBeforeAfterCommand(searchText As String, newText As String, Optional OnlySelection As Boolean = False, Optional InsertBefore As Boolean = False) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Save the current state of Track Changes and Author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            searchText = DecodeParagraphMarks(searchText)
            newText = DecodeParagraphMarks(newText)
            If String.IsNullOrWhiteSpace(searchText) Then
                CommandsList = $"Note: Empty insertion anchor (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            doc.Application.Activate()
            doc.Activate()

            ' Enable Track Changes
            doc.TrackRevisions = True

            ' Determine the range for the search
            Dim workrange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workrange = doc.Content
                Else
                    workrange = doc.Application.Selection.Range
                End If
            Else
                workrange = doc.Content
            End If

            Dim found As Boolean = False
            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End

            doc.Application.Selection.SetRange(workrange.Start, workrange.End)

            ' Add safety measures
            Dim maxIterations As Integer = 1000
            Dim iterationCount As Integer = 0
            Dim lastProcessedPosition As Integer = -1

            ' Loop through the content to find and insert at all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchText, doc.Application.Selection, True) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                ' Safety check for infinite loop
                iterationCount += 1
                If iterationCount > maxIterations Then
                    Debug.WriteLine($"ExecuteInsertBeforeAfterCommand: Max iterations ({maxIterations}) reached, breaking loop")
                    Exit Do
                End If

                ' Check if we're processing the same position again
                If doc.Application.Selection.Start = lastProcessedPosition Then
                    Debug.WriteLine("ExecuteInsertBeforeAfterCommand: Stuck at same position, advancing")
                    doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    doc.Application.Selection.Move(Word.WdUnits.wdCharacter, 1)
                    Continue Do
                End If
                lastProcessedPosition = doc.Application.Selection.Start

                ' Store the found range
                Dim foundRange As Word.Range = doc.Application.Selection.Range.Duplicate

                ' Inside ExecuteInsertBeforeAfterCommand, replace the TOC check block with this:
                Dim tocEnd As Integer = TocEndIfInside(foundRange, doc)
                If tocEnd > 0 Then
                    Debug.WriteLine("ExecuteInsertBeforeAfterCommand: Match in TOC -> skipping to end of TOC")
                    Dim searchLimit As Integer = If(OnlySelection, selectionEnd, doc.Content.End)
                    Dim continuePos As Integer = Math.Min(tocEnd, searchLimit)

                    If continuePos >= searchLimit Then
                        Exit Do
                    Else
                        doc.Application.Selection.SetRange(continuePos, searchLimit)
                        Continue Do
                    End If
                End If

                found = True

                Dim foundStart As Integer = foundRange.Start
                Dim foundEnd As Integer = foundRange.End
                Dim insertPosition As Integer

                If InsertBefore Then
                    insertPosition = foundStart
                Else
                    insertPosition = foundEnd
                End If

                ' Create a new range at the exact insertion point
                Dim insertRange As Word.Range = doc.Range(insertPosition, insertPosition)

                ' Insert the new text at the precise position
                insertRange.Text = newText

                ' Apply markdown conversion if needed
                If chkConvertMarkdown.Checked AndAlso newText.Length > 0 Then
                    Try
                        doc.Range(insertPosition, insertPosition + Len(newText)).Select()
                        Globals.ThisAddIn.ConvertMarkdownToWord()
                    Catch
                        ' Best effort - continue if conversion fails
                    End Try
                End If

                ' Calculate where to continue searching from
                Dim continuePosition As Integer
                If InsertBefore Then
                    ' Move past both inserted text and original found text
                    continuePosition = insertPosition + Len(newText) + (foundEnd - foundStart)
                Else
                    ' Move past the insertion point plus the new text
                    continuePosition = insertPosition + Len(newText)
                End If

                ' Make sure we're moving forward
                If continuePosition <= lastProcessedPosition Then
                    continuePosition = lastProcessedPosition + 1
                End If

                ' Adjust the working end if text was inserted
                If OnlySelection Then
                    selectionEnd = selectionEnd + Len(newText)
                End If

                ' Check if we've reached the end of the search range
                If OnlySelection Then
                    If continuePosition >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(continuePosition, selectionEnd)
                Else
                    If continuePosition >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(continuePosition, doc.Content.End)
                End If
            Loop

            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

            ' Restore original selection
            Try
                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()
            Catch
                ' If restoration fails, just collapse to start
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            End Try

            Return found

        Catch ex As System.Exception

#If DEBUG Then
        Debug.WriteLine("Error: " & ex.Message)
        Debug.WriteLine("Stacktrace: " & ex.StackTrace)
        System.Diagnostics.Debugger.Break()
#End If

            MsgBox("Error in ExecuteInsertBeforeAfterCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False

        Finally
            ' Restore the original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Function

    Private Function oldExecuteInsertBeforeAfterCommand(searchText As String, newText As String, Optional OnlySelection As Boolean = False, Optional InsertBefore As Boolean = False) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Save the current state of Track Changes and Author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            searchText = DecodeParagraphMarks(searchText)
            newText = DecodeParagraphMarks(newText)
            If String.IsNullOrWhiteSpace(searchText) Then
                CommandsList = $"Note: Empty insertion anchor (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            doc.Application.Activate()
            doc.Activate()

            ' Enable Track Changes
            doc.TrackRevisions = True

            ' Determine the range for the search
            Dim workrange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workrange = doc.Content
                Else
                    workrange = doc.Application.Selection.Range
                End If
            Else
                workrange = doc.Content
            End If

            Dim found As Boolean = False
            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End

            doc.Application.Selection.SetRange(workrange.Start, workrange.End)

            ' Add a safety counter to prevent infinite loops
            Dim maxIterations As Integer = 1000
            Dim iterationCount As Integer = 0

            ' Loop through the content to find and insert at all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchText, doc.Application.Selection, True) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                ' Safety check for infinite loop
                iterationCount += 1
                If iterationCount > maxIterations Then
                    Debug.WriteLine($"ExecuteInsertBeforeAfterCommand: Max iterations ({maxIterations}) reached, breaking loop")
                    Exit Do
                End If

                found = True

                ' Store the found range boundaries immediately after finding
                Dim foundRange As Word.Range = doc.Application.Selection.Range.Duplicate
                Dim foundOriginalEnd As Integer = foundRange.End  ' Store the original end position
                Dim insertPosition As Integer

                If InsertBefore Then
                    ' For insert before, use the start of the found range
                    insertPosition = foundRange.Start
                Else
                    ' For insert after, use the end of the found range
                    insertPosition = foundRange.End
                End If

                ' Create a new range at the exact insertion point
                Dim insertRange As Word.Range = doc.Range(insertPosition, insertPosition)

                ' Insert the new text at the precise position
                insertRange.Text = newText

                ' Apply markdown conversion if needed
                If chkConvertMarkdown.Checked Then
                    ' Select the newly inserted text for conversion
                    doc.Range(insertPosition, insertPosition + Len(newText)).Select()
                    Globals.ThisAddIn.ConvertMarkdownToWord()
                End If

                ' Calculate the position to continue searching from
                ' We need to move past both the found text and (if inserting before) the inserted text
                Dim continuePosition As Integer
                If InsertBefore Then
                    ' Skip past the inserted text AND the original found text
                    continuePosition = insertPosition + Len(newText) + (foundOriginalEnd - insertPosition)
                Else
                    ' We're already past the found text, just account for the inserted text
                    continuePosition = insertPosition + Len(newText)
                End If

                ' Adjust the working end if text was inserted
                selectionEnd = selectionEnd + Len(newText)

                ' Check if we've reached the end of the search range
                If OnlySelection Then
                    If continuePosition >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(continuePosition, selectionEnd)
                Else
                    If continuePosition >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(continuePosition, doc.Content.End)
                End If
            Loop

            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

            ' Restore original selection
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()

            Return found

        Catch ex As System.Exception

#If DEBUG Then
        Debug.WriteLine("Error: " & ex.Message)
        Debug.WriteLine("Stacktrace: " & ex.StackTrace)
        System.Diagnostics.Debugger.Break()
#End If

            MsgBox("Error in ExecuteInsertBeforeAfterCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False

        Finally
            ' Restore the original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Function



    Private Function ExecuteInsertCommand(newText As String) As Boolean
        Dim doc = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled = doc.TrackRevisions
        Try
            newText = DecodeParagraphMarks(newText)
            ' Ensure single paragraph delimiter style (Word uses Chr(13))
            newText = newText.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)
            doc.TrackRevisions = True
            Dim selection = doc.Application.Selection
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            selection.Text = newText
            If chkConvertMarkdown.Checked Then Globals.ThisAddIn.ConvertMarkdownToWord()
            Return True ' Success
        Catch ex As Exception
            MsgBox("Error in ExecuteInsertCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False ' Failed
        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Function


End Class

Partial Public Class frmAIChat

    ' Add a field
    Private _docClickHooked As Boolean = False

    ' Wire the document-level click handler (call this when the document is ready)
    Private Sub WireDocumentClick()
        If wbChat Is Nothing OrElse wbChat.Document Is Nothing Then Return
        Try
            ' Remove before add to avoid duplicates on re-init
            RemoveHandler wbChat.Document.Click, AddressOf Doc_Click
        Catch
        End Try
        AddHandler wbChat.Document.Click, AddressOf Doc_Click
        _docClickHooked = True
    End Sub

    ' Global click handler for the HTML document; finds nearest <a> and opens externally
    Private Sub Doc_Click(sender As Object, e As HtmlElementEventArgs)
        Try
            Dim el As HtmlElement = wbChat.Document.ActiveElement
            ' Walk up to the nearest anchor
            While el IsNot Nothing AndAlso Not String.Equals(el.TagName, "A", StringComparison.OrdinalIgnoreCase)
                el = el.Parent
            End While

            If el Is Nothing Then Return

            Dim href As String = el.GetAttribute("href")
            If String.IsNullOrWhiteSpace(href) Then Return

            ' Only handle external links
            Dim lower = href.Trim().ToLowerInvariant()
            If lower.StartsWith("http://") OrElse lower.StartsWith("https://") OrElse lower.StartsWith("mailto:") Then
                Process.Start(New ProcessStartInfo(href) With {.UseShellExecute = True})
                ' Prevent the WebBrowser from navigating internally
                If e IsNot Nothing Then
                    e.ReturnValue = False
                    e.BubbleEvent = False
                End If
            End If
        Catch
            ' ignore
        End Try
    End Sub


    ' HTML renderer for the chat history: an overlay WebBrowser on top of the hidden txtChatHistory.
    Private ReadOnly wbChat As New WebBrowser() With {
    .Dock = DockStyle.Fill,
    .AllowWebBrowserDrop = False,
    .IsWebBrowserContextMenuEnabled = True,
    .WebBrowserShortcutsEnabled = True,
    .ScriptErrorsSuppressed = True
}

    ' Queue + readiness flag so we can append even if the WebBrowser is not yet ready.
    Private _htmlReady As Boolean = False
    Private ReadOnly _htmlQueue As New List(Of String)()

    ' Extended Markdown pipeline for chat (advanced features + emoji + soft line breaks).
    Private ReadOnly _mdPipeline As MarkdownPipeline =
        New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseEmojiAndSmiley().
            UseSoftlineBreakAsHardlineBreak().
            Build()

    Private _lastThinkingId As String = Nothing

    ' Bridge to open links in the default browser from inside the WebBrowser control
    <System.Runtime.InteropServices.ComVisible(True)>
    Public Class BrowserBridge
        Public Sub OpenLink(url As String)
            Try
                If String.IsNullOrEmpty(url) Then Return
                Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
            Catch
                ' ignore
            End Try
        End Sub
    End Class

    ' Persist the inner HTML of the chat container (#chat) into My.Settings.
    Private Sub PersistChatHtml()
        Try
            If wbChat Is Nothing OrElse wbChat.Document Is Nothing Then Return
            Dim chat = wbChat.Document.GetElementById("chat")
            If chat Is Nothing Then Return
            My.Settings.LastChatHistoryHtml = chat.InnerHtml
            My.Settings.Save()
        Catch
            ' best-effort
        End Try
    End Sub

    ' Call from constructor, right after placing txtChatHistory in the TableLayoutPanel.
    Public Sub InitChatHtmlUI(host As TableLayoutPanel)
        If host Is Nothing Then Return

        txtChatHistory.Visible = False
        host.Controls.Add(wbChat, 0, 1)
        wbChat.BringToFront()

        wbChat.ObjectForScripting = New BrowserBridge()

        AddHandler wbChat.DocumentCompleted, AddressOf WbChat_DocumentCompleted
        AddHandler wbChat.Navigating, AddressOf WbChat_Navigating
        AddHandler wbChat.NewWindow, AddressOf WbChat_NewWindow
    End Sub

    Private Sub WbChat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            If e.Url IsNot Nothing Then
                Dim scheme = e.Url.Scheme.ToLowerInvariant()
                If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                    e.Cancel = True
                    Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
                End If
            End If
        Catch
            ' ignore
        End Try
    End Sub

    Private Sub WbChat_NewWindow(sender As Object, e As CancelEventArgs)
        e.Cancel = True
        Try
            Dim doc = wbChat.Document
            If doc IsNot Nothing AndAlso doc.ActiveElement IsNot Nothing Then
                Dim href = doc.ActiveElement.GetAttribute("href")
                If Not String.IsNullOrWhiteSpace(href) Then
                    Process.Start(New ProcessStartInfo(href) With {.UseShellExecute = True})
                End If
            End If
        Catch
            ' ignore
        End Try
    End Sub

    ' Call once in Load after controls are set up.
    Public Sub InitializeChatHtml()
        Dim baseSize As Single = If(Me IsNot Nothing AndAlso Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
        Dim fontPt As Single = System.Math.Max(baseSize + 1.0F, 10.0F)

        Dim css As String =
$"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
#chat{{padding:6px 8px;}}
.msg{{margin:6px 0;word-wrap:break-word;}}
.msg .who{{font-weight:600;margin-right:4px;}}
.msg.user .who{{color:#333;}}
.msg.assistant .who{{color:#003366;}}
.msg.thinking .content{{opacity:.75;font-style:italic;}}
/* No top gap when content is block-rendered */
.msg .content > *:first-child{{margin-top:0;}}
a{{color:#0068c9;text-decoration:underline;cursor:pointer;}}
a:visited{{color:#5a3694;}}
ul,ol{{margin:6px 0 6px 22px;}}
pre,code,kbd,samp{{font-family:Consolas,'Courier New',monospace;}}
pre{{white-space:pre-wrap;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:4px;padding:6px;}}
blockquote{{border-left:4px solid #e1e4e8;margin:6px 0;padding:6px 10px;background:#fafbfc;color:#333;}}
table{{border-collapse:collapse;margin:6px 0;}}
td,th{{border:1px solid #ddd;padding:4px 6px;}}"

        Dim html As String =
$"<!DOCTYPE html>
<html>
<head>
<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
<meta charset=""utf-8"">
<style>{css}</style>
<script type=""text/javascript"">
function wireLinks(root) {{
  var links = root.getElementsByTagName('a');
  for (var i = 0; i < links.length; i++) {{
    (function(a) {{
      a.setAttribute('target', '_self');    // avoid NewWindow for old IE
      a.setAttribute('rel', 'noopener');
      a.onclick = function() {{
        try {{ if (window.external && window.external.OpenLink) window.external.OpenLink(a.href); }} catch (e) {{}}
        if (window.event) window.event.returnValue = false; // IE8-
        return false;
      }};
    }})(links[i]);
  }}
}}
function appendMessage(html) {{
  var c = document.getElementById('chat');
  if (!c) return;
  var temp = document.createElement('div');
  temp.innerHTML = html;
  wireLinks(temp);
  while (temp.firstChild) {{
    c.appendChild(temp.firstChild);
  }}
  window.scrollTo(0, document.body.scrollHeight);
}}
function removeById(id) {{
  var el = document.getElementById(id);
  if (!el || !el.parentNode) return;
  el.parentNode.removeChild(el);
}}
</script>
</head>
<body>
  <div id=""chat""></div>
</body>
</html>"
        _htmlReady = False
        wbChat.DocumentText = html
    End Sub

    ' Clear the HTML chat entirely
    Public Sub ClearChatHtml()
        _htmlQueue.Clear()
        _htmlReady = False
        InitializeChatHtml()
    End Sub

    ' Safe HTML encode for plain text parts.
    Private Shared Function HtmlEncode(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("&", "&amp;").
                 Replace("<", "&lt;").
                 Replace(">", "&gt;").
                 Replace("""", "&quot;")
    End Function

    Private Shared Function InstrumentLinks(html As String) As String
        If String.IsNullOrEmpty(html) Then Return html
        Try
            Return System.Text.RegularExpressions.Regex.Replace(
                html,
                "(?is)<a\s+([^>]*?)\bhref\s*=\s*(?:'([^']*)'|""([^""]*)""|([^\s>]+))([^>]*)>",
                Function(m As System.Text.RegularExpressions.Match)
                    Dim pre = m.Groups(1).Value
                    Dim href = If(m.Groups(2).Success, m.Groups(2).Value, If(m.Groups(3).Success, m.Groups(3).Value, m.Groups(4).Value))
                    Dim post = m.Groups(5).Value
                    If String.IsNullOrWhiteSpace(href) Then Return m.Value
                    ' Already wired?
                    If m.Value.IndexOf("OpenLink", StringComparison.OrdinalIgnoreCase) >= 0 Then Return m.Value
                    Dim safeHref = href.Replace("""", "&quot;")
                    Dim onclickAttr = " onclick=""try{if(window.external&&window.external.OpenLink)window.external.OpenLink(this.href);}catch(e){};return false;"""
                    ' Force target=_self to avoid popup in old IE
                    Dim targetAttr = If(m.Value.IndexOf("target=", StringComparison.OrdinalIgnoreCase) >= 0, "", " target=""_self""")
                    Return $"<a {pre} href=""{safeHref}""{targetAttr}{onclickAttr}{post}>"
                End Function)
        Catch
            Return html
        End Try
    End Function


    ' Append a restored transcript as HTML:
    ' - "You:" messages are plain, HTML-encoded.
    ' - Assistant messages (AN5) are rendered from Markdown with commands removed.
    Public Sub AppendTranscriptToHtml(transcript As String)
        If String.IsNullOrEmpty(transcript) Then Return

        Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New String() {vbLf}, StringSplitOptions.None)
        Dim currentRole As String = Nothing
        Dim content As New System.Text.StringBuilder()

        Dim SubFlush As System.Action =
            Sub()
                If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                    content.Clear() : currentRole = Nothing : Return
                End If
                Dim htmlFrag As String
                If currentRole = "user" Then
                    Dim encoded = HtmlEncode(content.ToString()).Replace(vbLf, "<br>")
                    htmlFrag = $"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>"
                Else
                    ' Assistant: convert markdown and inline single <p> when possible.
                    Dim md = RemoveCommands(content.ToString())
                    Dim body = Markdown.ToHtml(md, _mdPipeline)
                    body = InstrumentLinks(body)
                    Dim t = If(body, "").Trim()
                    Dim isSingleParagraph As Boolean =
                        System.Text.RegularExpressions.Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                        Not System.Text.RegularExpressions.Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)
                    If isSingleParagraph Then
                        Dim inlineHtml As String = System.Text.RegularExpressions.Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
                        htmlFrag = $"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>{inlineHtml}</span></div>"
                    Else
                        htmlFrag = $"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><div class='content'>{body}</div></div>"
                    End If
                End If
                AppendHtml(htmlFrag)
                content.Clear()
                currentRole = Nothing
            End Sub

        For Each ln In lines
            If ln.StartsWith("You:", StringComparison.OrdinalIgnoreCase) Then
                SubFlush()
                currentRole = "user"
                content.Append(ln.Substring(4).TrimStart())
            ElseIf ln.StartsWith(AN5 & ":", StringComparison.OrdinalIgnoreCase) Then
                SubFlush()
                currentRole = "assistant"
                content.Append(ln.Substring((AN5 & ":").Length).TrimStart())
            Else
                If content.Length > 0 Then content.AppendLine()
                content.Append(ln)
            End If
        Next
        SubFlush()
        PersistChatHtml()
    End Sub

    ' Append a user message as HTML-encoded text (no Markdown for user input).
    Public Sub AppendUserHtml(text As String)
        Dim encoded = HtmlEncode(text).
                      Replace(vbCrLf, "<br>").
                      Replace(vbLf, "<br>").
                      Replace(vbCr, "<br>")
        AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
        PersistChatHtml()
    End Sub

    ' Show "Thinking..." placeholder and remember its DOM id.
    Public Sub ShowAssistantThinking()
        _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
        AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>Thinking...</span></div>")
    End Sub

    ' Remove the last "Thinking..." placeholder if present.
    Public Sub RemoveAssistantThinking()
        If String.IsNullOrEmpty(_lastThinkingId) Then Return
        Try
            If wbChat.Document IsNot Nothing Then
                wbChat.Document.InvokeScript("removeById", New Object() {_lastThinkingId})
            End If
        Catch
            ' Best effort; ignore.
        Finally
            _lastThinkingId = Nothing
        End Try
    End Sub

    ' Append an assistant message by converting Markdown -> HTML using Markdig.
    Public Sub AppendAssistantMarkdown(md As String)
        If md Is Nothing Then md = ""
        Dim body As String = Markdown.ToHtml(md, _mdPipeline)
        body = InstrumentLinks(body)
        Dim t As String = If(body, "").Trim()

        Dim isSingleParagraph As Boolean =
            System.Text.RegularExpressions.Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
            Not System.Text.RegularExpressions.Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

        If isSingleParagraph Then
            Dim inlineHtml As String = System.Text.RegularExpressions.Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
            AppendHtml($"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>{inlineHtml}</span></div>")
        Else
            AppendHtml($"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><div class='content'>{body}</div></div>")
        End If

        PersistChatHtml()
    End Sub

    Private Sub AppendHtml(fragment As String)
        If String.IsNullOrEmpty(fragment) Then Return

        ' If the WebBrowser isn't ready, buffer messages.
        If Not _htmlReady OrElse wbChat.Document Is Nothing Then
            _htmlQueue.Add(fragment)
            Return
        End If

        Try
            wbChat.Document.InvokeScript("appendMessage", New Object() {fragment})
        Catch
            ' If we hit a timing edge, queue and wait for next ready cycle.
            _htmlQueue.Add(fragment)
        End Try
    End Sub

    ' When the HTML document is ready, flush any queued fragments.
    Private Sub WbChat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        _htmlReady = True

        WireDocumentClick()

        If _htmlQueue.Count > 0 Then
            Try
                For Each frag In _htmlQueue
                    wbChat.Document.InvokeScript("appendMessage", New Object() {frag})
                Next
            Catch
            Finally
                _htmlQueue.Clear()
            End Try
        End If
    End Sub

End Class