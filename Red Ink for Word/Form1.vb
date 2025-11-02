' Red Ink for Word -- Chatbot Form Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 29.10.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.


Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports NAudio
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Class frmAIChat

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    Const AN As String = "Red Ink"
    Const AN5 As String = "Inky"   ' for Chatbox

    Const MarkerChar As String = ChrW(&HE000)

    Private PreceedingNewline As String = ""
    Private OldChat As String = ""
    Private UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
    Private SystemPrompt As String = ""

    Private WithEvents btnCopy As New Button() With {.Text = "Copy All", .AutoSize = True}
    Private WithEvents btnCopyLastAnswer As New Button() With {.Text = "Copy Last Answer", .AutoSize = True}
    Private WithEvents btnClear As New Button() With {.Text = "Clear", .AutoSize = True}
    Private WithEvents btnExit As New Button() With {.Text = "Quit", .AutoSize = True}
    Private WithEvents btnSend As New Button() With {.Text = "Send", .AutoSize = True}
    Private WithEvents btnSwitchModel As New Button() With {.Text = "Switch Model", .AutoSize = True}
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {.Text = "Include document", .AutoSize = True, .Checked = My.Settings.IncludeDocument}
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {.Text = "Include selection", .AutoSize = True, .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)}
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {.Text = "Grant write access", .AutoSize = True, .Checked = My.Settings.DoCommands}
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {.Text = "Not always on top", .AutoSize = True, .Checked = My.Settings.NotAlwaysOnTop}
    Private WithEvents chkConvertMarkdown As New System.Windows.Forms.CheckBox() With {.Text = "Do format", .AutoSize = True, .Checked = My.Settings.ConvertMarkdownInChat}


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

        ' Set the form's title and custom icon
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable ' Ensure border supports icons
        Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
        Me.TopMost = True ' Always on top

        ' Set the initial and minimum size of the form
        Me.MinimumSize = New Size(830, 521)

        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If

        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown

        ' Set up instructions label
        lblInstructions.Text = $"Enter your question and click 'Send' or Ctrl-Enter. You can allow the chatbot to perform actions on your document (search, replace, delete, insert)."
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
        If _context.INI_SecondAPI Then pnlButtons.Controls.Add(btnSwitchModel)
        pnlButtons.Controls.Add(btnExit)

        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkStayOnTop)
        pnlCheckboxes.Controls.Add(chkConvertMarkdown)


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

        If String.IsNullOrWhiteSpace(txtChatHistory.Text) Then
            Dim result = Await WelcomeMessage()
        Else
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()

        End If
        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()

    End Sub

    ' When the user clicks Send, we call the LLM with context.
    ' Then append the AI response to the conversation.

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland." & If(chkIncludeDocText.Checked, "\nYou have access to the user's document. \n", "") & If(chkIncludeselection.Checked, "\nYou have access to a selection of user's document. \n ", "") & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked), _context.SP_Add_ChatWord_Commands, "")
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

            ' Construct the full prompt
            Dim fullPrompt As New StringBuilder()

            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine("The user's document has the name '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' and has the following content: '" & docText & "'")
            End If
            If Not String.IsNullOrEmpty(selectionText) Then
                fullPrompt.AppendLine("In the user's document '" & Globals.ThisAddIn.Application.ActiveDocument.Name & "' the user has selected the following text: '" & selectionText & "'")
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

            _chatHistory.Add(("user", userPrompt.TrimEnd()))

            ' Add a placeholder for AI response while waiting
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory($"{AN5}: Thinking...")
                                End Sub)


            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, fullPrompt.ToString(), "", "", 0, _useSecondApi, True)

            aiResponse = aiResponse.TrimEnd()
            aiResponse = aiResponse.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ").Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ").Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
            aiResponse = aiResponse.Replace($"  *  ", "  " & ChrW(8226) & "  ")
            aiResponse = RemoveMarkdownFormatting(aiResponse)

            Dim CommandsString As String = ""
            If My.Settings.DoCommands And (chkIncludeselection.Checked Or chkIncludeDocText.Checked) Then
                CommandsString = aiResponse
                aiResponse = RemoveCommands(aiResponse)
                aiResponse = Regex.Replace(aiResponse, "[\r\n\s]+$", "")
            End If

            Debug.WriteLine("AI response: " & CommandsString)

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    RemoveLastLineFromChatHistory()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.TrimEnd().Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine) & Environment.NewLine)
                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        ExecuteAnyCommands(CommandsString, chkIncludeselection.Checked)
                                    End If
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse.TrimEnd()))

        Catch ex As System.Exception
            MsgBox("Error in btnSend_Click: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Private Async Function WelcomeMessage() As Task(Of String)

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland. " & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked), _context.SP_Add_ChatWord_Commands, "")
            txtUserInput.Text = ""

            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage} , such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.", "", "", 0, _useSecondApi, True)

            aiResponse = aiResponse.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf

            aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine))
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse))

            PreceedingNewline = Environment.NewLine

            Return ""

        Catch ex As System.Exception
            'MsgBox("Error in WelcomeMessage: " & ex.Message, MsgBoxStyle.Critical)
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

    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        _useSecondApi = Not _useSecondApi
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
    End Sub


    ' Clears the conversation from both the UI and saved settings.

    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.Save()
        Dim result = WelcomeMessage()
    End Sub


    ' Press Escape to close. Also button-based exit.

    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If
            My.Settings.LastChatHistory = conversation
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
        My.Settings.Save()

    End Sub


    ' Trigger the Send button on Ctrl+Enter in the user input textbox.

    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.Enter Then
            btnSend.PerformClick()
            e.Handled = True
        End If
    End Sub

    ' Reads the entire document's text.
    Private Function GetActiveDocumentText() As String
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
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
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ' Reads the current selection's text.
    Private Function GetCurrentSelectionText() As String
        Try
            ' Get the active document
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument

            ' Get the selection from the active window
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            If String.IsNullOrEmpty(sel.Text) Then
                chkIncludeselection.Checked = False
                Return ""
            Else
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
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ' Reads the entire document's text.

    Private Function OldGetActiveDocumentText() As String
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Return doc.Content.Text
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ' Reads the current selection's text.

    Private Function OldGetCurrentSelectionText() As String
        Try
            ' Get the active document
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument

            ' Get the selection from the active window
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            If String.IsNullOrEmpty(sel.Text) Then
                chkIncludeselection.Checked = False
                Return ""
            Else
                Return sel.Text
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

    Private Function ConvertMarkdownToHtml(markdown As String) As String
        Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
        Return Markdig.Markdown.ToHtml(markdown, pipeline)
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

    Public Sub ExecuteAnyCommands(teststring As String, OnlySelection As Boolean)

        Dim commands = ParseCommands(teststring)
        Dim topmost As Boolean = Me.TopMost

        Me.TopMost = False

        CommandsList = ""
        Dim LastCommandsList As String = ""

        If commands.Count() > 0 Then
            Globals.ThisAddIn.Application.Activate()
            'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):")
            System.Threading.Thread.Sleep(200)
        End If

        For Each pc In commands
            Debug.WriteLine($"Command: '{pc.Command}' wit '{pc.Argument1}' '{pc.Argument2}'")
            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If
            Select Case pc.Command.ToLower()
                Case "find"
                    CommandsList = $"Finding '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteFindCommand(pc.Argument1, OnlySelection)

                Case "addcomment"
                    CommandsList = $"Adding comment '{pc.Argument2}' to the text '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteAddComment(pc.Argument1, pc.Argument2, OnlySelection)

                Case "replycomment"
                    CommandsList = $"Replying to comment '{pc.Argument1}' to with '{pc.Argument2}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    'InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteReplyToCommentByIdToken(pc.Argument1, pc.Argument2)

                Case "replace"
                    If String.IsNullOrEmpty(pc.Argument2) Then
                        CommandsList = $"Deleting '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    Else
                        CommandsList = $"Replacing '{pc.Argument1}' with '{pc.Argument2}" & Environment.NewLine & CommandsList
                    End If
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteReplaceCommand(pc.Argument1, pc.Argument2, OnlySelection, MarkerChar)

                Case "insertafter"
                    CommandsList = $"Inserting '{pc.Argument2}' after '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, False)

                Case "insertbefore"
                    CommandsList = $"Inserting '{pc.Argument2}' before '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, True)

                Case "insert"
                    CommandsList = $"Inserting '{pc.Argument1}'" & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    Debug.WriteLine("ExecuteInsert")
                    ExecuteInsertCommand(pc.Argument1)

                Case Else
                    ' Unknown command or default
            End Select
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
        End If

        Me.TopMost = topmost
        Me.Focus()

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
    Private Sub ExecuteReplyToCommentByIdToken(
    ByVal idToken As String,
    ByVal replyText As String
)
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

            ' ——— existing logic ———
            If String.IsNullOrWhiteSpace(idToken) Then
                Debug.WriteLine("Add-Reply: Missing ID token.")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(replyText) Then
                Debug.WriteLine("Add-Reply: Reply text is empty.")
                Exit Sub
            End If

            Dim wordId As System.Nullable(Of Integer) = Nothing
            Dim pseudoHash As String = Nothing

            If Not TryParseCommentIdToken(idToken, wordId, pseudoHash) Then
                Debug.WriteLine("Add-Reply: Could not parse ID token (expected formats like '1234|abcdef' or 'id=1234;hash=abcdef').")
                Exit Sub
            End If

            Dim formatted As Boolean = chkConvertMarkdown.Checked
            Dim ok As Boolean = ThisAddIn.ReplyToWordComment(wordId, pseudoHash, replyText, formatted)
            If Not ok Then
                Debug.WriteLine("Add-Reply: Failed to add comment reply (target not found).")
            End If
            ' ——— end existing logic ———

        Finally
            ' Restore focus and selection to main text story; avoid leaving caret in a comment
            Try
                If app IsNot Nothing AndAlso doc IsNot Nothing AndAlso hadSel Then
                    app.ActiveWindow.Activate()
                    Dim s As Integer = Math.Max(doc.Content.Start, Math.Min(origStart, doc.Content.End))
                    Dim e As Integer = Math.Max(doc.Content.Start, Math.Min(origEnd, doc.Content.End))
                    doc.Range(s, e).Select() ' use a doc Range to force wdMainTextStory
                End If
            Catch
                ' best-effort restore; ignore failures
            End Try
        End Try
    End Sub



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

        ' 1) Fast path: split "id|hash"
        Dim pipeParts = s.Split(New Char() {"|"c}, 2, StringSplitOptions.None)
        If pipeParts.Length = 2 Then
            Dim left = pipeParts(0).Trim()
            Dim right = pipeParts(1).Trim()
            Dim idVal As Integer
            If Integer.TryParse(left, idVal) Then wordId = idVal
            If Not String.IsNullOrWhiteSpace(right) Then pseudoHash = right
            Return (wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash))
        End If

        ' 2) Labeled forms: allow separators ; , | or whitespace; allow labels wid/id and ph/hash/pseudohash
        ' Examples: "id=1234;hash=abcdef", "wid:1234 ph:abcdef"
        Dim idMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bwid|\bid|\bwordid)\s*[:=]\s*(?<id>-?\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If idMatch.Success Then
            Dim idVal As Integer
            If Integer.TryParse(idMatch.Groups("id").Value, idVal) Then wordId = idVal
        End If
        Dim hashMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bph|\bhash|\bpseudohash)\s*[:=]\s*(?<hash>[A-Za-z0-9_-]{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If hashMatch.Success Then
            pseudoHash = hashMatch.Groups("hash").Value.Trim()
        End If
        If wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash) Then Return True

        ' 3) Single token fallback:
        '    - all digits => id only
        '    - otherwise   => treat as hash
        Dim onlyDigits As Boolean = s.All(Function(ch) Char.IsDigit(ch))
        If onlyDigits Then
            Dim idVal As Integer
            If Integer.TryParse(s, idVal) Then wordId = idVal
            Return True
        Else
            ' Accept as hash if it looks non-empty
            If s.Length >= 6 Then
                pseudoHash = s
                Return True
            End If
        End If

        Return False
    End Function

    Private Sub ExecuteAddComment(
    ByVal searchTerm As String,
    ByVal commentText As String,
    Optional ByVal onlySelection As Boolean = False
)

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

        ' Validate inputs
        If String.IsNullOrWhiteSpace(searchTerm) Then
            Debug.WriteLine("AddComments: Search term is empty.")
            Return
        End If
        If String.IsNullOrWhiteSpace(commentText) Then
            Debug.WriteLine("AddComments: Comment text is empty.")
            Return
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
            Return
        End Try

        Try
            doc = app.ActiveDocument
        Catch
            Debug.WriteLine("AddComments: No active document found.")
            Return
        End Try
        If doc Is Nothing Then
            Debug.WriteLine("AddComments: No active document found.")
            Return
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
                        ThisAddIn.InsertMarkdownToComment(newComment.Range, commentText)
                    Else
                        newComment.Range.Text = commentText
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
        Return
    End Sub

    Private Sub ExecuteFindCommand(searchTerm As String, Optional OnlySelection As Boolean = False)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName
        Dim selectionStart As Integer = doc.Application.Selection.Start
        Dim selectionEnd As Integer = doc.Application.Selection.End

        Try
            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

            searchTerm = DecodeParagraphMarks(searchTerm)
            If String.IsNullOrWhiteSpace(searchTerm) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Exit Sub
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

            Dim found As Boolean = False

            Dim lastSelectionStart As Integer = -1 ' Track last selection position
            Dim stuckCounter As Integer = 0        ' Counter for repeated positions
            Dim maxStuckLimit As Integer = 2        ' Maximum allowed stuck occurrences

            ' Loop through the content to find and mark all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchTerm, doc.Application.Selection) = True

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

        Catch ex As System.Exception
            MsgBox("Error in ExecuteFindCommand: " & ex.Message)

        Finally
            ' Restore original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor

            ' Restore original selection
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()
        End Try
    End Sub


    Private Sub ExecuteReplaceCommand(oldText As String, newText As String, OnlySelection As Boolean, Marker As String)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try

            oldText = DecodeParagraphMarks(oldText)
            newText = DecodeParagraphMarks(newText)

            If String.IsNullOrWhiteSpace(oldText) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Exit Sub
            End If

            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True
            'doc.Application.UserName = AN

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
                newTextWithMarker = $"{newText.Substring(0, newText.Length - 2)}{Marker}{newText.Substring(newText.Length - 2)}"
            Else
                newTextWithMarker = newText
            End If

            Dim selectionStart As Integer = doc.Application.Selection.Start
                Dim selectionEnd As Integer = doc.Application.Selection.End
                doc.Application.Selection.SetRange(workRange.Start, workRange.End)
                Dim found As Boolean = False

            ' Loop through the content to find and replace all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(oldText, doc.Application.Selection) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    Exit Do
                End If

                found = True

                Dim isDeleted As Boolean = False
                For Each rev As Word.Revision In doc.Application.Selection.Range.Revisions
                    If rev.Type = Word.WdRevisionType.wdRevisionDelete Then
                        isDeleted = True
                        Exit For
                    End If
                Next

                ' Account for trackchanges being turned on, i.e. the old text remains
                Dim currentEnd As Integer = doc.Application.Selection.End
                If Not isDeleted Then
                    currentEnd = currentEnd + Len(newTextWithMarker)
                    selectionEnd = selectionEnd + Len(newTextWithMarker)
                    ' Replace the found text
                    doc.Application.Selection.Text = newTextWithMarker
                    If chkConvertMarkdown.Checked Then Globals.ThisAddIn.ConvertMarkdownToWord()
                End If

                ' Check if the collapsed selection has reached the end of the document or the selection
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

                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()


        Catch ex As System.Exception
            MsgBox("Error in ExecuteReplaceCommand: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub


    Private Sub ExecuteInsertBeforeAfterCommand(searchText As String, newText As String, Optional OnlySelection As Boolean = False, Optional InsertBefore As Boolean = False)
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' Save the current state of Track Changes and Author
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            searchText = DecodeParagraphMarks(searchText)
            newText = DecodeParagraphMarks(newText)
            If String.IsNullOrWhiteSpace(searchText) Then
                CommandsList = $"Note: Empty insertion anchor (ignored)." & Environment.NewLine & CommandsList
                Exit Sub
            End If

            doc.Application.Activate()
            doc.Activate()

            ' Enable Track Changes and set the author to 
            doc.TrackRevisions = True
            'doc.Application.UserName = AN

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

            ' Loop through the content to find and replace all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchText, doc.Application.Selection) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                found = True

                ' Account for trackchanges being turned on, i.e. the old text remains
                Dim currentEnd As Integer = doc.Application.Selection.End + Len(newText)
                selectionEnd = selectionEnd + Len(newText)

                ' Insert the found text
                If InsertBefore Then
                    doc.Application.Selection.InsertBefore(newText)
                Else
                    doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    doc.Application.Selection.Text = newText & doc.Application.Selection.Text
                End If
                If chkConvertMarkdown.Checked Then Globals.ThisAddIn.ConvertMarkdownToWord()

                ' Check if the collapsed selection has reached the end of the document or the selection
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

                doc.Application.Selection.SetRange(selectionStart, selectionEnd)
                doc.Application.Selection.Select()



                If Not found Then
                CommandsList = $"Note: The insertion point was not found." & Environment.NewLine & CommandsList
            End If

        Catch ex As System.Exception
            MsgBox("Error in ExecuteInsertBeforeAfterCommand: " & ex.Message, MsgBoxStyle.Critical)

        Finally
            ' Restore the original state of Track Changes and Author
            doc.TrackRevisions = trackChangesEnabled
            'doc.Application.UserName = originalAuthor
        End Try
    End Sub

    Private Sub ExecuteInsertCommand(newText As String)
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
        Catch ex As Exception
            MsgBox("Error in ExecuteInsertCommand: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Sub


End Class
