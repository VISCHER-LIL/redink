' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.Json.Serialization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DocumentFormat.OpenXml
Imports Google.Cloud.Speech.V1
Imports Google.Protobuf
Imports Grpc.Core
Imports Microsoft.Office.Interop.Word
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Vosk
Imports Whisper.net
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn

    Public Class TranscriptionForm

        Inherits Form

        ' --- P/Invoke for SetThreadExecutionState ---
        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
        End Function

        ' Constants for sleep prevention
        Private Const ES_CONTINUOUS As UInteger = &H80000000UI
        Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI
        Private Const ES_DISPLAY_REQUIRED As UInteger = &H2UI ' Optional: keeps the display on too

        Private _iSetTheSleepLock As Boolean = False

        Private RichTextBox1 As Forms.RichTextBox
        Private StartButton As Forms.Button
        Private StopButton As Forms.Button
        Private ClearButton As Forms.Button
        Private LoadButton As Forms.Button
        Private AudioButton As Forms.Button
        Private QuitButton As Forms.Button
        Private ProcessButton As Forms.Button
        Private cultureComboBox As Forms.ComboBox
        Private deviceComboBox As Forms.ComboBox
        Private processCombobox As Forms.ComboBox
        Private SpeakerIdent As System.Windows.Forms.CheckBox
        Private SpeakerDistance As Forms.TextBox
        Private Label1 As Label
        Private Label2 As Label
        Private StatusLabel As Label
        Private PartialTextLabel As Label
        Private ButtonPanel As Panel

        Private TranscriptPromptsTitles As New List(Of String)
        Private TranscriptPromptsLibrary As New List(Of String)

        Private recognizer As VoskRecognizer
        Private waveIn As WaveInEvent
        Private capturing As Boolean = False
        Private partialText As String = ""
        Private finalText As New StringBuilder()
        Private Const VoskTooltip = "Only for Vosk: Set similarity threshold for speaker identification (0.5-0.7 for real-time speaker tracking, 1.0-1.5 for meetings/interviews)"
        Private Const VoskToggle = "Iden"

        Private WhisperRecognizer As WhisperProcessor
        Private audioBuffer As New List(Of Single)
        Private STTCanceled As Boolean = False
        Private cts As CancellationTokenSource = New CancellationTokenSource()
        Private Const WhisperTooltip = "Only for Whisper: Select if text shall be translated to English and the threshold for detecting voice (default = 0.6, increase for noisy environments)"
        Private Const WhisperToggle = "Trans"
        Private STTModel As String = "whisper"

        Private GoogleSpeech As Boolean = False
        Private STTSecondAPI As Boolean = False
        Private IsGoogle As Boolean = False
        Private Const GoogleTooltip = "Only for Google: Set the maximum number of speakers expected for diarization (speaker tracking)"
        Private Const GoogleToggle = "Iden"
        Private googleReaderTask As System.Threading.Tasks.Task
        Private readerCts As CancellationTokenSource = New CancellationTokenSource()
        Private _stream As SpeechClient.StreamingRecognizeStream
        Private googleTranscriptStart As Integer = 0
        Private client As SpeechClient
        Private GoogleLanguageCode As String = ""
        Private audioQueue As New System.Collections.Concurrent.BlockingCollection(Of ByteString)()
        Private _googleStreamCompleted As Boolean = False
        Private Const STREAMING_LIMIT_MS As Integer = 290000  ' 4 Minuten 50 Sekunden
        Private streamingStartTime As DateTime

        Private ReadOnly ringBuffer As New Queue(Of Google.Protobuf.ByteString)()
        Private Const RING_BUFFER_SIZE As Integer = 50

        Private ReadOnly recoverySemaphore As New System.Threading.SemaphoreSlim(1, 1)
        Private writerTask As System.Threading.Tasks.Task

        ' The watchdog timer that will check for API responsiveness.
        Private _apiWatchdogTimer As System.Threading.Timer
        ' Tracks the last time we received ANY response (partial or final) from Google.
        ' We use a long representing Ticks for thread-safe updates.
        Private _lastApiResponseTicks As Long
        ' Configurable: The number of seconds of API silence before triggering a restart.
        Private Const API_RESPONSE_TIMEOUT_SECONDS As Integer = 3
        Private _lastKnownPartialResult As String = ""
        Private _justCommittedPartialText As String = ""

        ' Maps a temporary SpeakerTag (e.g., 1, 2) from the API to a consistent,
        ' human-readable label (e.g., "Speaker 1", "Speaker 2").
        Private _speakerTagToLabelMap As New Dictionary(Of Integer, String)

        ' Counter to ensure we always assign a new, unique speaker number.
        Private _nextSpeakerNumber As Integer = 1

        Private loopback As WasapiLoopbackCapture
        Private loopbackBuffer As BufferedWaveProvider
        Private loopbackCapture As WasapiLoopbackCapture
        Private loopbackRawProvider As BufferedWaveProvider
        Private loopbackResampler As MediaFoundationResampler
        Private _multiSourceSelected As Boolean = False

        Private ReadOnly Property MultiSourceEnabled As Boolean
            Get
                Return _multiSourceSelected
            End Get
        End Property

        Private sttAccessToken1 As String = String.Empty
        Private sttTokenExpiry1 As DateTime = DateTime.MinValue
        Private sttAccessToken2 As String = String.Empty
        Private sttTokenExpiry2 As DateTime = DateTime.MinValue



        ' Hilfs‐Methode: PrivateKey in 64-Zeichen-Zeilen brechen
        Public Shared Function FormatPrivateKey(rawKey As String) As String
            Dim noEscapes = rawKey.Replace("\n", "")
            Dim sb As New System.Text.StringBuilder()
            For i As Integer = 0 To noEscapes.Length - 1 Step 64
                Dim chunk = If(i + 64 <= noEscapes.Length,
                      noEscapes.Substring(i, 64),
                      noEscapes.Substring(i))
                sb.AppendLine(chunk)
            Next
            Return "-----BEGIN PRIVATE KEY-----" & vbLf &
           sb.ToString() &
           "-----END PRIVATE KEY-----" & vbLf
        End Function

        ' Neu: Holt lokal einen frischen STT-Token für die gewählte API
        Private Async Function GetFreshSTTToken(useSecond As Boolean) As System.Threading.Tasks.Task(Of String)

            Try
                Dim token As String
                Dim expiry As DateTime

                If useSecond Then
                    token = sttAccessToken2
                    expiry = sttTokenExpiry2
                Else
                    token = sttAccessToken1
                    expiry = sttTokenExpiry1
                End If

                If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                    ' Parameter je nach API auswählen
                    Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                    Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                    Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                    Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                    Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                    ' GoogleOAuthHelper konfigurieren
                    GoogleOAuthHelper.client_email = clientEmail
                    GoogleOAuthHelper.private_key = FormatPrivateKey(rawKey)
                    GoogleOAuthHelper.scopes = scopes
                    GoogleOAuthHelper.token_uri = authServer
                    GoogleOAuthHelper.token_life = life

                    ' neuen Token holen
                    Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                    Dim newExpiry As DateTime = DateTime.UtcNow.AddSeconds(life - 300)

                    If useSecond Then
                        sttAccessToken2 = newToken
                        sttTokenExpiry2 = newExpiry
                    Else
                        sttAccessToken1 = newToken
                        sttTokenExpiry1 = newExpiry
                    End If

                    token = newToken
                End If

                Return token

            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(
            $"Error fetching STT token: {ex.Message}",
            "Transcription Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        Public Sub New()
            ' Initialize UI Components
            InitializeComponents()

            Me.AutoScaleMode = AutoScaleMode.Dpi
            ''Me.AutoScaleMode = AutoScaleMode.Font

            ' Load available Vosk models
            Dim modelPath As String = Globals.ThisAddIn.INI_SpeechModelPath
            Dim modelsexist As Boolean = False

            Dim Endpoint As String = INI_Endpoint
            Dim Endpoint_2 As String = INI_Endpoint_2

            If Endpoint.Contains(GoogleIdentifier) And INI_OAuth2 Then
                STTSecondAPI = False
                IsGoogle = True
            ElseIf Endpoint_2.Contains(GoogleIdentifier) And INI_OAuth2_2 Then
                STTSecondAPI = True
                IsGoogle = True
            End If
            If IsGoogle And Not String.IsNullOrWhiteSpace(STTEndpoint) Then
                GoogleSpeech = True
                cultureComboBox.Items.Add(GoogleSTT_Desc)
                modelsexist = True
            End If

            If Directory.Exists(modelPath) Then
                For Each dir As String In Directory.GetDirectories(modelPath)
                    Dim dirName As String = System.IO.Path.GetFileName(dir)
                    If dirName.StartsWith("vosk-model") Then
                        cultureComboBox.Items.Add(dirName)
                        modelsexist = True
                    End If
                Next

                For Each file As String In Directory.GetFiles(modelPath)
                    Dim fileName As String = System.IO.Path.GetFileName(file)
                    If fileName.StartsWith("ggml") Then
                        cultureComboBox.Items.Add(fileName)
                        modelsexist = True
                    End If
                Next

            End If

            ' Pre-select the last used model if it exists in the list
            Dim lastModel As String = My.Settings.LastSpeechModel
            If Not String.IsNullOrEmpty(lastModel) AndAlso cultureComboBox.Items.Contains(lastModel) Then
                cultureComboBox.SelectedItem = lastModel
            End If

            AddHandler Me.cultureComboBox.MouseMove, AddressOf cultureComboBox_MouseMove

            LoadAudioDevices()

            AddHandler Me.deviceComboBox.MouseMove, AddressOf deviceComboBox_MouseMove

            AddHandler Me.deviceComboBox.SelectedIndexChanged, AddressOf Me.deviceComboBox_SelectedIndexChanged

            LoadAndPopulateProcessComboBox(Globals.ThisAddIn.INI_PromptLibPath_Transcript, processCombobox)

            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)

                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If

            ' Wire up event handlers
            AddHandler StartButton.Click, AddressOf StartButton_Click
            AddHandler StopButton.Click, AddressOf StopButton_Click
            AddHandler ClearButton.Click, AddressOf ClearButton_Click
            AddHandler LoadButton.Click, AddressOf LoadButton_Click
            AddHandler AudioButton.Click, AddressOf AudioButton_Click
            AddHandler QuitButton.Click, AddressOf QuitButton_Click
            AddHandler ProcessButton.Click, AddressOf ProcessButton_Click

            ' Make window resizable
            Me.MinimumSize = New System.Drawing.Size(800, 440)

            If Not modelsexist Then
                ShowCustomMessageBox($"No Vosk or Whisper models have been found at the configured path ('{modelPath}'). A model is necessary for transcribing. You can download models for free at {VoskSource} and {WhisperSource}.", $"{AN} Transcriptor")
                Me.Close()
            End If
        End Sub

        Private ToolTip As New Forms.ToolTip()

        Private Sub deviceComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
            ' runs on UI thread—safe to read SelectedItem here
            Dim s As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            _multiSourceSelected = Not String.IsNullOrEmpty(s) _
                            AndAlso s.EndsWith("(plus audio output)")
        End Sub

        Private Sub cultureComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.cultureComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.cultureComboBox, Me.cultureComboBox.Items(index).ToString())
                If Me.cultureComboBox.Items(index).startswith(GoogleSTT_Desc) Then
                    Me.SpeakerIdent.Text = GoogleToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, GoogleTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, GoogleTooltip)
                ElseIf Me.cultureComboBox.Items(index).startswith("ggml") Then
                    Me.SpeakerIdent.Text = WhisperToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, WhisperTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, WhisperTooltip)
                Else
                    Me.SpeakerIdent.Text = VoskToggle
                    ToolTip.SetToolTip(Me.SpeakerDistance, VoskTooltip)
                    ToolTip.SetToolTip(Me.SpeakerIdent, VoskTooltip)
                End If
            End If
        End Sub

        Private Sub deviceComboBox_MouseMove(sender As Object, e As MouseEventArgs)
            Dim index As Integer = Me.deviceComboBox.SelectedIndex
            If index >= 0 Then
                ToolTip.SetToolTip(Me.deviceComboBox, Me.deviceComboBox.Items(index).ToString())
            End If
        End Sub


        Public Sub ConfigureAudioOutputDevice()
            ' 1) Alle aktiven Render-Endpoints ermitteln
            Dim enumerator As New MMDeviceEnumerator()
            Dim devices As MMDeviceCollection =
        enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)

            ' 2) FriendlyNames und zugehörige IDs in parallele Arrays packen, inkl. Default als Index 0
            Dim totalCount As Integer = devices.Count + 1
            Dim deviceNames(totalCount - 1) As String
            Dim deviceIds(totalCount - 1) As String

            ' 2a) Default Audio Output Device (wie von WasapiLoopbackCapture)
            deviceNames(0) = "Default Audio Output Device"
            deviceIds(0) = String.Empty

            ' 2b) Alle anderen Geräte ab Index 1
            For i As Integer = 0 To devices.Count - 1
                deviceNames(i + 1) = devices(i).FriendlyName
                deviceIds(i + 1) = devices(i).ID
            Next

            ' 3) Aktuell in den Settings gespeichertes Device ermitteln (leere ID → Default)
            Dim currentDeviceId As String = My.Settings.AudioOutputDevice
            Dim currentDeviceName As String = String.Empty
            Dim idxSaved As Integer = Array.IndexOf(deviceIds, currentDeviceId)
            If idxSaved >= 0 Then
                currentDeviceName = deviceNames(idxSaved)
            End If

            ' 4) Prompt für den Auswahl-Dialog zusammenbauen
            Dim prompt As String = "Choose the audio output device for capturing"
            If Not String.IsNullOrEmpty(currentDeviceName) Then
                prompt &= $" (currently: {currentDeviceName})"
            End If
            prompt &= ":"

            ' 5) Auswahl-Dialog anzeigen
            Dim selection As String = ShowSelectionForm(
        prompt,
        $"{AN} Transcriptor",
        deviceNames)

            ' 6) Wenn Auswahl gültig, Index ermitteln und Settings setzen/clearen
            If Not String.IsNullOrEmpty(selection) AndAlso selection <> "esc" Then
                Dim chosenIndex As Integer = Array.IndexOf(deviceNames, selection)
                If chosenIndex >= 0 Then
                    If chosenIndex = 0 Then
                        ' Default gewählt → Setting leeren
                        My.Settings.AudioOutputDevice = String.Empty
                    Else
                        ' Konkrete Device-ID speichern
                        My.Settings.AudioOutputDevice = deviceIds(chosenIndex)
                    End If

                    Try
                        My.Settings.Save()
                    Catch ex As System.Exception
                        ' Volle Referenz auf Exception
                        ShowCustomMessageBox($"Error saving audio output device setting: {ex.Message}")
                    End Try
                End If
            End If
        End Sub

        Private Sub InitializeComponents()
            ' --- DPI‐aware form setup ---
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.AutoScaleMode = AutoScaleMode.Font
            Me.Text = $"{AN} Transcriptor (editable text, audio will not be stored)"
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' --- Create controls ---

            ' Transcript area
            Me.RichTextBox1 = New RichTextBox() With {
        .Font = New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point),
        .Multiline = True,
        .ScrollBars = RichTextBoxScrollBars.Vertical,
        .Dock = DockStyle.Fill
    }

            ' Selector labels
            Me.Label1 = New Label() With {.Text = "Model:", .AutoSize = True}
            Me.Label2 = New Label() With {.Text = "Source:", .AutoSize = True}

            ' Model / source dropdowns (start 50px wider)
            Me.cultureComboBox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 250
    }
            Me.deviceComboBox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 450
    }

            ' Speaker toggle + threshold
            Me.SpeakerIdent = New System.Windows.Forms.CheckBox() With {.Text = VoskToggle, .AutoSize = True}
            Me.SpeakerDistance = New System.Windows.Forms.TextBox() With {
        .Text = If(My.Settings.LastSpeakerDistance <= 0, "1.0", My.Settings.LastSpeakerDistance.ToString()),
        .Width = 50,
        .AutoSize = False
    }

            ' Status + partial text
            Me.StatusLabel = New Label() With {
        .Text = "Transcribing:",
        .AutoSize = True,
        .Dock = DockStyle.Top
    }
            Me.PartialTextLabel = New Label() With {
        .Text = "...",
        .AutoSize = True,
        .MinimumSize = New System.Drawing.Size(0, 70),
        .Dock = DockStyle.Top
    }

            ' Action buttons + bottom combobox
            Me.StartButton = New System.Windows.Forms.Button() With {.Text = "Start", .AutoSize = True}
            Me.StopButton = New System.Windows.Forms.Button() With {.Text = "Stop", .AutoSize = True, .Enabled = False}
            Me.ClearButton = New System.Windows.Forms.Button() With {.Text = "Clear", .AutoSize = True}
            Me.LoadButton = New System.Windows.Forms.Button() With {.Text = "Load", .AutoSize = True}
            Me.AudioButton = New System.Windows.Forms.Button() With {.Text = "Dev", .AutoSize = True}
            Me.QuitButton = New System.Windows.Forms.Button() With {.Text = "Quit", .AutoSize = True}
            Me.ProcessButton = New System.Windows.Forms.Button() With {.Text = "Process:", .AutoSize = True}
            Me.processCombobox = New System.Windows.Forms.ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 250
    }

            ' Add a little right‐margin so controls aren’t jammed
            Dim pad As New Padding(0, 0, 10, 0)
            For Each ctl In {Label1, cultureComboBox, Label2, deviceComboBox, SpeakerIdent, SpeakerDistance,
                     StartButton, StopButton, ClearButton, LoadButton, AudioButton, QuitButton, ProcessButton}
                ctl.Margin = pad
            Next
            processCombobox.Margin = pad

            ' --- Build layout ---

            ' Root: 3 rows—top selectors, middle transcript, bottom actions
            Dim root As New TableLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .ColumnCount = 1,
        .RowCount = 3,
        .Padding = New Padding(10)
    }
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row0: selectors
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100)) ' row1: transcript
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))    ' row2: actions

            ' Row 0: selectors laid out in a TableLayoutPanel so combos stretch
            Dim topRow As New TableLayoutPanel() With {
        .Dock = DockStyle.Top,
        .AutoSize = False,
        .Height = cultureComboBox.PreferredHeight + 10,
        .ColumnCount = 6,
        .RowCount = 1,
        .Padding = New Padding(0, 0, 0, 10)
    }
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            topRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            cultureComboBox.Dock = DockStyle.Fill
            deviceComboBox.Dock = DockStyle.Fill

            topRow.Controls.Add(Label1, 0, 0)
            topRow.Controls.Add(cultureComboBox, 1, 0)
            topRow.Controls.Add(Label2, 2, 0)
            topRow.Controls.Add(deviceComboBox, 3, 0)
            topRow.Controls.Add(SpeakerIdent, 4, 0)
            topRow.Controls.Add(SpeakerDistance, 5, 0)

            root.Controls.Add(topRow, 0, 0)

            ' Row 1: status, partial, then main RichTextBox
            Dim mid As New TableLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .ColumnCount = 1,
        .RowCount = 3
    }
            mid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            mid.Controls.Add(StatusLabel, 0, 0)
            mid.Controls.Add(PartialTextLabel, 0, 1)
            mid.Controls.Add(RichTextBox1, 0, 2)

            root.Controls.Add(mid, 0, 1)

            ' Row 2: bottom actions in a stretchy TableLayoutPanel
            Dim bottomRow As New TableLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .AutoSize = False,
        .Height = StartButton.PreferredSize.Height + 20,
        .ColumnCount = 8,
        .RowCount = 1,
        .Padding = New Padding(0, 10, 0, 0)
    }
            ' first six columns auto‐size, last column (processCombobox) fills
            For i = 1 To 7
                bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            Next
            bottomRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            processCombobox.Dock = DockStyle.Fill

            bottomRow.Controls.Add(StartButton, 0, 0)
            bottomRow.Controls.Add(StopButton, 1, 0)
            bottomRow.Controls.Add(ClearButton, 2, 0)
            bottomRow.Controls.Add(LoadButton, 3, 0)
            bottomRow.Controls.Add(AudioButton, 4, 0)
            bottomRow.Controls.Add(QuitButton, 5, 0)
            bottomRow.Controls.Add(ProcessButton, 6, 0)
            bottomRow.Controls.Add(processCombobox, 7, 0)

            root.Controls.Add(bottomRow, 0, 2)

            ' Swap in our root layout
            Me.Controls.Clear()
            Me.Controls.Add(root)

            ' Freeze minimum size once first shown
            AddHandler Me.Shown, Sub() Me.MinimumSize = Me.Size

            ' Set icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        End Sub


        Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
            Dim minWidth As Integer = SpeakerDistance.Left + SpeakerDistance.Width + 40
            If Me.Width < minWidth Then
                Me.Width = minWidth ' Force minimum width dynamically
            End If
        End Sub

        Private Async Function StopRecording() As System.Threading.Tasks.Task

            If loopbackCapture IsNot Nothing Then
                RemoveHandler loopbackCapture.DataAvailable, AddressOf OnLoopbackDataAvailable
                loopbackCapture.StopRecording()
                loopbackCapture.Dispose()
                loopbackCapture = Nothing
            End If

            If loopbackResampler IsNot Nothing Then
                loopbackResampler.Dispose()
                loopbackResampler = Nothing
                loopbackRawProvider = Nothing
            End If


            If waveIn IsNot Nothing Then
                RemoveHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
                RemoveHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
                waveIn.StopRecording()
                waveIn.Dispose()
                waveIn = Nothing
            End If

            CancelTranscription()

            If STTModel = "google" AndAlso _stream IsNot Nothing Then
                Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            End If

            If WhisperRecognizer IsNot Nothing Then
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper stopped...")
                Await WhisperRecognizer.DisposeAsync()
                WhisperRecognizer = Nothing
            End If

            ' Only release the sleep lock IF we were the ones who set it.
            If _iSetTheSleepLock Then
                ' We are responsible, so we release the lock.
                SetThreadExecutionState(ES_CONTINUOUS)
                _iSetTheSleepLock = False ' Reset our flag
                Debug.WriteLine("This form released the sleep lock.")
            Else
                ' We are not responsible, so we do nothing to the execution state.
                Debug.WriteLine("Another component is managing the sleep lock. This form took no action.")
            End If


        End Function


        Private Sub StopButton_Click(sender As Object, e As EventArgs)

            If Not capturing Then Return

            STTCanceled = True

            ' Verhindere Mehrfachklicks
            Me.StopButton.Enabled = False
            If STTModel <> "vosk" Then
                PartialTextLabel.Text = "Stopping…"
            End If

            System.Threading.Tasks.Task.Run(Async Function()
                                                Try
                                                    Await StopRecording()
                                                    If STTModel = "google" Then StopApiWatchdogTimer()
                                                Catch ex As System.Exception

                                                End Try

                                                Me.Invoke(Sub()
                                                              Me.StartButton.Enabled = True
                                                              Me.LoadButton.Enabled = True
                                                              Me.AudioButton.Enabled = True
                                                              Me.cultureComboBox.Enabled = True
                                                              Me.deviceComboBox.Enabled = True
                                                              Me.SpeakerIdent.Enabled = True
                                                              Me.SpeakerDistance.Enabled = True

                                                              If STTModel = "vosk" Then
                                                                  Addline(PartialTextLabel.Text)
                                                              End If
                                                              PartialTextLabel.Text = String.Empty
                                                          End Sub)
                                            End Function)

            capturing = False

        End Sub


        Private Sub ClearButton_Click(sender As Object, e As EventArgs)
            RichTextBox1.Invoke(Sub()
                                    RichTextBox1.Text = ""
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                End Sub)
        End Sub

        Private Sub FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.Closing

            If e.CloseReason = CloseReason.UserClosing Then
                If capturing Then

                    STTCanceled = True

                    Me.StopButton.Enabled = False
                    Me.AudioButton.Enabled = False
                    Me.QuitButton.Enabled = False
                    If STTModel <> "vosk" Then
                        PartialTextLabel.Text = "Stopping…"
                    End If

                    System.Threading.Tasks.Task.Run(Async Function()
                                                        Try
                                                            Await StopRecording()
                                                            If STTModel = "google" Then StopApiWatchdogTimer()
                                                        Catch ex As System.Exception

                                                        End Try

                                                        Me.Invoke(Sub()
                                                                      Me.StartButton.Enabled = False
                                                                      Me.LoadButton.Enabled = False

                                                                      If STTModel = "vosk" Then
                                                                          Addline(PartialTextLabel.Text)
                                                                      End If
                                                                      PartialTextLabel.Text = String.Empty
                                                                  End Sub)
                                                    End Function)

                    capturing = False

                End If
            End If
        End Sub

        Private Sub AudioButton_Click(sender As Object, e As EventArgs)
            ConfigureAudioOutputDevice()
        End Sub

        Private Sub QuitButton_Click(sender As Object, e As EventArgs)

            If capturing Then

                STTCanceled = True

                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = False
                Me.QuitButton.Enabled = False
                If STTModel <> "vosk" Then
                    PartialTextLabel.Text = "Stopping…"
                End If

                System.Threading.Tasks.Task.Run(Async Function()
                                                    Try
                                                        Await StopRecording()
                                                    Catch ex As System.Exception

                                                    End Try

                                                    Me.Invoke(Sub()
                                                                  Me.StartButton.Enabled = False
                                                                  Me.LoadButton.Enabled = False

                                                                  If STTModel = "vosk" Then
                                                                      Addline(PartialTextLabel.Text)
                                                                  End If
                                                                  PartialTextLabel.Text = String.Empty
                                                              End Sub)
                                                End Function)

                capturing = False

            End If
            Me.Close()
        End Sub

        Private Async Sub LoadButton_Click(sender As Object, e As EventArgs)
            If capturing Then Return

            Dim filepath As String = ""

            DragDropFormLabel = "Supported are audio files (*.wav, *.mp3, *.aac, *.m4a, *.mp4 and *.wma)"
            DragDropFormFilter = "Supported Files|*.wav;*.mp3;*.aac;*.m4a;*.mp4;*.wma|" &
                             "Wave files (*.wav)|*.wav|" &
                             "MP3 files (*.mp3)|*.mp3|" &
                             "AAC files (*.aac, *.m4a, *.mp4)|*.aac;*.m4a;*.mp4|" &
                             "WMA files (*.wma)|*.wma|" &
                             "All files|*.*"

            Using form As New DragDropForm()

                If form.ShowDialog() = DialogResult.OK Then
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    filepath = form.SelectedFilePath
                    If Not File.Exists(filepath) Then
                        ShowCustomMessageBox("The selected file was not found.")
                        Return
                    End If
                Else
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    Return
                End If
            End Using
            DragDropFormLabel = ""
            DragDropFormFilter = ""

            Dim splash As New Slib.Splashscreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        ' Ask user for language code
                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "ESC", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            Return
                        End If

                        If Not GoogleSTTsupportedLanguages.Any(Function(code) code.Trim().Normalize().IndexOf(language, StringComparison.OrdinalIgnoreCase) = 0) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages))
                            Return
                        End If

                        ' Configure the streaming recognizer

                        GoogleLanguageCode = language

                    Case "vosk"
                        StartVosk()

                    Case "whisper"

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If

                        StartWhisper(language)
                        STTCanceled = False
                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")

                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select

                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

                Select Case STTModel
                    Case "google"
                        googleTranscriptStart = RichTextBox1.TextLength
                        Dim methodChoice As Integer = ShowCustomYesNoBox("Select your Google transcription method (you may have to try which one works better):", "Send chunks (faster)", "Stream (less gaps)")

                        Debug.WriteLine("Choice = " & methodChoice)

                        If methodChoice = 0 Then
                            splash.Close()
                            Return
                        End If

                        ' Splash schließen, UI ist bereits deaktiviert
                        splash.Close()

                        splash = New Slib.Splashscreen($"Transcribing file ...")
                        splash.Show()
                        splash.Refresh()

                        Try

                            ' Chunking vs. Streaming aufrufen
                            If methodChoice = 1 Then
                                Await GoogleChunkedTranscribeAudioFile(filepath)
                            Else
                                Await GoogleFileStreamTranscription(filepath)
                            End If

                        Catch ex As Exception
                            splash.Close()
                            ShowCustomMessageBox($"Error in Transcribing File using Google: {ex.Message}")
                        Finally
                            splash.Close()
                            Me.Invoke(Sub()
                                          capturing = False
                                          StartButton.Enabled = True
                                          StopButton.Enabled = False
                                          LoadButton.Enabled = True
                                          AudioButton.Enabled = True
                                          cultureComboBox.Enabled = True
                                          deviceComboBox.Enabled = True
                                          SpeakerIdent.Enabled = True
                                          SpeakerDistance.Enabled = True
                                      End Sub)
                        End Try

                    Case "vosk"
                        VoskTranscribeAudioFile(filepath)
                    Case "whisper"
                        WhisperTranscribeAudioFile(filepath)
                        ShowCustomMessageBox($"Transcription using Whisper has started In the background. You can continue working. Do not quit Word. Press 'Stop' to stop transcription.")
                End Select

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an Error starting the transcription engine (Error: {ex.Message}).")

            End Try

        End Sub


        Private Async Sub ProcessButton_Click(sender As Object, e As EventArgs)
            If processCombobox.SelectedIndex >= 0 Then
                Dim selectedIndex As Integer = processCombobox.SelectedIndex
                If selectedIndex < TranscriptPromptsLibrary.Count Then
                    Dim OtherPrompt As String = TranscriptPromptsLibrary(selectedIndex)
                    Dim SelectedText As String = ""
                    If String.IsNullOrWhiteSpace(RichTextBox1.SelectedText) Then
                        SelectedText = RichTextBox1.Text
                    Else
                        SelectedText = RichTextBox1.SelectedText
                    End If
                    Dim LLMResult As String = Await LLM(OtherPrompt & " Current Date is: " & DateTime.Now.ToString("dd MMM yyyy", CultureInfo.CurrentCulture), SelectedText, "", "", 0, False)

                    Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                    Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

                    If wordApp.Documents.Count > 0 Then
                        ' Collapse any existing selection towards the end
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' Insert the markdown text
                        InsertTextWithMarkdown(selection, LLMResult, True)
                    End If
                End If
            End If
        End Sub

        Private Sub LoadAudioDevices()
            deviceComboBox.Items.Clear()
            Dim i As Integer = 0
            For i = 0 To WaveInEvent.DeviceCount - 1
                Dim capabilities = WaveInEvent.GetCapabilities(i)
                Dim micName As String = $"{i}: {capabilities.ProductName}"
                '  a) plain mic
                deviceComboBox.Items.Add(micName)
                '  b) mic + system audio
                deviceComboBox.Items.Add($"{micName} (plus audio output)")
            Next

            ' Select default device (if available)
            Dim lastAudioSource As String = My.Settings.LastAudioSource
            If Not String.IsNullOrEmpty(lastAudioSource) AndAlso deviceComboBox.Items.Contains(lastAudioSource) Then
                deviceComboBox.SelectedItem = lastAudioSource
            ElseIf deviceComboBox.Items.Count > 0 Then
                deviceComboBox.SelectedIndex = 0
            End If
            Dim sel = TryCast(deviceComboBox.SelectedItem, String)
            _multiSourceSelected = (sel IsNot Nothing AndAlso sel.EndsWith("(plus audio output)"))
        End Sub

        Private Sub StartVosk()
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())
            Dim model As New Model(modelpath)
            recognizer = New VoskRecognizer(model, 16000.0F)
            If Me.SpeakerIdent.Checked Then
                ' Get the first available speaker model in the directory
                Dim speakerModelPath As String = System.IO.Directory.GetDirectories(System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\"), "vosk-model*").FirstOrDefault()
                If String.IsNullOrEmpty(speakerModelPath) Then
                    ShowCustomMessageBox($"No speaker model found (at {System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), "Speaker\")}. Speaker recognition will be disabled.")
                    Me.SpeakerIdent.Checked = False
                Else
                    Dim speakerModel As SpkModel = New SpkModel(speakerModelPath)
                    recognizer.SetSpkModel(speakerModel)

                End If

                Debug.WriteLine("Vosk recognizer initialized")

            End If

            recognizer.SetMaxAlternatives(0) ' Forces earlier finalization
            recognizer.SetWords(True) ' Enable word timestamps
            recognizer.SetPartialWords(True) ' Partial words emitted faster
        End Sub

        Private Sub StartWhisper(Optional language As String = "auto")
            Dim modelpath As String = System.IO.Path.Combine(ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath), Me.cultureComboBox.SelectedItem.ToString())

            ' Load the model using WhisperFactory with the specified runtime options
            Dim factory As WhisperFactory = WhisperFactory.FromPath(modelpath)

            ' Configure the builder with language, threads, etc.
            If Me.SpeakerIdent.Checked Then
                Dim builder = factory.CreateBuilder() _
                    .WithLanguage(language) _
                    .WithThreads(Environment.ProcessorCount) _
                    .WithNoSpeechThreshold(Single.Parse(Me.SpeakerDistance.Text)) _
                    .WithTemperature(0.3) _
                    .WithTranslate()

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            Else
                Dim builder = factory.CreateBuilder() _
                    .WithLanguage(language) _
                    .WithThreads(Environment.ProcessorCount) _
                    .WithNoSpeechThreshold(Single.Parse(Me.SpeakerDistance.Text)) _
                    .WithTemperature(0.3)

                ' Build the recognizer
                WhisperRecognizer = builder.Build()
            End If
        End Sub

        Private Async Function StartGoogleSTT() As System.Threading.Tasks.Task
            ' ─── 1) Interceptor definieren, der bei jedem neuen Streaming-Aufruf einen frischen Token holt ───

            Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
                    Async Function(contextCall, metadata)
                        ' Nicht mehr context.GetFresh…, sondern unser lokaler Helper
                        Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                        metadata.Add("Authorization", $"Bearer {tokenToSend}")
                        Await System.Threading.Tasks.Task.CompletedTask
                    End Function
                )

            ' ─── 2) Baue die ChannelCredentials mit Secure SSL + unserem Interceptor ───
            Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
                            Grpc.Core.ChannelCredentials.SecureSsl,
                            callCreds
                        )

            ' ─── 3) Erzeuge einen brandneuen SpeechClient, der das obige channelCreds verwendet ───
            Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                            .Endpoint = STTEndpoint,
                            .ChannelCredentials = channelCreds
                        }
            client = builder.Build()

            ' ─── 4) Öffne die Streaming-Verbindung mit InitializeGoogleStream() ───
            '      Das ruft im Hintergrund “_stream = client.StreamingRecognize() …” und sendet die 
            '      StreamingConfig per WriteAsync. Beim ersten WriteAsync wird der Interceptor aktiv.
            Await InitializeGoogleStream()

            SyncLock ringBuffer
                ringBuffer.Clear()
            End SyncLock

            StartAudioQueueWriter()

        End Function

        Private Sub ResetGoogleStreamFlag()
            _googleStreamCompleted = False
        End Sub

        Private Async Function InitializeGoogleStream() As System.Threading.Tasks.Task

            streamingStartTime = DateTime.UtcNow
            ResetGoogleStreamFlag()

            Try
                ' Bidirektionales Streaming öffnen

                If Me.SpeakerIdent.Checked Then

                    'Dim maxSpk As Integer = CInt(Math.Ceiling(Double.Parse(Me.SpeakerDistance.Text)))

                    Dim minSpeakers As Integer = 2
                    Dim maxSpeakers As Integer = 6 ' Standard-Maximum, anpassen falls nötig

                    ' Versuchen Sie, die Werte aus der UI zu lesen, mit sicheren Standardwerten
                    Try
                        ' Annahme: SpeakerDistance ist jetzt MaxCount und ein neues TextFeld ist MinCount
                        maxSpeakers = CInt(Double.Parse(Me.SpeakerDistance.Text))
                    Catch
                        ' Bei Fehler Standardwerte verwenden
                    End Try

                    ' Die Werte auf den von Google unterstützten Bereich begrenzen
                    minSpeakers = System.Math.Max(2, minSpeakers)
                    maxSpeakers = System.Math.Max(minSpeakers, maxSpeakers)

                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
                        .Config = New RecognitionConfig With {
                            .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                            .SampleRateHertz = 16000,
                            .LanguageCode = GoogleLanguageCode,
                            .EnableAutomaticPunctuation = True,
                            .EnableSpokenPunctuation = True,
                            .EnableWordTimeOffsets = False,
                            .EnableWordConfidence = False,
                            .Model = "latest_long",
                            .UseEnhanced = True,
                            .DiarizationConfig = New SpeakerDiarizationConfig With {
                                .EnableSpeakerDiarization = Me.SpeakerIdent.Checked,
                                        .MinSpeakerCount = minSpeakers,
                                    .MaxSpeakerCount = maxSpeakers
                                                            }
                        },
                        .InterimResults = True,
                        .SingleUtterance = False
                    }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.streamingConfig = streamingConfig})


                Else
                    _stream = client.StreamingRecognize()
                    Dim streamingConfig As New StreamingRecognitionConfig With {
                    .Config = New RecognitionConfig With {
                        .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                        .SampleRateHertz = 16000,
                        .LanguageCode = GoogleLanguageCode,
                        .EnableAutomaticPunctuation = True,
                            .EnableSpokenPunctuation = True,
                            .EnableWordTimeOffsets = False,
                            .EnableWordConfidence = False,
                            .Model = "latest_long",
                            .UseEnhanced = True
                                },
                                .InterimResults = True,
                                .SingleUtterance = False
                            }
                    Await _stream.WriteAsync(New StreamingRecognizeRequest With {.streamingConfig = streamingConfig})
                End If

                'StartAudioQueueWriter()

            Catch ex As System.Exception

                ShowCustomMessageBox("No speaker diarization available for this language (or other error).", $"{GoogleSTT_Desc} Language Code")
                _stream = client.StreamingRecognize()
                Dim streamingConfig As New StreamingRecognitionConfig With {
                .Config = New RecognitionConfig With {
                    .Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                    .SampleRateHertz = 16000,
                    .LanguageCode = GoogleLanguageCode
                            },
                            .InterimResults = True
                        }
                _stream.WriteAsync(New StreamingRecognizeRequest With {.streamingConfig = streamingConfig}).Wait()

            End Try

        End Function

        Private Async Sub StartButton_Click(sender As Object, e As EventArgs)

            If capturing Then
                Return
            End If

            Dim splash As New Slib.Splashscreen($"Loading model...")
            splash.Show()
            splash.Refresh()

            cts = New CancellationTokenSource()
            STTCanceled = False
            audioBuffer.Clear()

            Try
                If Me.cultureComboBox.SelectedItem.ToString().StartsWith(GoogleSTT_Desc) Then
                    STTModel = "google"
                ElseIf Me.cultureComboBox.SelectedItem.ToString().StartsWith("ggml") Then
                    STTModel = "whisper"
                Else
                    STTModel = "vosk"
                End If

                Select Case STTModel

                    Case "google"

                        readerCts = New CancellationTokenSource()

                        Dim language As String = ShowSelectionForm("Select the language code you want to transcribe in:", $"{GoogleSTT_Desc}", GoogleSTTsupportedLanguages)

                        language = language.Trim()

                        ' first handle empty or escape
                        If String.IsNullOrWhiteSpace(language) OrElse String.Equals(language, "esc", StringComparison.OrdinalIgnoreCase) Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If

                        ' now do a true case‑insensitive lookup
                        If Not GoogleSTTsupportedLanguages.Any(
                                Function(code)
                                    Return String.Equals(code, language, StringComparison.OrdinalIgnoreCase)
                                End Function) Then
                            splash.Close()
                            ShowCustomMessageBox("This language code is not supported. Supported are: " & String.Join(", ", GoogleSTTsupportedLanguages), $"{GoogleSTT_Desc} Language Code")
                            STTCanceled = True
                            Return
                        End If

                        Try
                            GoogleLanguageCode = language
                            Await StartGoogleSTT()
                        Catch ex As System.Exception
                            ShowCustomMessageBox("Error starting transcription service: {ex.Message}", $"{GoogleSTT_Desc}")
                            STTCanceled = True
                            Return
                        End Try

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                        googleTranscriptStart = RichTextBox1.TextLength

                        _speakerTagToLabelMap.Clear()
                        _nextSpeakerNumber = 1

                        Me.googleReaderTask = StartGoogleReaderTask()

                    Case "vosk"

                        StartVosk()

                        If Not StartRecording() Then
                            splash.Close()
                            Return
                        End If

                    Case "whisper"

                        ' Define supported ISO 639-1 language codes

                        Dim language As String = ShowCustomInputBox("Enter the language ISO code you want Whisper to transcribe (e.g. en, de, fr, etc.) or go with 'auto':", "Whisper Language Code", True, "auto")

                        language = language.ToLower()

                        If String.IsNullOrWhiteSpace(language) Or language = "esc" Or Not WhisperSupportedLanguages.Contains(language.ToLower()) Then
                            splash.Close()
                            If Not WhisperSupportedLanguages.Contains(language.ToLower()) And language <> "esc" Then
                                ShowCustomMessageBox("This language code is not supported. Supported are: Afrikaans (af), Albanian (sq), Amharic (am), Arabic (ar), Armenian (hy), Assamese (as), Azerbaijani (az), Bashkir (ba), Basque (eu), Belarusian (be), Bengali (bn), Bosnian (bs), Breton (br), Bulgarian (bg), Catalan (ca), Chinese (zh), Croatian (hr), Czech (cs), Danish (da), Dutch (nl), English (en), Estonian (et), Faroese (fo), Finnish (fi), French (fr), Galician (gl), Georgian (ka), German (de), Greek (el), Gujarati (gu), Haitian Creole (ht), Hausa (ha), Hebrew (he), Hindi (hi), Hungarian (hu), Icelandic (is), Indonesian (id), Italian (it), Japanese (ja), Javanese (jv), Kannada (kn), Kazakh (kk), Khmer (km), Kinyarwanda (rw), Kirghiz (ky), Korean (ko), Latvian (lv), Lithuanian (lt), Luxembourgish (lb), Macedonian (mk), Malagasy (mg), Malay (ms), Malayalam (ml), Maltese (mt), Maori (mi), Marathi (mr), Mongolian (mn), Myanmar (my), Nepali (ne), Norwegian (no), Occitan (oc), Pashto (ps), Persian (fa), Polish (pl), Portuguese (pt), Punjabi (pa), Romanian (ro), Russian (ru), Sanskrit (sa), Serbian (sr), Sindhi (sd), Sinhala (si), Slovak (sk), Slovenian (sl), Somali (so), Spanish (es), Sundanese (su), Swahili (sw), Swedish (sv), Tagalog (tl), Tajik (tg), Tamil (ta), Tatar (tt), Telugu (te), Thai (th), Turkish (tr), Ukrainian (uk), Urdu (ur), Uzbek (uz), Vietnamese (vi), Welsh (cy), Yiddish (yi), Yoruba (yo), Zulu (zu)")
                            End If
                            STTCanceled = True
                            Return
                        End If
                        StartWhisper(language)

                        If Not StartRecording() Then
                            splash.Close()
                            STTCanceled = True
                            Return
                        End If
                        STTCanceled = False

                        PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is listening and working... (no partial results shown, please wait)")
                    Case Else
                        splash.Close()
                        ShowCustomMessageBox($"No valid model selected. Please select a model.")
                        Return

                End Select
                My.Settings.LastAudioSource = Me.deviceComboBox.SelectedItem.ToString()
                My.Settings.LastSpeechModel = Me.cultureComboBox.SelectedItem.ToString()
                My.Settings.LastSpeakerEnabled = Me.SpeakerIdent.Checked
                similarityThreshold = Double.Parse(Me.SpeakerDistance.Text)
                If STTModel = "google" Then
                    If similarityThreshold < 1 Then similarityThreshold = 1.0
                Else
                    If similarityThreshold = 0 Then similarityThreshold = 1.0
                    If similarityThreshold < 0.2 Then similarityThreshold = 0.2
                    If similarityThreshold > 2.5 Then similarityThreshold = 2.5
                End If
                My.Settings.LastSpeakerDistance = similarityThreshold

                My.Settings.Save()

                If STTModel = "google" Then StartApiWatchdogTimer()

                capturing = True
                Me.StartButton.Enabled = False
                Me.cultureComboBox.Enabled = False
                Me.deviceComboBox.Enabled = False
                Me.SpeakerIdent.Enabled = False
                Me.SpeakerDistance.Enabled = False
                Me.StopButton.Enabled = True
                Me.LoadButton.Enabled = False
                Me.AudioButton.Enabled = False
                splash.Close()

            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"There has been an error starting the transcription engine (Error: {ex.Message}).")

            End Try
        End Sub


        Private Function StartRecording() As Boolean

            Dim ss As String = TryCast(Me.deviceComboBox.SelectedItem, String)
            Dim deviceIndex As Integer

            Dim pos As Integer = If(ss?.IndexOf(":"c), -1)
            If pos < 0 OrElse Not Integer.TryParse(ss.Substring(0, pos), deviceIndex) Then
                ShowCustomMessageBox($"Invalid device selection: '{ss}'")
                Return False
            End If

            waveIn = New WaveInEvent() With {
                    .DeviceNumber = deviceIndex,
                    .WaveFormat = New WaveFormat(16000, 1)
                }


            If MultiSourceEnabled Then

                ' Versuche, das in den Einstellungen gesetzte Ausgabegerät zu verwenden
                Dim audioOutputDeviceId As String = My.Settings.AudioOutputDevice
                Dim chosenDevice As MMDevice = Nothing

                'Debug.WriteLine("audioOutputDeviceId=" & audioOutputDeviceId)

                If Not String.IsNullOrEmpty(audioOutputDeviceId) Then
                    Try
                        Dim enumerator As New MMDeviceEnumerator()
                        chosenDevice = enumerator.GetDevice(audioOutputDeviceId)
                    Catch ex As System.Exception
                        ' Ungültige ID oder Gerät nicht gefunden → Fallback auf Default
                        chosenDevice = Nothing
                    End Try
                End If

                ' 1) LoopbackCapture mit spezifischem Gerät oder Default erstellen
                If chosenDevice IsNot Nothing Then
                    loopbackCapture = New WasapiLoopbackCapture(chosenDevice)
                Else
                    loopbackCapture = New WasapiLoopbackCapture()
                End If

                ' 2) Raw-Provider in native Format
                loopbackRawProvider = New BufferedWaveProvider(loopbackCapture.WaveFormat) With {
                                .DiscardOnBufferOverflow = True
                            }
                AddHandler loopbackCapture.DataAvailable, Sub(s, ev)
                                                              loopbackRawProvider.AddSamples(ev.Buffer, 0, ev.BytesRecorded)
                                                          End Sub

                ' 3) Resample von native → Mic-Format (16 kHz mono 16-bit)
                loopbackResampler = New MediaFoundationResampler(loopbackRawProvider, waveIn.WaveFormat) With {
                            .ResamplerQuality = 60
                        }

                ' 4) Aufnahme starten
                Try
                    loopbackCapture.StartRecording()
                Catch ex As System.Exception
                    ' Gerät evtl. exklusiv belegt → Fallback auf Mic-only
                    ShowCustomMessageBox("Cannot capture system audio: Device is in exclusive use or invalid. Continuing with mic only.")
                    loopbackCapture.Dispose()
                    loopbackCapture = Nothing
                    loopbackResampler?.Dispose()
                    loopbackResampler = Nothing
                    loopbackRawProvider = Nothing
                End Try
            End If

            If STTModel = "google" Then
                AddHandler waveIn.DataAvailable, AddressOf OnGoogleDataAvailable
            Else
                AddHandler waveIn.DataAvailable, AddressOf OnAudioDataAvailable
            End If
            waveIn.StartRecording()

            ' Always request that the system stay awake.
            ' The function returns the PREVIOUS state.
            Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

            ' Now, check if the SYSTEM_REQUIRED flag was already set in the previous state.
            ' We use a bitwise AND. If the result is 0, the flag was NOT set before our call.
            If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
                ' The lock was NOT active before. Therefore, *we* are responsible for releasing it later.
                _iSetTheSleepLock = True
                Debug.WriteLine("Sleep lock was not active. This form has now acquired it.")
            Else
                ' The lock was ALREADY active. We are not responsible for releasing it.
                _iSetTheSleepLock = False
                Debug.WriteLine("Sleep lock was already active. This form will not release it.")
            End If


            Return True

        End Function

        Private Sub StartApiWatchdogTimer()
            ' Initialize the last response time to now.
            System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

            ' Dispose of any existing timer to prevent orphans.
            _apiWatchdogTimer?.Dispose()

            ' Create a new timer that will call the CheckApiResponse method every 1000ms (1 sec).
            _apiWatchdogTimer = New System.Threading.Timer(
                                        AddressOf CheckApiResponse,
                                        Nothing,
                                        TimeSpan.FromSeconds(1),
                                        TimeSpan.FromSeconds(1)
                                    )
        End Sub

        Private Sub StopApiWatchdogTimer()
            _apiWatchdogTimer?.Dispose()
            _apiWatchdogTimer = Nothing
        End Sub

        Private Sub CheckApiResponse(state As Object)
            ' If we are not capturing, or a recovery is already in progress, do nothing.
            If Not capturing OrElse recoverySemaphore.CurrentCount = 0 Then
                Return
            End If

            ' Atomically read the last response time.
            Dim lastResponseTime As New DateTime(System.Threading.Interlocked.Read(_lastApiResponseTicks))

            ' Check if the elapsed time has exceeded our timeout.
            If (DateTime.UtcNow - lastResponseTime).TotalSeconds > API_RESPONSE_TIMEOUT_SECONDS Then
                ' The API has not responded in time. The stream is likely hung.
                Debug.WriteLine($"[ApiWatchdog] No API response for >{API_RESPONSE_TIMEOUT_SECONDS}s. Forcing stream recovery.")

                ' Stop the timer to prevent it from re-triggering while we recover.
                StopApiWatchdogTimer()

                ' Use our existing thread-safe recovery method to restart the stream.
                ' Then, after recovery, the watchdog will be restarted by the recovery logic itself.
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
            End If
        End Sub

        Private Function StartGoogleReaderTask() As System.Threading.Tasks.Task
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                Catch
                End Try
            End If

            readerCts = New CancellationTokenSource()

            Dim newTask = System.Threading.Tasks.Task.Run(
                                    Async Sub()
                                        Dim token = readerCts.Token
                                        Try
                                            Dim enumerator = _stream.GetResponseStream().GetAsyncEnumerator(token)

                                            While Await enumerator.MoveNextAsync()
                                                System.Threading.Interlocked.Exchange(_lastApiResponseTicks, DateTime.UtcNow.Ticks)

                                                For Each result In enumerator.Current.Results
                                                    If result.IsFinal Then
                                                        ' --- NEW, CORRECTED FINAL RESULT LOGIC ---
                                                        If result.Alternatives.Count > 0 Then
                                                            Dim bestAlternative = result.Alternatives(0)
                                                            Dim finalTranscript As String = bestAlternative.Transcript.Trim()
                                                            _lastKnownPartialResult = ""

                                                            ' Check if we should ignore this result because it's a duplicate of a
                                                            ' partial result that was just committed during recovery.
                                                            If Not String.IsNullOrEmpty(_justCommittedPartialText) AndAlso
                                                               String.Equals(finalTranscript, _justCommittedPartialText.Trim(), StringComparison.OrdinalIgnoreCase) Then

                                                                ' This is a duplicate. Log it, clear the flag, and do nothing more.
                                                                Debug.WriteLine($"[ReaderTask] Ignoring duplicate final result: '{finalTranscript}'")
                                                                _justCommittedPartialText = ""

                                                            Else
                                                                ' This is a new, valid final result. Clear the "ignore" flag and proceed.
                                                                _justCommittedPartialText = ""

                                                                ' Now, apply formatting based on diarization settings.
                                                                If Me.SpeakerIdent.Checked AndAlso bestAlternative.Words.Count > 0 Then

                                                                    Dim currentSegment As New System.Text.StringBuilder()
                                                                    ' Get the label for the very first word's speaker.
                                                                    Dim currentSpeakerLabel As String = GetSpeakerLabel(bestAlternative.Words(0).SpeakerTag)

                                                                    For Each wordInfo In bestAlternative.Words
                                                                        Dim wordSpeakerLabel As String = GetSpeakerLabel(wordInfo.SpeakerTag)

                                                                        If wordSpeakerLabel <> currentSpeakerLabel Then
                                                                            ' The speaker has changed. Commit the previous speaker's segment.
                                                                            Dim segmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                            Addline(segmentToCommit)

                                                                            ' Start a new segment for the new speaker.
                                                                            currentSegment.Clear()
                                                                            currentSpeakerLabel = wordSpeakerLabel
                                                                        End If

                                                                        ' Append the current word to the segment.
                                                                        currentSegment.Append(wordInfo.Word & " ")
                                                                    Next

                                                                    ' After the loop, commit the final segment.
                                                                    If currentSegment.Length > 0 Then
                                                                        Dim finalSegmentToCommit As String = $"{currentSpeakerLabel}: {currentSegment.ToString().Trim()}"
                                                                        Addline(finalSegmentToCommit)
                                                                    End If
                                                                Else
                                                                    ' --- Standard non-diarization logic ---
                                                                    ' Just use the simple Addline method with the raw transcript.
                                                                    Addline(finalTranscript)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        ' --- Interim result logic (remains the same) ---
                                                        If result.Alternatives.Count > 0 Then
                                                            Dim partialTranscript = result.Alternatives(0).Transcript
                                                            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialTranscript)
                                                            _lastKnownPartialResult = partialTranscript
                                                        End If
                                                    End If
                                                Next
                                            End While

                                            ' --- The rest of the Catch blocks remain the same ---
                                        Catch ex As OperationCanceledException
                                            Debug.WriteLine($"[ReaderTask] Gracefully cancelled via OperationCanceledException.")
                                        Catch rex As RpcException
                                            If token.IsCancellationRequested OrElse rex.StatusCode = StatusCode.Cancelled Then
                                                Debug.WriteLine($"[ReaderTask] Gracefully cancelled via RpcException (Status: {rex.StatusCode}).")
                                            Else
                                                Debug.WriteLine($"[ReaderTask] Unexpected RpcException (Status: {rex.StatusCode}). Requesting recovery...")
                                                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                                            End If
                                        Catch ex As Exception
                                            Debug.WriteLine($"[ReaderTask] UNEXPECTED FATAL ERROR: {ex.ToString()}")
                                        End Try
                                    End Sub)
            Return newTask
        End Function

        Private Function GetSpeakerLabel(speakerTag As Integer) As String
            ' Check if we've already seen this tag in this session.
            If _speakerTagToLabelMap.ContainsKey(speakerTag) Then
                ' Yes, return the consistent label we already assigned.
                Return _speakerTagToLabelMap(speakerTag)
            Else
                ' No, this is a new speaker tag. Assign it a new label.
                Dim newLabel As String = $"Speaker {_nextSpeakerNumber}"
                _nextSpeakerNumber += 1

                ' Store the mapping for future use.
                _speakerTagToLabelMap.Add(speakerTag, newLabel)

                Return newLabel
            End If
        End Function



        Private Async Function TryRecoverGoogleStreamAsync() As System.Threading.Tasks.Task

            ' Check if there is a pending partial result that we need to commit.
            If Not String.IsNullOrWhiteSpace(_lastKnownPartialResult) Then
                ' Create a copy to avoid any potential race conditions.
                Dim partialToCommit As String = _lastKnownPartialResult

                ' Reset the class-level variable immediately.
                _lastKnownPartialResult = ""
                _justCommittedPartialText = partialToCommit

                ' Use the existing Addline method to append it to the RichTextBox.
                ' Addline is already thread-safe as it uses Invoke.
                Debug.WriteLine($"[TryRecover] Committing lost partial result: '{partialToCommit}'")
                Addline(partialToCommit)
            End If

            ' Asynchronously wait to acquire the semaphore. If another thread already has it,
            ' this thread will wait here without blocking a thread-pool thread.
            Await recoverySemaphore.WaitAsync()
            Try
                ' Now that we have the lock, perform the actual recovery.
                ' Any other threads calling this method will be waiting on the line above.
                Debug.WriteLine($"[TryRecover] Acquired semaphore. Starting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                Await RecoverGoogleStream()
                streamingStartTime = DateTime.UtcNow ' Reset the timer *after* successful recovery
                Me.Invoke(Sub() StartApiWatchdogTimer())
            Finally
                ' CRITICAL: Always release the semaphore in a Finally block to prevent deadlocks.
                recoverySemaphore.Release()
                Debug.WriteLine($"[TryRecover] Released semaphore. ts={DateTime.UtcNow:HH:mm:ss.fff}")
            End Try
        End Function

        Private Sub StartAudioQueueWriter()
            writerTask = System.Threading.Tasks.Task.Run(
                                        Async Sub()
                                            Try
                                                ' This loop will automatically exit when the queue is completed by
                                                ' SafeCompleteAndDisposeGoogleStreamAsync.
                                                For Each chunk As Google.Protobuf.ByteString In audioQueue.GetConsumingEnumerable()
                                                    Try
                                                        ' If the stream was disposed during recovery, exit the writer immediately.
                                                        If _stream Is Nothing Then
                                                            Debug.WriteLine("[Writer] Stream is null. Exiting task.")
                                                            Return
                                                        End If

                                                        ' Send the audio chunk.
                                                        Await _stream.WriteAsync(New StreamingRecognizeRequest With {.AudioContent = chunk})

                                                    Catch ex As RpcException
                                                        ' A gRPC error occurred (e.g., the stream was cancelled).
                                                        ' This is an expected part of the shutdown/recovery cycle.
                                                        ' We just log it and exit the writer task gracefully.
                                                        Debug.WriteLine($"[Writer] RpcException (Status: {ex.StatusCode}). Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As NullReferenceException
                                                        ' This can happen if _stream is set to Nothing by another thread.
                                                        Debug.WriteLine("[Writer] Stream became null. Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As InvalidOperationException
                                                        ' This can happen if the stream is used after being closed.
                                                        Debug.WriteLine($"[Writer] InvalidOperationException (likely closed stream). Exiting writer task.")
                                                        Return ' Exit the task.

                                                    Catch ex As Exception
                                                        ' Catch any other unexpected error.
                                                        Debug.WriteLine($"[Writer] Unhandled exception in write loop: {ex.GetType().Name}. Exiting writer task.")
                                                        Return ' Exit the task gracefully.
                                                    End Try
                                                Next

                                            Catch ex As InvalidOperationException
                                                ' This exception occurs if GetConsumingEnumerable is called on a collection
                                                ' that has already been marked as complete and then disposed. This is an
                                                ' expected and normal part of the recovery cycle.
                                                Debug.WriteLine("[Writer] Task ending gracefully due to completed or disposed audio queue.")

                                            Catch ex As Exception
                                                ' A truly unexpected error occurred at the task level.
                                                Debug.WriteLine($"[Writer] UNEXPECTED FATAL ERROR in writer task: {ex.ToString()}")
                                            End Try
                                        End Sub)
        End Sub

        Private Async Sub OnGoogleDataAvailable(sender As Object, e As WaveInEventArgs)
            If _googleStreamCompleted Then Return

            'Debug.WriteLine($"[OnGoogleDataAvailable] start  ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

            Dim now = DateTime.UtcNow
            Dim elapsed = (now - streamingStartTime).TotalMilliseconds

            ' ——————— 1) MIX IN LOOPBACK, falls aktiviert ———————
            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim chunk As Google.Protobuf.ByteString =
                Google.Protobuf.ByteString.CopyFrom(e.Buffer, 0, e.BytesRecorded)

            ' 1) Ins Ring-Buffer schreiben (maximal 50 Chunks)
            SyncLock ringBuffer
                ringBuffer.Enqueue(chunk)
                If ringBuffer.Count > RING_BUFFER_SIZE Then ringBuffer.Dequeue()
            End SyncLock

            'Debug.WriteLine($"[OnGoogleDataAvailable] afterRing   ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count}")

            ' 2) In die Queue schreiben (sofern offen)
            If Not audioQueue.IsAddingCompleted Then
                audioQueue.Add(chunk)
            End If

            'Debug.WriteLine($"[OnGoogleDataAvailable] afterQueue  ts={DateTime.UtcNow:HH:mm:ss.fff} queue={audioQueue.Count}")

            ' 3) Timeout prüfen und global recovern
            If elapsed > STREAMING_LIMIT_MS Then
                Debug.WriteLine($"[OnGoogleDataAvailable] Timeout detected. Requesting recovery... ts={DateTime.UtcNow:HH:mm:ss.fff}")
                ' Fire-and-forget the recovery task so we don't block the audio processing event.
                System.Threading.Tasks.Task.Run(Async Sub() Await TryRecoverGoogleStreamAsync())
                ' We now reset the timer *inside* the safe recovery method, not here.
                streamingStartTime = DateTime.UtcNow ' Resetting here is fine to prevent this from firing again immediately
                Return
            End If


        End Sub


        Private Async Function RecoverGoogleStream() As System.Threading.Tasks.Task
            Debug.WriteLine($"[RecoverGoogleStream] Starting...")

            ' --- 1. SHUTDOWN OLD COMPONENTS ---
            ' Store the old task before we overwrite the class-level variable.
            Dim oldReaderTask As System.Threading.Tasks.Task = Me.googleReaderTask

            ' Cancel the old reader's token source.
            If readerCts IsNot Nothing Then
                Try
                    readerCts.Cancel()
                    Debug.WriteLine($"[RecoverGoogleStream] Old CancellationTokenSource cancelled.")
                Catch ex As Exception
                    ' Ignore
                End Try
            End If

            ' Gracefully complete and dispose of the old stream object.
            ' This will help the old reader task exit cleanly.
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Debug.WriteLine($"[RecoverGoogleStream] Old stream disposed.")

            ' Now, explicitly wait for the old reader task to finish.
            ' This is the KEY to preventing the race condition.
            If oldReaderTask IsNot Nothing Then
                Try
                    Await oldReaderTask
                    Debug.WriteLine($"[RecoverGoogleStream] Old reader task has completed.")
                Catch ex As Exception
                    ' We expect exceptions here (like TaskCanceled), so we just log and continue.
                    Debug.WriteLine($"[RecoverGoogleStream] Awaiting old reader task threw: {ex.GetType().Name}")
                End Try
            End If

            ' --- 2. INITIALIZE NEW COMPONENTS ---

            ' This block is mostly the same, creating the new client and stream.
            Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
            Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {newToken}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function)
            Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
            Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
                                .Endpoint = STTEndpoint,
                                .ChannelCredentials = channelCreds
                            }
            client = builder.Build()

            ' Initialize the new stream.
            ' It's important that the call to InitializeGoogleStream happens *after*
            ' the old stream and reader are fully dead.
            ResetGoogleStreamFlag()
            Await InitializeGoogleStream()
            Debug.WriteLine($"[RecoverGoogleStream] New stream initialized.")

            ' --- 3. START NEW TASKS ---

            ' First, start the writer task. It needs a fresh audioQueue.
            ' You have a bug in SafeCompleteAndDisposeGoogleStreamAsync where you permanently close the queue.
            ' Let's fix that too. First, we need a NEW audio queue.
            audioQueue = New System.Collections.Concurrent.BlockingCollection(Of ByteString)()
            StartAudioQueueWriter()
            Debug.WriteLine($"[RecoverGoogleStream] New writer task started.")

            ' Now that everything old is gone, start the new reader task
            ' and assign it to our class-level variable.
            Me.googleReaderTask = StartGoogleReaderTask()
            Debug.WriteLine($"[RecoverGoogleStream] New reader task started.")
        End Function


        Private Async Function xRecoverGoogleStream() As System.Threading.Tasks.Task

            'Debug.WriteLine($"[RecoverGoogleStream] start      ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

            Try
                ' ─── 1) Neuer Token & Client wie gehabt ───
                Dim newToken As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {newToken}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function)
                Dim channelCreds = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl, callCreds)
                Dim builder = New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
            .Endpoint = STTEndpoint,
            .ChannelCredentials = channelCreds
        }
                client = builder.Build()

                ' ─── 2) Stream neu initialisieren ───
                streamingStartTime = DateTime.UtcNow
                ResetGoogleStreamFlag()
                Await InitializeGoogleStream()

                'Debug.WriteLine($"[RecoverGoogleStream] inited     ts={DateTime.UtcNow:HH:mm:ss.fff}")

                ' ─── 3) Offset zurücksetzen ───
                Dim offset As Integer = 0
                Me.Invoke(Sub() offset = RichTextBox1.TextLength)
                googleTranscriptStart = offset

                ' ─── 4) Ring-Buffer wieder in die Queue spielen ───
                SyncLock ringBuffer
                    For Each oldChunk In ringBuffer
                        audioQueue.Add(oldChunk)
                    Next
                End SyncLock

                'Debug.WriteLine($"[RecoverGoogleStream] requeued   ts={DateTime.UtcNow:HH:mm:ss.fff} ring={ringBuffer.Count} queue={audioQueue.Count}")

                ' ─── 5) Reader neu starten ───
                StartGoogleReaderTask()

                SyncLock ringBuffer
                    ringBuffer.Clear()
                End SyncLock


                'Debug.WriteLine($"[RecoverGoogleStream] completed  ts={DateTime.UtcNow:HH:mm:ss.fff}")


            Catch ex As System.Exception
                'Debug.WriteLine($"[RecoverGoogleStream] ERROR      ts={DateTime.UtcNow:HH:mm:ss.fff} ex={ex.Message}")

            End Try
        End Function



        Private Async Function xxSafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Beende den Stream sauber
            Try
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted Then
                    Await _stream.WriteCompleteAsync()
                    _googleStreamCompleted = True
                    ' ► KEIN CompleteAdding() hier!
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"Error in SafeComplete…: {ex.Message}")
            End Try

            ' 2) Jetzt erst: Queue schließen und auf writerTask warten
            audioQueue.CompleteAdding()
            Await writerTask

            ' 3) Finally: Stream-Objekt freigeben
            If _stream IsNot Nothing Then
                _stream.Dispose()
                _stream = Nothing
            End If
        End Function

        Private Async Function SafeCompleteAndDisposeGoogleStreamAsync(token As CancellationToken) As System.Threading.Tasks.Task
            ' 1) Beende den Stream sauber
            Try
                ' ONLY try to complete the stream if it's still valid AND hasn't been forcibly cancelled.
                If _stream IsNot Nothing AndAlso Not _googleStreamCompleted AndAlso Not token.IsCancellationRequested Then
                    Await _stream.WriteCompleteAsync()
                End If
            Catch ex As RpcException When ex.StatusCode = StatusCode.Cancelled
                ' This is an expected exception if the stream was cancelled while we tried to complete it.
                ' We can safely ignore it and proceed with cleanup.
                Debug.WriteLine($"[SafeComplete] Ignored expected RpcException (Cancelled).")
            Catch ex As Exception
                ' Catch other potential errors but don't let them stop the cleanup process.
                Debug.WriteLine($"[SafeComplete] Error during WriteCompleteAsync: {ex.Message}")
            End Try

            _googleStreamCompleted = True

            ' 2) Wait for the writerTask to finish.
            ' It will finish either because the queue was completed or it hit an exception.
            If writerTask IsNot Nothing AndAlso Not writerTask.IsCompleted Then
                Try
                    ' Don't try to complete the queue if it's already done.
                    If Not audioQueue.IsAddingCompleted Then
                        audioQueue.CompleteAdding()
                    End If
                    Await writerTask
                Catch ex As Exception
                    Debug.WriteLine($"[SafeComplete] Error while awaiting writerTask: {ex.Message}")
                End Try
            End If

            ' 3) Finally: Stream-Objekt freigeben
            ' This is a local method call, not the gRPC object. This is safe.
            _stream?.Dispose()
            _stream = Nothing
        End Function



        Private Function ConvertAudioToFloat(buffer As Byte()) As Single()
            ' Each sample = 2 bytes (16-bit), so half as many float samples
            Dim floatArray As Single() = New Single((buffer.Length \ 2) - 1) {}

            ' Convert raw 16-bit PCM -> -1.0f..+1.0f
            For i As Integer = 0 To buffer.Length - 2 Step 2
                Dim sample As Short = BitConverter.ToInt16(buffer, i)
                floatArray(i \ 2) = sample / 32768.0F
            Next

            Return floatArray
        End Function

        Private Sub OnLoopbackDataAvailable(sender As Object, e As WaveInEventArgs)
            ' Buffer the system audio for later mixing
            loopbackBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded)
        End Sub


        Private Async Sub OnAudioDataAvailable(sender As Object, e As WaveInEventArgs)

            If MultiSourceEnabled AndAlso loopbackCapture IsNot Nothing AndAlso loopbackResampler IsNot Nothing Then
                Dim mixBuf(e.BytesRecorded - 1) As Byte
                ' read the same # of bytes from our resampler (16kHz mono 16-bit)
                Dim bytesRead = loopbackResampler.Read(mixBuf, 0, e.BytesRecorded)
                If bytesRead > 0 Then
                    For i As Integer = 0 To bytesRead - 1 Step 2
                        Dim micSample As Integer = BitConverter.ToInt16(e.Buffer, i)
                        Dim outSample As Integer = BitConverter.ToInt16(mixBuf, i)
                        Dim summedSample As Integer = micSample + outSample

                        ' clamp to Int16
                        If summedSample > Short.MaxValue Then summedSample = Short.MaxValue
                        If summedSample < Short.MinValue Then summedSample = Short.MinValue

                        Dim ba() As Byte = BitConverter.GetBytes(CShort(summedSample))
                        e.Buffer(i) = ba(0)
                        e.Buffer(i + 1) = ba(1)
                    Next
                End If
            End If

            Dim buffer As Byte() = e.Buffer
            Dim bytesRecorded As Integer = e.BytesRecorded

            ' Convert to 16-bit PCM samples
            Dim sampleCount As Integer = CInt(bytesRecorded / 2)
            Dim samples(sampleCount - 1) As Single ' Float array for normalized audio

            For i As Integer = 0 To sampleCount - 1
                ' Convert 16-bit PCM to float (-1.0 to 1.0)
                Dim sample As Short = BitConverter.ToInt16(buffer, i * 2)
                Dim floatSample As Single = sample / 32768.0F
                samples(i) = floatSample
            Next

            ' **Normalize Samples**
            Dim maxSample As Single = samples.Max(Function(x) System.Math.Abs(x))
            If maxSample > 0 Then
                Dim gain As Single = 1.0F / maxSample ' Compute normalization factor
                For i As Integer = 0 To sampleCount - 1
                    samples(i) *= gain ' Apply normalization
                Next
            End If

            ' Convert back to 16-bit PCM
            For i As Integer = 0 To sampleCount - 1
                Dim normalizedSample As Short = CShort(samples(i) * 32767)
                Dim bytes As Byte() = BitConverter.GetBytes(normalizedSample)
                buffer(i * 2) = bytes(0)
                buffer(i * 2 + 1) = bytes(1)
            Next

            Select Case STTModel
                Case "vosk"
                    If recognizer IsNot Nothing AndAlso capturing Then
                        Dim jsonResult As String = ""
                        jsonResult = If(recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded),
                                                recognizer.Result, recognizer.PartialResult)
                        ProcessTranscriptionJson(jsonResult)
                    End If

                Case "whisper"

                    If WhisperRecognizer Is Nothing Then Return

                    Try
                        ' Convert audio buffer to float array
                        Dim whispersamples As Single() = ConvertAudioToFloat(e.Buffer)

                        ' Append to buffer
                        audioBuffer.AddRange(whispersamples)
                        ' Only process when buffer has enough data 
                        If audioBuffer.Count < 32000 Then Return ' Adjust threshold based on sample rate
                        ' Copy buffered audio and clear buffer
                        Dim processSamples = audioBuffer.ToArray()
                        audioBuffer.Clear()
                        e.Buffer.Initialize() ' Clear the buffer    

                        ' Process transcription asynchronously
                        Await ProcessWhisper(processSamples)
                    Catch ex As Exception
                        Debug.WriteLine($"Error in OnAudioDataAvailable: {ex.Message}")
                    End Try
            End Select

        End Sub

        Private Async Function ProcessWhisper(samples As Single()) As System.Threading.Tasks.Task
            Try
                If STTCanceled Then Return

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                ' Iterate over the transcription results (only once)
                Dim enumerator = segments.GetAsyncEnumerator()

                If Await enumerator.MoveNextAsync() Then ' Only process the first result batch
                    Dim result As SegmentData = enumerator.Current
                    Dim text As String = result.Text

                    'Debug.WriteLine(text)
                    text = Regex.Replace(text, "\[.*?\]", String.Empty)
                    text = Regex.Replace(text, "\*.*?\*", String.Empty)

                    If Not String.IsNullOrWhiteSpace(text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If
                End If

                Await enumerator.DisposeAsync()

            Catch ex As Exception
                Debug.WriteLine($"Error in ProcessWhisper: {ex.Message}")
            End Try
        End Function

        Public Async Function WhisperTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task

            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Whisper is reading and transcribing your file...")

                Dim samples As Single() = LoadAudioToFloatArray(filepath) ' Use LoadAudioToFloatArray for MP3/FLAC

                Dim segments As IAsyncEnumerable(Of SegmentData) = WhisperRecognizer.ProcessAsync(samples)

                Dim enumerator = segments.GetAsyncEnumerator()

                Dim Exited As Boolean = False

                While Await enumerator.MoveNextAsync()

                    If cts.Token.IsCancellationRequested Then
                        Exited = True
                        Exit While
                    End If

                    Dim result As SegmentData = enumerator.Current
                    Dim Text = result.Text
                    If Not String.IsNullOrWhiteSpace(Text) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(Text & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                End While
                Await enumerator.DisposeAsync()

                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.AudioButton.Enabled = True
                Me.LoadButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If

            Catch ex As Exception
                Debug.WriteLine($"Error in WhisperTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        Public Sub CancelTranscription()
            If cts IsNot Nothing Then
                cts.Cancel()
            End If
        End Sub


        Public Function LoadAudioToFloatArray(filepath As String) As Single()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono (Whisper requires this format)
                Dim waveFormat = New WaveFormat(16000, 1) ' 16kHz, Mono
                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Convert to floating point explicitly
                    Dim floatProvider As ISampleProvider = resampler.ToSampleProvider()

                    ' Read audio data into a floating-point array
                    Dim samples As New List(Of Single)()
                    Dim buffer As Single() = New Single(1024 - 1) {} ' Buffer for PCM float samples
                    Dim samplesRead As Integer

                    Do
                        samplesRead = floatProvider.Read(buffer, 0, buffer.Length)
                        If samplesRead > 0 Then
                            samples.AddRange(buffer.Take(samplesRead))
                        End If
                    Loop While samplesRead > 0

                    Return samples.ToArray()
                End Using
            End Using
        End Function


        ''' <summary>
        ''' Teilt eine Audiodatei in <60 s-Slices, ruft RecognizeAsync auf und beendet sich danach selbst.
        ''' </summary>
        Public Async Function GoogleChunkedTranscribeAudioFile(filepath As String) _
        As System.Threading.Tasks.Task

            ' ─── 0) Stelle sicher, dass client initialisiert ist ───
            If client Is Nothing Then
                Dim tokenToSend As String = Await GetFreshSTTToken(STTSecondAPI)
                Dim callCreds As Grpc.Core.CallCredentials = Grpc.Core.CallCredentials.FromInterceptor(
            Async Function(contextCall, metadata)
                metadata.Add("Authorization", $"Bearer {tokenToSend}")
                Await System.Threading.Tasks.Task.CompletedTask
            End Function
        )
                Dim channelCreds As Grpc.Core.ChannelCredentials = Grpc.Core.ChannelCredentials.Create(
            Grpc.Core.ChannelCredentials.SecureSsl,
            callCreds
        )
                Dim builder As New Google.Cloud.Speech.V1.SpeechClientBuilder() With {
            .Endpoint = STTEndpoint,
            .ChannelCredentials = channelCreds
        }
                client = builder.Build()
            End If

            ' ─── 1) Lade PCM-Daten (16 kHz, mono, 16 Bit) ───
            Dim pcmData As Byte() = LoadAudioToPCM(filepath)

            ' ─── 2) Chunk-Parameter ───
            Dim bytesPerSec As Integer = 16000 * 2      ' 32 000 B/s
            Dim sliceLenSec As Integer = 50
            Dim overlapSec As Integer = 2
            Dim sliceSize As Integer = sliceLenSec * bytesPerSec
            Dim overlapSize As Integer = overlapSec * bytesPerSec
            Dim offset As Integer = 0

            ' ─── 3) Schleife über alle Slices ───
            While offset < pcmData.Length AndAlso Not STTCanceled
                Dim endPos = System.Math.Min(offset + sliceSize, pcmData.Length)
                Dim slice(endPos - offset - 1) As Byte
                Array.Copy(pcmData, offset, slice, 0, endPos - offset)

                ' ─── 4) RecognitionConfig bauen ───
                Dim config As New Google.Cloud.Speech.V1.RecognitionConfig With {
            .Encoding = Google.Cloud.Speech.V1.RecognitionConfig.Types.AudioEncoding.Linear16,
            .SampleRateHertz = 16000,
            .LanguageCode = GoogleLanguageCode,
            .EnableAutomaticPunctuation = True,
            .Model = "latest_long",
            .UseEnhanced = True
        }
                Dim audio As Google.Cloud.Speech.V1.RecognitionAudio =
            Google.Cloud.Speech.V1.RecognitionAudio.FromBytes(slice)

                ' ─── 5) Sync-API-Call ───
                Dim response As Google.Cloud.Speech.V1.RecognizeResponse =
            Await client.RecognizeAsync(config, audio)

                ' ─── 6) Ergebnisse anhängen ───
                For Each result As Google.Cloud.Speech.V1.SpeechRecognitionResult In response.Results
                    If result.Alternatives.Count > 0 Then
                        Addline(result.Alternatives(0).Transcript)
                    End If
                Next

                ' ─── 7) Wenn das letzte Slice war, Abbruch ───
                If endPos >= pcmData.Length Then
                    Exit While
                End If

                ' ansonsten Offset mit Überlappung weiter
                offset = endPos - overlapSize
                If offset < 0 Then offset = 0
            End While

            ' ─── 8) Abschlussmeldung ───
            ShowCustomMessageBox("Chunked transcription complete.", $"{AN} Transcriptor")
        End Function



        ''' <summary>
        ''' Transcribe a local file via StreamingRecognize by feeding the existing
        ''' audioQueue → writerTask → readerTask pipeline, then cleanly shutting down.
        ''' </summary>
        Public Async Function GoogleFileStreamTranscription(filepath As String) As System.Threading.Tasks.Task
            ' 1) UI-Status
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = $"{GoogleSTT_Desc} streaming file…")

            ' 2) Initialisiere den Stream und Reader
            readerCts = New CancellationTokenSource()
            Await StartGoogleSTT()                             ' öffnet _stream & schreibt Config
            googleTranscriptStart = RichTextBox1.TextLength
            googleReaderTask = StartGoogleReaderTask()         ' startet das Lesen der Antworten

            ' 3) Queue & Writer zurücksetzen
            audioQueue = New BlockingCollection(Of Google.Protobuf.ByteString)()
            StartAudioQueueWriter()                            ' schreibt später aus audioQueue in _stream

            ' 4) PCM-Daten laden
            Dim pcmFull As Byte() = LoadAudioToPCM(filepath)

            ' 5) RIFF-Header entfernen, falls WAV
            Dim pcmData = If(
        pcmFull.Length > 44 AndAlso
        System.Text.Encoding.ASCII.GetString(pcmFull, 0, 4) = "RIFF",
        pcmFull.Skip(44).ToArray(),
        pcmFull
    )

            ' 6) In file-Tempo (16 kHz) in die Queue legen
            Const chunkSize As Integer = 4096
            Dim bytesPerSec As Integer = 16000 * 2  ' 16 kHz × 16 Bit Mono = 32 000 B/s
            Dim pos As Integer = 0

            While pos < pcmData.Length AndAlso Not STTCanceled
                Dim len = System.Math.Min(chunkSize, pcmData.Length - pos)
                Dim chunk = Google.Protobuf.ByteString.CopyFrom(pcmData, pos, len)
                audioQueue.Add(chunk)

                ' → hier wird das Tempo gedrosselt:
                Dim delayMs = CInt(1000.0 * len / bytesPerSec)
                Await System.Threading.Tasks.Task.Delay(delayMs)

                pos += len
            End While

            ' 7) Queue schließen → Writer weiß, dass kein Nachschub mehr kommt
            audioQueue.CompleteAdding()

            ' 8) Stream sauber beenden & auf readerTask warten
            Await SafeCompleteAndDisposeGoogleStreamAsync(readerCts.Token)
            Await googleReaderTask

            ' 9) Cleanup & UI wieder freigeben
            StopApiWatchdogTimer()
            PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
            ShowCustomMessageBox("Streaming transcription complete.", $"{AN} Transcriptor")
            Me.Invoke(Sub()
                          capturing = False
                          StartButton.Enabled = True
                          StopButton.Enabled = False
                          LoadButton.Enabled = True
                          AudioButton.Enabled = True
                          cultureComboBox.Enabled = True
                          deviceComboBox.Enabled = True
                          SpeakerIdent.Enabled = True
                          SpeakerDistance.Enabled = True
                      End Sub)
        End Function


        Public Async Function VoskTranscribeAudioFile(filepath As String) As System.Threading.Tasks.Task
            Try
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "Vosk is reading and transcribing your file... press 'Esc' to abort")

                Dim Exited As Boolean = False

                ' Load PCM audio directly (no float conversion needed)
                Dim pcmData As Byte() = LoadAudioToPCM(filepath)

                ' Initialize Vosk recognizer 
                recognizer.Reset()

                ' Stream PCM data to Vosk recognizer
                Dim chunkSize As Integer = 4096 ' Process in small chunks
                Dim offset As Integer = 0

                While offset < pcmData.Length

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Exited = True
                        Exit While
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        ' Exit the loop
                        Exited = True
                        Exit While
                    End If

                    Dim chunkLength As Integer = System.Math.Min(chunkSize, pcmData.Length - offset)
                    Dim chunk As Byte() = pcmData.Skip(offset).Take(chunkLength).ToArray()

                    ' Feed the chunk into the recognizer
                    Dim resultAvailable As Boolean = recognizer.AcceptWaveform(chunk, chunk.Length)

                    ' Retrieve transcription
                    Dim resultText As String
                    If resultAvailable Then
                        Dim resultJson As String = recognizer.Result()
                        resultText = ExtractTextFromJson(resultJson)
                    Else
                        Dim partialJson As String = recognizer.PartialResult()
                        resultText = ExtractTextFromJson(partialJson)
                    End If

                    ' Update UI with transcribed text
                    If Not String.IsNullOrWhiteSpace(resultText) And Not STTCanceled Then
                        Me.Invoke(Sub()
                                      RichTextBox1.AppendText(resultText & vbCrLf)
                                      RichTextBox1.ScrollToCaret()
                                  End Sub)
                    End If

                    offset += chunkLength
                End While

                ' Get final result
                Dim finalResultJson As String = recognizer.FinalResult()
                Dim finalText As String = ExtractTextFromJson(finalResultJson)

                If Not String.IsNullOrWhiteSpace(finalText) Then
                    Me.Invoke(Sub()
                                  RichTextBox1.AppendText(finalText & vbCrLf)
                                  RichTextBox1.ScrollToCaret()
                              End Sub)
                End If

                ' Reset flags and UI
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")
                STTCanceled = True
                Await StopRecording()
                capturing = False
                Me.StartButton.Enabled = True
                Me.StopButton.Enabled = False
                Me.LoadButton.Enabled = True
                Me.AudioButton.Enabled = True
                Me.cultureComboBox.Enabled = True
                Me.deviceComboBox.Enabled = True
                Me.SpeakerIdent.Enabled = True
                Me.SpeakerDistance.Enabled = True
                PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = "")

                If Exited Then
                    ShowCustomMessageBox("Transcription aborted.")
                Else
                    ShowCustomMessageBox("The transcription of your file is complete.")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error in VoskTranscribeAudioFile: {ex.Message}")
            End Try
        End Function

        Private Function ExtractTextFromJson(jsonString As String) As String
            Try
                Dim json As JObject = JObject.Parse(jsonString)
                If json.ContainsKey("text") Then
                    Return json("text").ToString()
                Else
                    Return String.Empty
                End If
            Catch ex As Exception
                Debug.WriteLine($"JSON Parsing Error: {ex.Message}")
                Return String.Empty
            End Try
        End Function


        Public Function LoadAudioToPCM(filepath As String) As Byte()
            Using reader As New MediaFoundationReader(filepath) ' Supports MP3, WAV, FLAC, etc.
                ' Convert audio to 16kHz Mono PCM (Vosk requires this format)
                Dim waveFormat = New WaveFormat(16000, 16, 1) ' 16kHz, 16-bit, Mono

                Using resampler As New MediaFoundationResampler(reader, waveFormat)
                    resampler.ResamplerQuality = 60

                    ' Use MemoryStream to store PCM data
                    Using memoryStream As New MemoryStream()
                        Using pcmWriter As New WaveFileWriter(memoryStream, waveFormat)
                            Dim buffer(4096 - 1) As Byte
                            Dim bytesRead As Integer

                            Do
                                bytesRead = resampler.Read(buffer, 0, buffer.Length)
                                If bytesRead > 0 Then
                                    pcmWriter.Write(buffer, 0, bytesRead)
                                End If
                            Loop While bytesRead > 0

                            pcmWriter.Flush()
                        End Using

                        ' Return raw PCM byte array
                        Return memoryStream.ToArray()
                    End Using
                End Using
            End Using
        End Function

        Private Sub ProcessTranscriptionJson(jsonString As String)
            Try
                Dim jsonObject As JObject = JObject.Parse(jsonString)



                If jsonObject.ContainsKey("text") AndAlso jsonObject("text") IsNot Nothing Then
                    Dim completedLine As String = jsonObject("text").ToString()
                    If Not String.IsNullOrWhiteSpace(completedLine) Then

                        ' Check if speaker embeddings are available
                        If jsonObject.ContainsKey("spk") AndAlso jsonObject("spk").Type = JTokenType.Array Then
                            Dim speakerArray As JArray = jsonObject("spk")
                            Dim speakerEmbedding As List(Of Double) = speakerArray.Select(Function(x) CDbl(x)).ToList()

                            ' Identify the speaker using cosine similarity
                            Dim speakerID As String = IdentifySpeaker(speakerEmbedding)
                            completedLine = $"{speakerID}: " & completedLine
                        End If

                        ' Add line to UI or output
                        Addline(completedLine)
                    End If
                ElseIf jsonObject.ContainsKey("partial") AndAlso jsonObject("partial") IsNot Nothing Then
                    partialText = jsonObject("partial").ToString()
                    PartialTextLabel.Invoke(Sub() PartialTextLabel.Text = partialText)
                End If

            Catch ex As Exception
                MessageBox.Show("Error in ProcessTranscriptionJson: " & ex.Message, "Error")
            End Try
        End Sub

        ' Dictionary to store multiple embeddings per speaker for better matching
        Dim knownSpeakers As New Dictionary(Of String, List(Of List(Of Double)))
        Dim similarityThreshold As Double = 1.0 ' Adjusted for Euclidean Distance

        Private Function IdentifySpeaker(newEmbedding As List(Of Double)) As String
            ' Normalize new embedding
            newEmbedding = NormalizeEmbedding(newEmbedding)

            Dim bestMatch As String = "Unknown"
            Dim bestDistance As Double = Double.MaxValue

            For Each kvp In knownSpeakers
                Dim existingEmbeddings As List(Of List(Of Double)) = kvp.Value

                ' Compute similarity with the average embedding of the stored speaker
                Dim avgEmbedding As List(Of Double) = GetAverageEmbedding(existingEmbeddings)
                Dim distance As Double = EuclideanDistance(avgEmbedding, newEmbedding)

                ' Consider as the same speaker if distance is below threshold
                If distance < bestDistance AndAlso distance < similarityThreshold Then
                    bestMatch = kvp.Key
                    bestDistance = distance
                End If
            Next

            ' If no match, assign a new speaker ID
            If bestMatch = "Unknown" Then
                Dim newSpeakerID As String = "Speaker " & (knownSpeakers.Count + 1).ToString()
                knownSpeakers(newSpeakerID) = New List(Of List(Of Double)) From {newEmbedding}
                Return newSpeakerID
            Else
                ' Store the new embedding for future matches (stabilizes detection)
                knownSpeakers(bestMatch).Add(newEmbedding)

                ' Limit stored embeddings to the last 5 to prevent memory overuse
                If knownSpeakers(bestMatch).Count > 5 Then
                    knownSpeakers(bestMatch).RemoveAt(0)
                End If

                Return bestMatch
            End If
        End Function

        ' Normalize the embedding (ensures embeddings are comparable)
        Private Function NormalizeEmbedding(embedding As List(Of Double)) As List(Of Double)
            Dim norm As Double = System.Math.Sqrt(embedding.Sum(Function(x) x * x))
            If norm = 0 Then Return embedding
            Return embedding.Select(Function(x) x / norm).ToList()
        End Function

        ' Compute the average embedding
        Private Function GetAverageEmbedding(embeddings As List(Of List(Of Double))) As List(Of Double)
            Dim embeddingSize As Integer = embeddings(0).Count
            Dim avgEmbedding As New List(Of Double)(New Double(embeddingSize - 1) {})

            ' Sum up all embeddings
            For Each emb In embeddings
                For i As Integer = 0 To embeddingSize - 1
                    avgEmbedding(i) += emb(i)
                Next
            Next

            ' Divide by the number of stored embeddings
            For i As Integer = 0 To embeddingSize - 1
                avgEmbedding(i) /= embeddings.Count
            Next

            Return avgEmbedding
        End Function

        ' Compute Euclidean Distance between two speaker embeddings
        Private Function EuclideanDistance(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim sum As Double = 0
            For i As Integer = 0 To vec1.Count - 1
                sum += (vec1(i) - vec2(i)) ^ 2
            Next
            Return System.Math.Sqrt(sum)
        End Function


        ' Function to compute cosine similarity between two speaker embeddings
        Private Function CosineSimilarity(vec1 As List(Of Double), vec2 As List(Of Double)) As Double
            Dim dotProduct As Double = vec1.Zip(vec2, Function(a, b) a * b).Sum()
            Dim magnitude1 As Double = System.Math.Sqrt(vec1.Sum(Function(a) a * a))
            Dim magnitude2 As Double = System.Math.Sqrt(vec2.Sum(Function(b) b * b))

            If magnitude1 = 0 OrElse magnitude2 = 0 Then
                Return 0
            End If

            Return dotProduct / (magnitude1 * magnitude2)
        End Function



        Private Sub Addline(completedline As String)
            completedline = completedline.Trim()

            SyncLock finalText
                finalText.AppendLine(completedline)
            End SyncLock

            ' This block is now deadlock-safe because it only writes to the UI.
            RichTextBox1.Invoke(Sub()
                                    ' Clear the partial text label
                                    PartialTextLabel.Text = ""

                                    ' Append the new completed line. AppendText is generally a safe "write" operation.
                                    RichTextBox1.AppendText(completedline & vbCrLf)

                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub


        Private Sub ReplaceAndAddLine(fullTranscript As String)
            RichTextBox1.Invoke(Sub()
                                    ' 1) select everything from the start index to the end…
                                    RichTextBox1.Select(googleTranscriptStart, RichTextBox1.TextLength - googleTranscriptStart)
                                    ' 2) replace it with the entire new transcript
                                    RichTextBox1.SelectedText = fullTranscript & Environment.NewLine
                                    ' 3) reset the caret to the end
                                    RichTextBox1.SelectionStart = RichTextBox1.Text.Length
                                    RichTextBox1.ScrollToCaret()
                                    If STTModel = "google" Then googleTranscriptStart = RichTextBox1.TextLength
                                End Sub)
        End Sub



        Public Sub LoadAndPopulateProcessComboBox(filePath As String, processComboBox As Forms.ComboBox)
            ' Execute LoadPrompts function
            Dim resultCode As Integer = LoadTranscriptPrompts(ExpandEnvironmentVariables(filePath))

            ' Clear the combo box before populating
            processComboBox.Items.Clear()

            ' Check if prompts were successfully loaded
            If resultCode = 0 AndAlso TranscriptPromptsTitles.Count > 0 Then
                ' Add the titles to the combo box
                For Each title As String In TranscriptPromptsTitles
                    processComboBox.Items.Add(title)
                Next
            End If
        End Sub

        Private Function LoadTranscriptPrompts(filePath As String) As Integer

            ' Initialize the return code to 0 (no error)
            Dim returnCode As Integer = 0

            filePath = ExpandEnvironmentVariables(filePath)

            'Debug.WriteLine($"Filepath = {filePath}")

            Try
                ' Verify the file exists
                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The transcript prompt library file was not found.")
                    Return 1
                End If

                TranscriptPromptsTitles.Clear()
                TranscriptPromptsLibrary.Clear()

                ' Read all lines from the file
                Dim lines = System.IO.File.ReadAllLines(filePath)

                For Each line As String In lines
                    ' Trim leading and trailing spaces
                    Dim trimmedLine = line.Trim()

                    ' Ignore empty lines and lines starting with ';'
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then
                        ' Split the line by the delimiter '|'
                        Dim promptData = trimmedLine.Split("|"c)

                        ' Ensure there are at least two parts (title and prompt)
                        If promptData.Length >= 2 Then
                            Dim title = promptData(0).Trim()
                            Dim prompt = String.Join("|", promptData.Skip(1)).Trim()

                            ' Add title and prompt to the respective lists
                            TranscriptPromptsTitles.Add(title)
                            TranscriptPromptsLibrary.Add(prompt)
                        End If
                    End If
                Next

                ' Check if no prompts were found
                If TranscriptPromptsLibrary.Count = 0 Then
                    returnCode = 3
                    ShowCustomMessageBox("No prompts have been found in the configured transcript prompt library file.")
                End If

            Catch ex As System.IO.FileNotFoundException
                returnCode = 1
                ShowCustomMessageBox("The transcript prompt library file was not found: " & ex.Message)

            Catch ex As IndexOutOfRangeException
                returnCode = 2
                ShowCustomMessageBox("The format of the transcript prompt library file is not correct (is a '|' or text thereafter missing?): " & ex.Message)

            Catch ex As Exception
                returnCode = 99
                ShowCustomMessageBox("An unexpected error occurred while loading transcript prompts: " & ex.Message)
            End Try

            Return returnCode
        End Function

    End Class

    Public Class StopForm
        Inherits Form

        Public Property StopRequested As Boolean = False

        Public Sub New()
            Me.Text = "Transkription stoppen"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Width = 200
            Me.Height = 100

            Dim btnStop As New System.Windows.Forms.Button() With {
            .Text = "Stop",
            .Dock = DockStyle.Fill
        }
            AddHandler btnStop.Click, Sub(s, e)
                                          Me.StopRequested = True
                                          Me.Close()
                                      End Sub

            Me.Controls.Add(btnStop)
        End Sub
    End Class



End Class
