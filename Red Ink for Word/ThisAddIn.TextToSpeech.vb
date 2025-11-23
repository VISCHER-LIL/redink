' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Speech.Synthesis
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports Microsoft.Office.Interop.Word
Imports NAudio.Wave
Imports NetOffice.PowerPointApi
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    Public Async Sub CreatePodcast()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        HostName = My.Settings.Hostname
        GuestName = My.Settings.Guestname
        TargetAudience = My.Settings.TargetAudience
        Duration = My.Settings.Duration
        Language = My.Settings.Language
        DialogueContext = My.Settings.DialogueContext
        ExtraInstructions = My.Settings.ExtraInstructions

        Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Host name", HostName),
                    New SLib.InputParameter("Guest name", GuestName),
                    New SLib.InputParameter("Target audience", TargetAudience),
                    New SLib.InputParameter("Context, background info", DialogueContext),
                    New SLib.InputParameter("Target length", Duration),
                    New SLib.InputParameter("Language of dialogue", Language),
                    New SLib.InputParameter("Extra instructions", ExtraInstructions)
                    }

        If ShowCustomVariableInputForm("Please enter the following parameters to take into account when creating your podcast script:", $"Create Podcast Script", params) Then

            HostName = params(0).Value.ToString()
            GuestName = params(1).Value.ToString()
            TargetAudience = params(2).Value.ToString()
            DialogueContext = params(3).Value.ToString()
            Duration = params(4).Value.ToString()
            Language = params(5).Value.ToString()
            ExtraInstructions = params(6).Value.ToString()

            My.Settings.Hostname = HostName
            My.Settings.Guestname = GuestName
            My.Settings.TargetAudience = TargetAudience
            My.Settings.DialogueContext = DialogueContext
            My.Settings.Duration = Duration
            My.Settings.Language = Language
            My.Settings.ExtraInstructions = ExtraInstructions
            My.Settings.Save()

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Podcast), True, False, False, False, False, 3, True, False, True, False, 0, False, "", True)

        End If

    End Sub



    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetThreadExecutionState(ByVal esFlags As UInteger) As UInteger
    End Function

    Private Const ES_CONTINUOUS As UInteger = &H80000000UI
    Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI

    ' Flag to track if the TTS engine is responsible for the current sleep lock.
    Private Shared _ttsAcquiredTheSleepLock As Boolean = False

    ''' <summary>
    ''' Cooperatively acquires a system sleep lock for TTS operations.
    ''' It checks if a lock is already active before taking responsibility for it.
    ''' </summary>
    Public Shared Sub AcquireTTSSleepLock()
        ' Always request that the system stay awake.
        ' The function returns the PREVIOUS state.
        Dim previousState As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)

        ' Check if the SYSTEM_REQUIRED flag was already set in the previous state.
        If (previousState And ES_SYSTEM_REQUIRED) = 0 Then
            ' The lock was NOT active before. Therefore, the TTS engine is now responsible.
            _ttsAcquiredTheSleepLock = True
            Debug.WriteLine("[TTS] Sleep lock was not active. TTS has now acquired it.")
        Else
            ' The lock was ALREADY active. The TTS engine is not responsible for releasing it.
            _ttsAcquiredTheSleepLock = False
            Debug.WriteLine("[TTS] Sleep lock was already active. TTS will not release it.")
        End If
    End Sub

    ''' <summary>
    ''' Cooperatively releases the system sleep lock, but only if the TTS
    ''' engine was the component that originally acquired it.
    ''' </summary>
    Public Shared Sub ReleaseTTSSleepLock()
        ' Only release the sleep lock IF we were the ones who set it.
        If _ttsAcquiredTheSleepLock Then
            ' We are responsible, so we release the lock.
            SetThreadExecutionState(ES_CONTINUOUS)
            _ttsAcquiredTheSleepLock = False ' Reset our flag
            Debug.WriteLine("[TTS] TTS has released the sleep lock.")
        Else
            ' We are not responsible, so we do nothing.
            Debug.WriteLine("[TTS] Another component is managing the sleep lock. TTS took no action.")
        End If
    End Sub

    Public Enum TTSEngine
        Google = 0
        OpenAI = 1
    End Enum

    Public Shared TTS_SelectedEngine As TTSEngine = TTSEngine.Google

    Public Sub DetectTTSEngines()
        ' — split auth endpoints —

        Dim auth1 As String = ThisAddIn.INI_Endpoint
        Dim auth2 As String = ThisAddIn.INI_Endpoint_2

        ' — split TTS endpoints —
        Dim ttsEps = If(String.IsNullOrEmpty(ThisAddIn.INI_TTSEndpoint),
                     Array.Empty(Of String)(),
                     INI_TTSEndpoint.Split("¦"c))
        Dim tts1 As String = If(ttsEps.Length > 0, ttsEps(0), "")
        Dim tts2 As String = If(ttsEps.Length > 1, ttsEps(1), "")

        ' reset
        TTS_googleAvailable = False : TTS_googleSecondary = False
        TTS_openAIAvailable = False : TTS_openAISecondary = False
        TTS_GoogleEndpoint = "" : TTS_OpenAIEndpoint = ""

        ' — Google (needs OAuth2 flags) —
        If auth1.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = False
        End If
        If auth2.Contains(GoogleIdentifier) AndAlso ThisAddIn.INI_OAuth2_2 Then
            TTS_googleAvailable = True
            TTS_googleSecondary = True
        End If

        ' — OpenAI (no OAuth2) —
        If auth1.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = False
        End If
        If auth2.Contains(OpenAIIdentifier) Then
            TTS_openAIAvailable = True
            TTS_openAISecondary = True
        End If

        ' — assign TTS URIs based on identifier match —
        If tts1.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts1
        If tts2.Contains(GoogleIdentifier) Then TTS_GoogleEndpoint = tts2

        If tts1.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts1
        If tts2.Contains(OpenAIIdentifier) Then TTS_OpenAIEndpoint = tts2

        ' if neither engine auth-configured, bail early
        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
            Return
        End If
    End Sub

    Private Shared Function UseSecondaryFor(engine As TTSEngine) As Boolean
        If engine = TTSEngine.Google Then
            Return TTS_googleSecondary
        Else
            Return TTS_openAISecondary
        End If
    End Function


    ' Token-Cache für TTS
    Private Shared ttsAccessToken1 As String = String.Empty
    Private Shared ttsTokenExpiry1 As DateTime = DateTime.MinValue
    Private Shared ttsAccessToken2 As String = String.Empty
    Private Shared ttsTokenExpiry2 As DateTime = DateTime.MinValue

    Private Shared Async Function GetFreshTTSToken(useSecond As Boolean) _
    As System.Threading.Tasks.Task(Of String)

        Try
            Dim token As String
            Dim expiry As DateTime

            If useSecond Then
                token = ttsAccessToken2
                expiry = ttsTokenExpiry2
            Else
                token = ttsAccessToken1
                expiry = ttsTokenExpiry1
            End If

            ' Wenn kein Token oder abgelaufen, neuen holen
            If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= expiry Then
                ' Parameter je nach gewählter API
                Dim clientEmail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail)
                Dim scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes)
                Dim rawKey = If(useSecond, INI_APIKey_2, INI_APIKey)
                Dim authServer = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint)
                Dim life = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry)

                ' GoogleOAuthHelper konfigurieren
                GoogleOAuthHelper.client_email = clientEmail
                GoogleOAuthHelper.private_key = TranscriptionForm.FormatPrivateKey(rawKey)
                GoogleOAuthHelper.scopes = scopes
                GoogleOAuthHelper.token_uri = authServer
                GoogleOAuthHelper.token_life = life

                ' neuen Token holen
                Dim newToken As String = Await GoogleOAuthHelper.GetAccessToken()
                Dim newExpiry = DateTime.UtcNow.AddSeconds(life - 300)

                If useSecond Then
                    ttsAccessToken2 = newToken
                    ttsTokenExpiry2 = newExpiry
                Else
                    ttsAccessToken1 = newToken
                    ttsTokenExpiry1 = newExpiry
                End If

                token = newToken
            End If

            Return token

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(
            $"Error fetching TTS token: {ex.Message}",
            "TTS Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function

    Public Shared cts As New CancellationTokenSource()

    Private Shared Async Function GenerateOpenAITTSAsync(
        input As String,
        languageCode As String,
        voiceName As String,
        pitch As Double,
        speakingRate As Double
    ) As Task(Of Byte())

        Try

            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim apiKey = If(TTS_openAISecondary, DecodedAPI_2, DecodedAPI)

            Debug.WriteLine($"[TTS] OpenAI endpoint = '{TTS_OpenAIEndpoint}'")
            Debug.WriteLine($"[TTS] OpenAI API Key = '{apiKey}'")

            Using client As New System.Net.Http.HttpClient()
                client.DefaultRequestHeaders.Authorization =
                New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)

                ' build JSON
                Dim j = New JObject From {
                {"model", TTS_OpenAI_Model},
                {"input", input},
                {"voice", voiceName},
                {"response_format", "mp3"},
                {"instructions", ""}
            }

                Dim content = New StringContent(j.ToString(), Encoding.UTF8, "application/json")

                ' POST to the detected OpenAI endpoint
                Dim resp = Await client.PostAsync(TTS_OpenAIEndpoint, content).ConfigureAwait(False)
                If resp.IsSuccessStatusCode Then
                    Return Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                Else
                    Dim err = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Throw New System.Exception($"OpenAI TTS Error {resp.StatusCode}: {err}")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateOpenAITTSAsync: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function



    Public Shared Async Function GenerateAudioFromText(input As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional nossml As Boolean = False, Optional Pitch As Double = 0, Optional SpeakingRate As Double = 1, Optional CurrentPara As String = "") As Task(Of Byte())

        AcquireTTSSleepLock()

        Try

            Dim eng = TTS_SelectedEngine

            If eng = TTSEngine.OpenAI Then
                ' strip off “ — Beschreibung” if present
                Dim rawVoice = voiceName.Split(" "c)(0)
                Return Await GenerateOpenAITTSAsync(input,
                                       languageCode,
                                       rawVoice,
                                       Pitch,
                                       SpeakingRate)
            End If

            Using httpClient As New HttpClient()

                Dim AccessToken As String = Await GetFreshTTSToken(UseSecondaryFor(TTSEngine.Google))
                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If


                If String.IsNullOrEmpty(AccessToken) Then
                    ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                    Return Nothing
                End If

                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim requestBody As JObject

                'Debug.WriteLine(input)

                Dim jsonPayload As String

                If input.Trim().StartsWith("{") Then
                    jsonPayload = input
                Else

                    Dim textlabel As String = "text"
                    Dim ssmlPattern As String = "<[^>]+>"  ' Matches any tag-like structure <...>

                    If nossml Then
                        input = Regex.Replace(input, ssmlPattern, String.Empty)
                    Else
                        If Regex.IsMatch(input, ssmlPattern) Then
                            If Not input.Trim().StartsWith("<speak>") Then
                                input = "<speak>" & input & "</speak>"
                            End If
                            textlabel = "ssml"
                        End If
                    End If

                    ' Process as single-speaker plain text
                    requestBody = New JObject From {
                    {"input", New JObject From {{$"{textlabel}", input}}},
                    {"voice", New JObject From {
                        {"languageCode", languageCode},
                        {"name", voiceName}
                    }},
                    {"audioConfig", New JObject From {
                        {"audioEncoding", "MP3"},
                        {"pitch", Pitch},
                        {"speakingRate", SpeakingRate},
                        {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                    }}
                }
                    jsonPayload = requestBody.ToString()
                End If
                ' Convert payload to JSON
                Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")

                Try
                    ' Make API request

                    If Len(input) > TTSLargeText Then
                        Dim t As New Thread(Sub()
                                                ShowCustomMessageBox("Audio generation has started and runs in the background. Press 'Esc' to abort.).", "", 3, "", True)
                                            End Sub)
                        t.SetApartmentState(ApartmentState.STA)
                        t.Start()
                    End If

                    Dim response As HttpResponseMessage = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content, cts.Token).ConfigureAwait(False)

                    ' Error Handling: Check if API call failed
                    If response Is Nothing Then
                        ShowCustomMessageBox("Error generating audio: No response from Google TTS API.")
                        Return Nothing
                    End If

                    Dim responseString As String = Await response.Content.ReadAsStringAsync()

                    ' Debug output: Show API response for troubleshooting
                    Debug.WriteLine($"Google TTS API Response: {responseString}")

                    If response.IsSuccessStatusCode Then
                        Dim responseJson As JObject = JObject.Parse(responseString)

                        ' Check if "audioContent" exists in response
                        If responseJson.ContainsKey("audioContent") Then
                            Dim audioBase64 As String = responseJson("audioContent").ToString()
                            Return System.Convert.FromBase64String(audioBase64)
                        Else
                            ShowCustomMessageBox("Error generating audio: 'audioContent' not found in response.")
                            Return Nothing
                        End If
                    Else
                        ShowCustomMessageBox($"Error generating audio: API returned status {response.StatusCode}. Response: {responseString}{If(String.IsNullOrEmpty(CurrentPara), "", "Text: " & CurrentPara) & " [in clipboard]"}).")
                        If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(response.StatusCode & vbCrLf & vbCrLf & responseString & vbCrLf & vbCrLf & CurrentPara)
                        Return Nothing
                    End If
                Catch ex As TaskCanceledException
                    ShowCustomMessageBox("Audio generation aborted.")
                    Return Nothing
                Catch ex As Exception
                    MessageBox.Show($"Error in GenerateAudioFromText (HTTP): {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Nothing
                End Try

            End Using
        Catch ex As Exception
            MessageBox.Show($"Error in GenerateAudioFromText: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Finally
            ' If IsNothing(prevExecState) Then ...
            ReleaseTTSSleepLock()
        End Try

    End Function


    Public Function ParseTextToConversation(text As String) As List(Of Tuple(Of String, String))
        Dim conversation As New List(Of Tuple(Of String, String))
        Dim currentSpeaker As String = ""
        Dim currentText As String = ""

        Dim paragraphs As String() = text.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each para As String In paragraphs
            Dim trimmedText As String = para.Trim()
            If String.IsNullOrEmpty(trimmedText) Then Continue For

            ' Check if the paragraph starts with a speaker tag
            Dim newSpeaker As String = ""
            If hostTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "H"
                trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
            ElseIf guestTags.Any(Function(tag) trimmedText.StartsWith(tag, StringComparison.OrdinalIgnoreCase)) Then
                newSpeaker = "G"
                trimmedText = trimmedText.Substring(trimmedText.IndexOf(":"c) + 1).Trim()
            End If

            ' If a new speaker is detected, store the previous entry and start a new one
            If newSpeaker <> "" Then
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
                End If
                currentSpeaker = newSpeaker
                currentText = trimmedText
            Else
                ' Continue the current speaker's dialogue
                If Not String.IsNullOrEmpty(currentSpeaker) Then
                    currentText &= " " & trimmedText
                End If
            End If
        Next

        ' Add the last entry
        If Not String.IsNullOrEmpty(currentSpeaker) Then
            conversation.Add(Tuple.Create(currentSpeaker, currentText.Trim()))
        End If

        Return conversation
    End Function


    Async Sub GenerateAndPlayPodcastAudio(
        conversation As List(Of Tuple(Of String, String)),
        filepath As String,
        languagecode As String,
        hostVoice As String,
        guestVoice As String,
        pitch As Double,
        speakingrate As Double,
        nossml As Boolean
    )

        Try

            Dim outputFiles As New List(Of String)

            ' ensure a valid output path
            If String.IsNullOrWhiteSpace(filepath) Then
                filepath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            End If

            ' defaults
            If String.IsNullOrEmpty(languagecode) Then languagecode = "en-US"
            If String.IsNullOrEmpty(hostVoice) Then hostVoice = "en-US-Studio-O"
            If String.IsNullOrEmpty(guestVoice) Then guestVoice = "en-US-Casual-K"

            Dim Exited As Boolean = False
            Dim eng = TTS_SelectedEngine

            Using httpClient As New HttpClient()
                ' — set Authorization header once, based on engine —
                If eng = TTSEngine.Google Then
                    Debug.WriteLine($"[TTS] Using Google TTS engine with endpoint '{TTS_GoogleEndpoint}'")
                    ' Google: fetch OAuth token
                    Dim token = Await GetFreshTTSToken(TTS_googleSecondary)
                    If String.IsNullOrEmpty(token) Then
                        ShowCustomMessageBox("Error generating audio - authentication failed (no token).")
                        Return
                    End If
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", token)
                Else
                    Debug.WriteLine($"[TTS] Using OpenAI TTS engine with endpoint '{TTS_OpenAIEndpoint}'")
                    ' OpenAI: use API key
                    Dim key = If(TTS_openAISecondary, INI_APIKey_2, INI_APIKey)
                    httpClient.DefaultRequestHeaders.Authorization =
                    New Net.Http.Headers.AuthenticationHeaderValue("Bearer", key)
                End If

                ' start “running in background” message
                Dim t As New Thread(Sub()
                                        ShowCustomMessageBox(
                                        "Audio generation has started and runs in the background. Press 'Esc' to abort.",
                                        "", 3, "", True)
                                    End Sub)
                t.SetApartmentState(ApartmentState.STA)
                t.Start()

                ' process each speaker snippet
                For i = 0 To conversation.Count - 1

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then Exited = True : Exit For
                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then Exited = True : Exit For

                    Dim speaker = conversation(i).Item1
                    Dim text = conversation(i).Item2
                    Dim voice = If(speaker = "H", hostVoice, guestVoice)

                    ' handle SSML stripping/wrapping
                    Dim textlabel = "text"
                    If Not nossml Then
                        If Regex.IsMatch(text, "<[^>]+>") AndAlso Not text.Trim().StartsWith("<speak>") Then
                            text = $"<speak>{text}</speak>"
                            textlabel = "ssml"
                        End If
                    Else
                        text = Regex.Replace(text, "<[^>]+>", "")
                    End If

                    Dim audioBytes As Byte()

                    If eng = TTSEngine.Google Then
                        ' — Google path —
                        Dim requestBody = New JObject From {
                        {"input", New JObject From {{textlabel, text}}},
                        {"voice", New JObject From {
                            {"languageCode", languagecode},
                            {"name", voice}
                        }},
                        {"audioConfig", New JObject From {
                            {"audioEncoding", "MP3"},
                            {"pitch", pitch},
                            {"speakingRate", speakingrate},
                            {"effectsProfileId", New JArray("small-bluetooth-speaker-class-device")}
                        }}
                    }

                        Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                        Dim resp = Await httpClient.PostAsync(TTS_GoogleEndpoint & "text:synthesize", content)
                        Dim respStr = Await resp.Content.ReadAsStringAsync()
                        Dim respJson = JObject.Parse(respStr)

                        If respJson.ContainsKey("audioContent") Then
                            audioBytes = System.Convert.FromBase64String(respJson("audioContent").ToString())
                        Else
                            ShowCustomMessageBox("Error: no audioContent in Google response.")
                            Continue For
                        End If

                    Else
                        ' — OpenAI path —
                        ' strip off any “ — Beschreibung” from the combo text
                        Dim rawVoice = voice.Split(" "c)(0)
                        audioBytes = Await GenerateOpenAITTSAsync(text, languagecode, rawVoice, pitch, speakingrate)
                    End If

                    Debug.WriteLine($"Generated audio of {audioBytes.Length} for speaker {speaker} ({voice}) with text length {text.Length} characters.")

                    ' save snippet
                    Dim tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_podcast_temp_{i}.mp3")
                    File.WriteAllBytes(tempFile, audioBytes)
                    outputFiles.Add(tempFile)

                    ' throttle
                    Await System.Threading.Tasks.Task.Delay(1000)
                Next

                ' merge & cleanup
                If Not Exited Then MergeAudioFiles(outputFiles, filepath)
                For Each f In outputFiles : File.Delete(f) : Next
            End Using

            If Exited Then
                ShowCustomMessageBox("Multi-speaker audio generation aborted.")
            Else
                If ShowCustomYesNoBox(
                    $"Your multi-speaker audio sequence has been generated ('{filepath}') and is ready to be played. Play it?",
                    "Yes", "No (file remains available)") = 1 Then
                    PlayAudio(filepath)
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error generating podcast audio: {ex.Message}")
        Finally

        End Try
    End Sub

    Public Sub MergeAudioFiles(inputFiles As System.Collections.Generic.List(Of System.String), outputFile As System.String)
        If inputFiles Is Nothing OrElse inputFiles.Count = 0 Then Throw New ArgumentException("No input files.")
        Dim take As Integer = inputFiles.Count

        ' 1) Concatenate to a temporary WAV (uniform PCM)
        Dim tempWav As String = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".wav")
        Dim targetFormat As New NAudio.Wave.WaveFormat(44100, 16, 2) ' 44.1kHz, 16-bit, stereo

        Using writer As New NAudio.Wave.WaveFileWriter(tempWav, targetFormat)
            For i = 0 To take - 1
                Dim inPath = inputFiles(i)
                If Not System.IO.File.Exists(inPath) Then Continue For

                Debug.WriteLine(inPath)

                ' AudioFileReader decodes MP3/WAV/etc. to 32-bit float
                Using src As New NAudio.Wave.AudioFileReader(inPath)
                    ' Resample/convert to the target PCM format
                    Using resampler As New NAudio.Wave.MediaFoundationResampler(src, targetFormat)
                        resampler.ResamplerQuality = 60
                        Dim buffer(8192 - 1) As Byte
                        While True
                            Dim read = resampler.Read(buffer, 0, buffer.Length)
                            If read = 0 Then Exit While
                            writer.Write(buffer, 0, read)
                        End While
                    End Using
                End Using
            Next
        End Using

        ' 2) If caller wants MP3, encode with Media Foundation; otherwise leave as WAV
        Dim ext = System.IO.Path.GetExtension(outputFile)
        If String.Equals(ext, ".mp3", StringComparison.OrdinalIgnoreCase) Then
            Try
                NAudio.MediaFoundation.MediaFoundationApi.Startup()
                Using wavReader As New NAudio.Wave.WaveFileReader(tempWav)
                    ' 192 kbps; adjust as needed
                    NAudio.Wave.MediaFoundationEncoder.EncodeToMp3(wavReader, outputFile, 192000)
                End Using
            Catch ex As Exception
                ' Fallback: deliver WAV if MP3 encoder is unavailable (e.g., Windows N without Media Feature Pack)
                Dim wavFallback = System.IO.Path.ChangeExtension(outputFile, ".wav")
                System.IO.File.Copy(tempWav, wavFallback, True)
                Throw New InvalidOperationException("Media Foundation MP3 encoder unavailable. Wrote WAV instead: " & wavFallback, ex)
            Finally
                NAudio.MediaFoundation.MediaFoundationApi.Shutdown()
                Try : System.IO.File.Delete(tempWav) : Catch : End Try
            End Try
        Else
            ' Caller requested a non-MP3 extension; give them the WAV
            System.IO.File.Copy(tempWav, outputFile, True)
            Try : System.IO.File.Delete(tempWav) : Catch : End Try
        End If
    End Sub


    ' Function to save audio to a file
    Public Shared Sub SaveAudioToFile(audioData As Byte(), filePath As String)
        Try
            If audioData IsNot Nothing AndAlso audioData.Length > 0 Then
                File.WriteAllBytes(filePath, audioData)
                Debug.WriteLine($"Audio file saved: {filePath}")
            Else
                Debug.WriteLine("No audio received.")
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error saving file: {ex.Message}")
        End Try
    End Sub

    ' Function to play the generated MP3 audio using NAudio
    Public Shared Sub PlayAudio(filePath As String)


        Dim splash As New SLib.SplashScreen($"Playing MP3... press 'Esc' to abort")
        If File.Exists(filePath) Then
            splash.Show()
            splash.Refresh()
        End If

        Try

            If File.Exists(filePath) Then

                Using mp3Reader As New Mp3FileReader(filePath)
                    Using waveOut As New WaveOutEvent()
                        waveOut.Init(mp3Reader)
                        waveOut.Play()

                        ' Keep playing until the audio ends
                        While waveOut.PlaybackState = PlaybackState.Playing
                            Thread.Sleep(100)
                            System.Windows.Forms.Application.DoEvents()
                            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                                Exit While
                            End If
                            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                                Exit While
                            End If
                        End While

                        ' Stop playback
                        waveOut.Stop()
                    End Using ' Automatically disposes waveOut
                End Using ' Automatically disposes mp3Reader

                splash.Close()

            Else
                splash.Close()
                ShowCustomMessageBox("Audio file not found.")
            End If
        Catch ex As Exception
            splash.Close()
            ShowCustomMessageBox($"Error playing audio: {ex.Message}")
        End Try
    End Sub

    Shared Async Sub GenerateAndPlayAudio(textToSpeak As String, filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O")

        Dim Temporary As Boolean = (filepath = "")

        Dim audioBytes As Byte() = Await System.Threading.Tasks.Task.Run(Function() GenerateAudioFromText(textToSpeak, languageCode, voiceName).Result)

        Try
            If audioBytes IsNot Nothing Then
                If Temporary Then
                    filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
                End If
                SaveAudioToFile(audioBytes, filepath)
                Dim Result As Integer = 1
                If Len(textToSpeak) > TTSLargeText Then
                    Result = ShowCustomYesNoBox("Your audio sequence has been generated " & If(Temporary, "", $"('{filepath}') ") & "and is ready to be played. Play it?", "Yes", If(Temporary, "No", "No (file remains available)"))
                End If
                If Result = 1 Then
                    PlayAudio(filepath)
                End If
                If Temporary Then
                    System.IO.File.Delete(filepath)
                End If
            End If
        Catch ex As System.Exception

        End Try
    End Sub


    Public Sub ReadPodcast(Text As String)

        Dim NoSSML As Boolean = My.Settings.NoSSML
        Dim Pitch As Double = My.Settings.Pitch
        Dim SpeakingRate As Double = My.Settings.Speakingrate

        ' Create an array of InputParameter objects.
        Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML)
                    }

        Dim conversation As List(Of Tuple(Of String, String)) = ParseTextToConversation(Text)
        Dim hasHost As Boolean = conversation.Any(Function(t) t.Item1 = "H")
        Dim hasGuest As Boolean = conversation.Any(Function(t) t.Item1 = "G")

        If hasHost AndAlso hasGuest Then
            Using frm As New TTSSelectionForm("Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Text-to-Speech - Select Voices", True) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish to use for creating your audio file and configure where to save it.", $"{AN} Google Text-to-Speech - Select Voices", True)
                If frm.ShowDialog() = DialogResult.OK Then
                    Dim selectedVoices As List(Of String) = frm.SelectedVoices
                    Dim selectedLanguage As String = frm.SelectedLanguage
                    Dim outputPath As String = frm.SelectedOutputPath

                    Debug.WriteLine("Voices=" & selectedVoices(0))
                    Debug.WriteLine("TTS_SelectedEngine=" & TTS_SelectedEngine)

                    ' Call the procedure (the parameters are passed ByRef).
                    If ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your podcast audio file:", $"Create Podcast Audio", params) Then

                        ' After OK is clicked, update your original variables:
                        Pitch = CDbl(params(0).Value)
                        SpeakingRate = CDbl(params(1).Value)
                        NoSSML = CBool(params(2).Value)

                        My.Settings.NoSSML = NoSSML
                        My.Settings.Pitch = Pitch
                        My.Settings.Speakingrate = SpeakingRate
                        My.Settings.Save()

                        GenerateAndPlayPodcastAudio(conversation, outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), Pitch, SpeakingRate, NoSSML)
                    End If
                End If
            End Using
        Else
            ' Missing either Host or Guest
            ShowCustomMessageBox($"No conversation was found. Use '{hostTags(0)}' and '{guestTags(0)}' to dedicate content to the host and guest.")
        End If

    End Sub


    Public Async Sub GenerateAndPlayAudioFromSelectionParagraphs(filepath As String, Optional languageCode As String = "en-US", Optional voiceName As String = "en-US-Studio-O", Optional voiceNameAlt As String = "")

        Dim CurrentPara As String = ""

        Try

            Dim Temporary As Boolean = (filepath = "")
            Dim Alternate As Boolean = True

            If Temporary Then
                filepath = System.IO.Path.Combine(ExpandEnvironmentVariables("%TEMP%"), $"{AN2}_temp.mp3")
            End If

            If voiceNameAlt = "" Then Alternate = False

            ' Get the current Word selection.
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = app.Selection
            If selection Is Nothing OrElse selection.Paragraphs.Count = 0 Then
                ShowCustomMessageBox("No text selected.")
                Return
            End If

            Dim NoSSML As Boolean = My.Settings.NoSSML
            Dim Pitch As Double = My.Settings.Pitch
            Dim SpeakingRate As Double = My.Settings.Speakingrate
            Dim ReadTitleNumbers As Boolean = False
            Dim CleanText As Boolean = False
            Dim CleanTextPrompt As String = My.Settings.CleanTextPrompt
            If String.IsNullOrWhiteSpace(CleanTextPrompt) Then CleanTextPrompt = SP_CleanTextPrompt

            ' Create an array of InputParameter objects.
            Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Pitch", Pitch),
                    New SLib.InputParameter("Speaking Rate", SpeakingRate),
                    New SLib.InputParameter("No SSML", NoSSML),
                    New SLib.InputParameter("Title Numbers", ReadTitleNumbers),
                    New SLib.InputParameter("Clean text", CleanText)
                    }

            ' Call the procedure (the parameters are passed ByRef).
            If Not ShowCustomVariableInputForm("Please enter the following parameters to apply when creating your audio file based on your text:", $"Create Audio", params) Then Return

            Pitch = CDbl(params(0).Value)
            SpeakingRate = CDbl(params(1).Value)
            NoSSML = CBool(params(2).Value)
            ReadTitleNumbers = CBool(params(3).Value)
            CleanText = CBool(params(4).Value)

            My.Settings.NoSSML = NoSSML
            My.Settings.Pitch = Pitch
            My.Settings.Speakingrate = SpeakingRate
            My.Settings.Save()

            If CleanText Then
                CleanTextPrompt = ShowCustomInputBox("Please enter the prompt to 'clean' the text with (each paragraph will be submitted to this prompt)", "Create Audio", False, CleanTextPrompt).Trim()
                If CleanTextPrompt = "ESC" Then Return
                If CleanTextPrompt = "" Then
                    CleanText = False
                Else
                    My.Settings.CleanTextPrompt = CleanTextPrompt
                    My.Settings.Save()
                End If
            End If

            Dim totalParagraphs As Integer = selection.Paragraphs.Count
            Dim tempFiles As New List(Of String)
            Dim paragraphIndex As Integer = 0
            Dim sentenceEndPunctuation As String() = {".", "!", "?", ";", ":", ",", ")", "]", "}"}
            Dim bracketedTextPattern As String = "^\s*[\(\[\{][^\)\]\}]*[\)\]\}]\s*$"

            Dim voiceName1 As String = voiceName
            Dim voiceName2 As String = voiceNameAlt
            Dim currentVoiceName As String = voiceName1
            Dim firstTitleEncountered As Boolean = False
            Dim LastTextWasTitle As Boolean = False

            Dim cleanedTextBuilder As New System.Text.StringBuilder()

            ShowProgressBarInSeparateThread($"{AN} Audio Generation", "Starting audio generation...")
            ProgressBarModule.CancelOperation = False

            Dim silenceFileAfterBullet As String = Await GenerateSilenceAudioFileAsync(0.3)
            Dim silenceFileTitle As String = Await GenerateSilenceAudioFileAsync(0.7)
            Dim silenceFileRegular As String = Await GenerateSilenceAudioFileAsync(0.3)

            ' Process each paragraph in the selection.
            For Each para As Microsoft.Office.Interop.Word.Paragraph In selection.Paragraphs
                ' Allow the user to abort by pressing Escape.
                If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Or (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
                    For Each file In tempFiles
                        Try
                            If IO.File.Exists(file) Then IO.File.Delete(file)
                        Catch ex As Exception
                            Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                        End Try
                    Next
                    ShowCustomMessageBox("Audio generation aborted by user.")
                    ProgressBarModule.CancelOperation = True
                    Return
                End If

                ' Get the trimmed paragraph text.
                Dim paraText As String

                ' Check if the paragraph has numbering
                If Not String.IsNullOrEmpty(para.Range.ListFormat.ListString) And ReadTitleNumbers Then
                    ' Include the numbering before the paragraph text
                    paraText = para.Range.ListFormat.ListString.Trim("."c) & vbCrLf & para.Range.Text.Trim()
                Else
                    ' No numbering, just take the paragraph text
                    paraText = para.Range.Text.Trim()
                End If


                ' Skip paragraphs that are empty...
                If String.IsNullOrWhiteSpace(paraText) Or Regex.IsMatch(paraText, bracketedTextPattern) Then Continue For
                ' ...or that contain only numbers or control characters.
                If Regex.IsMatch(paraText, "^[\d\p{C}\s]+$") Then Continue For

                Dim lastChar As String = paraText.Substring(paraText.Length - 1)

                ' Check if the last character is one of the defined punctuation marks
                If Not sentenceEndPunctuation.Contains(lastChar) Then
                    ' Append a period
                    paraText = paraText & "."
                End If

                ' Determine if this paragraph is part of a bullet list.
                Dim isBullet As Boolean = False
                If para.Range.ListFormat IsNot Nothing AndAlso para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering Then
                    isBullet = True
                End If

                ' Determine if the paragraph “looks like” a title.
                Dim isTitle As Boolean = False
                Dim styleName As String = ""
                Try
                    Dim styleObj As Word.Style = TryCast(para.Range.Style, Word.Style)
                    If styleObj IsNot Nothing Then
                        styleName = styleObj.NameLocal.ToLowerInvariant()
                    Else
                        styleName = String.Empty
                    End If
                Catch ex As Exception
                    Debug.WriteLine("Error retrieving style: " & ex.Message)
                End Try
                If styleName.Contains("heading") Then
                    isTitle = True
                Else
                    Dim lineCount As Long = para.Range.ComputeStatistics(WdStatistic.wdStatisticLines)
                    If lineCount <= 2 Then
                        isTitle = True
                    End If
                    If Not paraText.EndsWith(".") Then
                        isTitle = True
                    End If
                End If

                Debug.WriteLine("Para = " & paraText & vbCrLf & vbCrLf)
                Debug.WriteLine("IsTitle = " & isTitle & vbCrLf)
                CurrentPara = Left(paraText, 400) & "..."

                If isTitle AndAlso Alternate Then
                    If Not firstTitleEncountered Then
                        firstTitleEncountered = True
                        ' For the very first title, keep the current voice unchanged.
                    Else
                        If Not LastTextWasTitle Then
                            ' Switch the voice if the last paragraph was not a title.
                            Debug.WriteLine("Switching ...")
                            If currentVoiceName = voiceName1 Then
                                currentVoiceName = voiceName2
                            Else
                                currentVoiceName = voiceName1
                            End If
                        End If
                    End If
                    LastTextWasTitle = True
                Else
                    LastTextWasTitle = False
                End If

                ' Set the maximum value if you know the total number of steps.
                GlobalProgressMax = totalParagraphs

                ' Update the current progress value and status label.
                GlobalProgressValue = paragraphIndex + 1
                GlobalProgressLabel = $"Paragraph {paragraphIndex + 1} of {totalParagraphs} (some may be skipped)"

                ' For bullet lists, insert a short pause BEFORE the paragraph.
                If isBullet Then
                    Dim silenceFileBefore As String = Await GenerateSilenceAudioFileAsync(0.1)
                    If Not String.IsNullOrEmpty(silenceFileBefore) Then tempFiles.Add(silenceFileBefore)
                End If

                If CleanText Then
                    ' Remove any unwanted characters from the paragraph text.
                    paraText = Await LLM(CleanTextPrompt, "<TEXTTOPROCESS>" & paraText & "</TEXTTOPROCESS>", "", "", 0, False, True)
                    paraText = paraText.Trim().Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()
                    CurrentPara = Left(CurrentPara, 100) & $"... [cleaned: {Left(paraText, 400)}...]"
                    Debug.WriteLine("Cleaned Para = " & paraText & vbCrLf & vbCrLf)

                End If

                ' Generate the audio for the paragraph via your TTS API.
                Dim paragraphAudioBytes As Byte() = Await GenerateAudioFromText(paraText, languageCode, currentVoiceName, NoSSML, Pitch, SpeakingRate, CurrentPara)

                CurrentPara = ""

                If paragraphAudioBytes IsNot Nothing Then
                    If CleanText Then
                        cleanedTextBuilder.AppendLine(paraText)
                        cleanedTextBuilder.AppendLine() ' Leerzeile zwischen Absätzen
                    End If
                    Dim tempParaFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_temp_para_{paragraphIndex}.mp3")
                    File.WriteAllBytes(tempParaFile, paragraphAudioBytes)
                    tempFiles.Add(tempParaFile)
                    Debug.WriteLine("Created " & tempParaFile)
                Else
                    ' If audio generation failed, skip this paragraph.
                    Debug.WriteLine("Creation failed")
                    Continue For
                End If

                ' For bullet lists, insert a short pause AFTER the paragraph.
                If isBullet Then

                    If Not String.IsNullOrEmpty(silenceFileAfterBullet) Then tempFiles.Add(silenceFileAfterBullet)
                End If

                ' After each paragraph, add an extra pause:
                ' • Use a medium pause (0.7 sec) for titles.
                ' • Otherwise use a short pause (0.3 sec).
                If isTitle Then
                    If Not String.IsNullOrEmpty(silenceFileTitle) Then tempFiles.Add(silenceFileTitle)
                Else
                    If Not String.IsNullOrEmpty(silenceFileRegular) Then tempFiles.Add(silenceFileRegular)
                End If

                Await System.Threading.Tasks.Task.Delay(1000) ' Delay to not overhwelm the API

                paragraphIndex += 1
            Next

            ' If no valid paragraphs were found, notify the user.
            If tempFiles.Count = 0 Then
                ShowCustomMessageBox("No valid paragraphs found For audio generation; skipping empty ones And {...}, [...] And (...).")
                Return
            End If

            If Not ProgressBarModule.CancelOperation Then
                ' Merge all the temporary audio files into one final file.
                GlobalProgressLabel = $"Merging audio {totalParagraphs} snippets..."
                MergeAudioFiles(tempFiles, filepath)
            End If

            If Not ProgressBarModule.CancelOperation AndAlso CleanText Then
                Try
                    Dim txtPath As String = System.IO.Path.ChangeExtension(filepath, ".txt")
                    System.IO.File.WriteAllText(txtPath, cleanedTextBuilder.ToString(), System.Text.Encoding.UTF8) ' überschreibt ohne Rückfrage
                Catch ex As System.Exception
                    ' Fehler still schlucken, Ablauf geht ungestört weiter
                    Debug.WriteLine("Error writing cleaned text file: " & ex.Message)
                End Try
            End If

            'Cleanup Temporary files.
            For Each file In tempFiles
                Try
                    If IO.File.Exists(file) Then IO.File.Delete(file)
                Catch ex As Exception
                    Debug.WriteLine($"Error deleting temp file {file}: {ex.Message}")
                End Try
            Next

            If Not ProgressBarModule.CancelOperation Then
                ProgressBarModule.CancelOperation = True
                ' Play the merged audio file.
                PlayAudio(filepath)
                If Temporary Then
                    System.IO.File.Delete(filepath)
                End If
            Else
                ProgressBarModule.CancelOperation = True
                ShowCustomMessageBox("Audio generation aborted by user.")
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Error generating audio from selected paragraphs ({ex.Message}{If(String.IsNullOrEmpty(CurrentPara), "", "; Text: " & CurrentPara) & " [in clipboard]"}).")
            If Not String.IsNullOrEmpty(CurrentPara) Then SLib.PutInClipboard(ex.Message & vbCrLf & vbCrLf & CurrentPara)
        End Try
    End Sub

    Private Async Function GenerateSilenceAudioFileAsync(durationSeconds As Double) As Task(Of String)
        Return Await System.Threading.Tasks.Task.Run(Function() GenerateSilenceAudioFile(durationSeconds))
    End Function

    ' Synchronous helper that creates a buffer of silence and encodes it to MP3.
    Private Function GenerateSilenceAudioFile(durationSeconds As Double) As String
        Try
            ' Set audio format parameters.
            Dim sampleRate As Integer = 24000       ' Adjust as needed to match your TTS output.
            Dim channels As Integer = 1
            Dim bitsPerSample As Integer = 16
            Dim blockAlign As Integer = channels * (bitsPerSample \ 8)
            Dim totalSamples As Integer = CInt(sampleRate * durationSeconds)
            Dim totalBytes As Integer = totalSamples * blockAlign

            ' Create a buffer filled with zeros (silence).
            Dim silenceBytes(totalBytes - 1) As Byte
            ' (The array is automatically initialized to zeros.)

            ' Generate a temporary file name.
            Dim tempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{AN2}_silence_{CInt(durationSeconds * 1000)}ms.mp3")

            ' Wrap the silence buffer in a MemoryStream and then a RawSourceWaveStream.
            Using ms As New MemoryStream(silenceBytes)
                Dim waveFormat As New WaveFormat(sampleRate, bitsPerSample, channels)
                Using waveStream As New RawSourceWaveStream(ms, waveFormat)
                    ' Encode the silence to MP3.
                    MediaFoundationEncoder.EncodeToMp3(waveStream, tempFile)
                End Using
            End Using

            Return tempFile
        Catch ex As Exception
            Debug.WriteLine($"Error generating silence audio: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' Legacy Text To Speech

    Private synth As New SpeechSynthesizer()

    Public Shared Sub SelectVoiceByNumber()
        ' Ensure the SpeechSynthesizer is available
        Dim synth As New SpeechSynthesizer()

        ' (1) Retrieve all available voices
        Dim installedVoices As List(Of InstalledVoice) = synth.GetInstalledVoices().ToList()
        Dim voiceNames As New List(Of String)()

        ' (2) Populate voice list
        Dim sb As New StringBuilder()
        sb.AppendLine("Available voices for Text-to-Speech:" & vbCrLf)

        For i As Integer = 0 To installedVoices.Count - 1
            Dim voiceInfo As VoiceInfo = installedVoices(i).VoiceInfo
            voiceNames.Add(voiceInfo.Name)
            sb.AppendLine($"{i}: {voiceInfo.Name}")
        Next

        If voiceNames.Count = 0 Then
            ShowCustomMessageBox("No voices available on this system.", "Text-to-Speech")
            Return
        End If

        Dim UserInput As String = ShowCustomInputBox(sb.ToString(), "Select Voice for Text Reader", True)

        If String.IsNullOrWhiteSpace(UserInput) Then Return

        Dim selectedIndex As Integer
        If Integer.TryParse(UserInput, selectedIndex) AndAlso selectedIndex >= 0 AndAlso selectedIndex < voiceNames.Count Then
            ' Get the selected voice name
            Dim chosenVoice As String = voiceNames(selectedIndex)
            Try
                synth.SelectVoice(chosenVoice)
                My.Settings.LastVoice = chosenVoice
                My.Settings.Save()

                synth.Speak($"Hello! I am now using the voice: {chosenVoice}")
            Catch ex As Exception
                MsgBox("Error selecting voice: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Else
            ShowCustomMessageBox("Invalid voice number entered.", "Text-to-Speech")
        End If
    End Sub

    Public Sub SpeakSelectedText()

        Debug.WriteLine("Status: " & synth.State.ToString())

        If synth.State = SynthesizerState.Speaking Then
            synth.SpeakAsyncCancelAll()
            ShowCustomMessageBox("Reading out aborted.", "Text-to-Speech")
            Return
        End If

        Try
            ' Get the active Word application
            Dim wordApp As Word.Application = Globals.ThisAddIn.Application

            ' Get the selected text
            Dim selectedText As String = wordApp.Selection.Text.Trim()

            If String.IsNullOrEmpty(selectedText) Then
                ShowCustomMessageBox("No text selected in Word.", "Text-to-Speech")
                Return
            End If

            ' Speak the selected text

            synth.SelectVoice(My.Settings.LastVoice)

            synth.SpeakAsync(selectedText)

            ShowCustomMessageBox($"Reading out the selected text (using {My.Settings.LastVoice}). You can stop this by again calling this function.", "Text-to-Speech")

        Catch ex As Exception
            MessageBox.Show("Error in SpeakSelectedText: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



End Class
