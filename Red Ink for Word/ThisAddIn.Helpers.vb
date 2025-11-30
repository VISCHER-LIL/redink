' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn

    Public Function INILoadFail() As Boolean
        If Not INIloaded Then
            If Not StartupInitialized Then
                DelayedStartupTasks()
                RemoveStartupHandlers()
                If Not INIloaded Then Return True
                Return False
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Return True
                End If
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Friend NotInheritable Class WordUndoScope
        Implements System.IDisposable

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private ReadOnly _undo As Microsoft.Office.Interop.Word.UndoRecord
        Private ReadOnly _iStarted As System.Boolean

        Public Sub New(app As Microsoft.Office.Interop.Word.Application, Optional name As System.String = Nothing)
            _app = app
            _undo = _app.UndoRecord

            ' Word < 2013 (Version < 15.0) hat kein UndoRecord.
            Dim ver As System.Version = New System.Version(_app.Version)
            If ver.Major < 15 Then
                Return
            End If

            ' Nur starten, wenn gerade kein anderer Custom-Record läuft
            If Not _undo.IsRecordingCustomRecord Then
                If name IsNot Nothing AndAlso name.Length > 0 Then
                    _undo.StartCustomRecord(name)
                Else
                    _undo.StartCustomRecord("VSTO-Aktion")
                End If
                _iStarted = True
            End If
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Try
                If _iStarted AndAlso _undo.IsRecordingCustomRecord Then
                    _undo.EndCustomRecord()
                End If
            Catch ex As System.Exception
                ' Nichts werfen – wir sind in Dispose
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' Ersetzt jede Sequenz \\X durch \\uXXXX (doppelter Backslash!).
    ''' Aus \\; wird \\u003B, aus \\< wird \\u003C, usw.
    ''' </summary>
    Public Function HideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\(.)",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim c As Char = m.Groups(1).Value(0)
                Dim hex As String = System.Convert.ToInt32(c).ToString("X4")
                Return "\\u" & hex
            End Function)
    End Function

    ''' <summary>
    ''' Ersetzt jede Sequenz \\uXXXX (doppelter Backslash!) zurück in das jeweilige Zeichen.
    ''' Aus \\u003B wird ;, aus \\u003C wird &lt;, usw.
    ''' </summary>
    Public Function UnHideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\u([0-9A-Fa-f]{4})",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim code As Integer = Integer.Parse(m.Groups(1).Value, System.Globalization.NumberStyles.HexNumber)
                Return System.Convert.ToChar(code).ToString()
            End Function)
    End Function

    Public Function GatherSelectedDocuments(Optional IncludeName As Boolean = True,
                                            Optional IncludeNone As System.Boolean = False,
                                            Optional ExceptCurrent As Boolean = False,
                                            Optional SilentAndGetAll As Boolean = False) As System.String
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Collect all open documents (unique by FullName/Name to avoid duplicates from multiple windows)
            Dim docList As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            For Each doc As Microsoft.Office.Interop.Word.Document In app.Documents
                Dim key As System.String = If(Not System.String.IsNullOrEmpty(doc.FullName), doc.FullName, doc.Name)
                If Not seen.Contains(key) Then
                    seen.Add(key)
                    docList.Add(doc)
                End If
            Next

            ' Optionally exclude the currently active document
            If ExceptCurrent Then
                Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing
                Try
                    activeDoc = app.ActiveDocument
                Catch
                    activeDoc = Nothing
                End Try
                If activeDoc IsNot Nothing Then
                    For i As System.Int32 = docList.Count - 1 To 0 Step -1
                        If System.Object.ReferenceEquals(docList(i), activeDoc) Then
                            docList.RemoveAt(i)
                        End If
                    Next
                End If
            End If

            If docList.Count = 0 Then
                Return "NONE"
            End If

            ' If silent mode requested: return all (after optional exclusion) without prompting
            If SilentAndGetAll Then
                Return BuildDocumentsResult(docList, IncludeName)
            End If

            ' Build selection items for each open document
            Dim selItems As New System.Collections.Generic.List(Of SelectionItem)()
            For i As System.Int32 = 0 To docList.Count - 1
                Dim d As Microsoft.Office.Interop.Word.Document = docList(i)
                selItems.Add(New SelectionItem($"{d.Name} ({d.FullName})", i + 1))
            Next

            ' Add “All open documents” and optional “None”
            Dim indexAll As System.Int32 = selItems.Count + 1
            selItems.Add(New SelectionItem("Add all open documents", indexAll))

            Dim indexNone As System.Int32 = -1
            If IncludeNone Then
                indexNone = indexAll + 1
                selItems.Add(New SelectionItem("Do not add any document", indexNone))
            End If

            ' Prompt user (default/highlight on "All")
            Dim itemsArray As SelectionItem() = selItems.ToArray()
            Dim picked As System.Int32 = SelectValue(itemsArray, indexAll, "Choose document to add …")

            ' User cancelled or invalid choice
            If picked < 1 Then
                Return System.String.Empty
            End If

            ' User explicitly chose "None"
            If IncludeNone AndAlso picked = indexNone Then
                Return System.String.Empty
            End If

            ' Determine targets based on selection
            Dim targets As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document)()
            If picked = indexAll Then
                targets.AddRange(docList)
            Else
                If picked - 1 >= 0 AndAlso picked - 1 < docList.Count Then
                    targets.Add(docList(picked - 1))
                Else
                    Return System.String.Empty
                End If
            End If

            Return BuildDocumentsResult(targets, IncludeName)

        Catch ex As System.Exception
            Return "ERROR " & ex.Message
        End Try
    End Function

    ' Helper to build the concatenated document content string
    Private Function BuildDocumentsResult(docs As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document),
                                          includeName As System.Boolean) As System.String
        Dim insertedDocuments As System.String = System.String.Empty
        Dim tagIndex As System.Int32 = 1

        For Each d As Microsoft.Office.Interop.Word.Document In docs
            If includeName Then insertedDocuments &= $"Here follows document no. {tagIndex} with the name '" & d.Name & "': " & vbCrLf
            insertedDocuments &= $"<DOCUMENT{tagIndex}>" & vbCrLf
            insertedDocuments &= d.Content.Text & vbCrLf
            insertedDocuments &= $"</DOCUMENT{tagIndex}>" & vbCrLf
            tagIndex += 1
        Next

        If System.String.IsNullOrEmpty(insertedDocuments) Then
            ShowCustomMessageBox("No content could be retrieved from the selected document(s).")
            Return System.String.Empty
        End If

        Return insertedDocuments
    End Function


    Public Function FindLongTextInChunks(ByVal findText As String, ByRef selection As Word.Selection, Optional Skipdeleted As Boolean = True) As Boolean

        Debug.WriteLine("Entering into FindLongTextAnchoredFast")

        Dim answer As Boolean = WordSearchHelper.FindLongTextAnchoredFast(selection, findText, Skipdeleted)

        Debug.WriteLine("Text found: " & "'" & selection.Text & "'")

        Return answer

    End Function

    Private Function oldGetSelectedTextLength() As Integer
        Try
            ' Get the active Word application
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Get the current selection in the active document
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

            ' Check if there is any selected text
            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Split the text on whitespace to count words,
            ' ignoring empty entries from multiple spaces/newlines
            Dim words = selectedText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                                       StringSplitOptions.RemoveEmptyEntries)
            Return words.Length

        Catch ex As System.Exception ' Explicitly referencing System.Exception
            ' Handle any exceptions and return 0 if an error occurs
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Counts real words in the current selection: sequences of letters (Unicode) optionally joined by internal apostrophes or hyphens; numeric/mixed tokens are ignored.
    ''' </summary>
    Private Function GetSelectedTextLength() As Integer
        Try
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Pattern:
            ' \b                Word boundary
            ' [\p{L}]+          One or more Unicode letters
            ' (?:['’\-‑–][\p{L}]+)*  Optional internal apostrophe/hyphen/dash + letters (e.g. don't, mother-in-law, rock’n’roll)
            ' \b                Word boundary
            ' Excludes tokens containing digits or starting with punctuation.
            Dim pattern As String = "\b[\p{L}]+(?:['’\-‑–][\p{L}]+)*\b"

            Return Regex.Matches(selectedText, pattern).Count
        Catch ex As System.Exception
            Return 0
        End Try
    End Function


    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        template = Regex.Replace(template, "{Codebasis}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack_2}", "", RegexOptions.IgnoreCase)

        Dim result As String = template

        Dim placeholderPattern As String = "\{([^}]+)\}"
        Dim matches As MatchCollection = Regex.Matches(template, placeholderPattern)

        For Each m As Match In matches
            Dim placeholder As String = m.Value          ' e.g. "{Name}"
            Dim varName As String = m.Groups(1).Value    ' e.g. "Name"

            ' Debug.WriteLine($"placeholder = {placeholder}  Varname = {varName}")
            ' Search for Field
            Dim fieldInfo = Me.GetType().GetField(varName)
            If fieldInfo IsNot Nothing Then
                Dim fieldValue = fieldInfo.GetValue(Me)
                If fieldValue IsNot Nothing Then
                    result = result.Replace(placeholder, fieldValue.ToString())
                End If
                Continue For
            End If

            ' Search for Property
            Dim propInfo = Me.GetType().GetProperty(varName)
            If propInfo IsNot Nothing Then
                Dim propValue = propInfo.GetValue(Me)
                If propValue IsNot Nothing Then
                    result = result.Replace(placeholder, propValue.ToString())
                End If
            End If
        Next

        Return result
    End Function


    ' Lightweight scope to show a progress window and enable cancellation for long-running operations.
    ' Works with the existing ProgressBarModule and DPIProgressForm/ProgressForm.
    Public NotInheritable Class ProgressScope
        Implements IDisposable

        Private ReadOnly _cts As System.Threading.CancellationTokenSource = New System.Threading.CancellationTokenSource()
        Private _uiThread As System.Threading.Thread
        Private _form As System.Windows.Forms.Form
        Private ReadOnly _useDpiForm As Boolean
        Private _closed As Integer = 0

        ' Start the scope and show a modeless progress window reading from ProgressBarModule.* every 250ms.
        Public Sub New(headerText As String,
                   initialLabel As String,
                   Optional max As Integer = 100,
                   Optional useDpiForm As Boolean = True)
            _useDpiForm = useDpiForm

            ' Initialize global progress state
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, max)
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = If(initialLabel, "")

            ' Spin up a dedicated STA UI thread with its own message loop for the progress form
            _uiThread = New System.Threading.Thread(
            Sub()
                Try
                    System.Windows.Forms.Application.EnableVisualStyles()
                    _form = If(_useDpiForm,
                               CType(New DPIProgressForm(headerText, initialLabel), System.Windows.Forms.Form),
                               CType(New ProgressForm(headerText, initialLabel), System.Windows.Forms.Form))

                    ' Run form (timer inside form pulls ProgressBarModule.* and closes itself when CancelOperation=True)
                    System.Windows.Forms.Application.Run(_form)
                Catch
                    ' Swallow — we always attempt to clean up in Dispose.
                End Try
            End Sub
        )
            _uiThread.IsBackground = True
            _uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
            _uiThread.Start()
        End Sub

        ' Report progress in a threadsafe way via your global ProgressBarModule.
        Public Shared Sub Report(current As Integer,
                             Optional max As Integer = -1,
                             Optional label As String = Nothing)
            If max >= 1 Then ProgressBarModule.GlobalProgressMax = max
            If label IsNot Nothing Then ProgressBarModule.GlobalProgressLabel = label
            ProgressBarModule.GlobalProgressValue = System.Math.Max(0, System.Math.Min(current, ProgressBarModule.GlobalProgressMax))
        End Sub

        ' Request cancellation (also triggered by the Cancel button in the UI)
        Public Sub RequestCancel()
            _cts.Cancel()
            ProgressBarModule.CancelOperation = True
        End Sub

        ' Check this frequently at safe points (between steps/chunks) to bail out early.
        Public ReadOnly Property CancelRequested As Boolean
            Get
                Return ProgressBarModule.CancelOperation OrElse _cts.IsCancellationRequested
            End Get
        End Property

        ' Bubble a CancellationToken if you prefer tokens.
        Public ReadOnly Property Token As System.Threading.CancellationToken
            Get
                Return _cts.Token
            End Get
        End Property

        ' In class ProgressScope
        Public Sub Dispose() Implements IDisposable.Dispose
            If System.Threading.Interlocked.Exchange(_closed, 1) <> 0 Then Return
            Try
                ' Signal cancel so the form’s timer path also closes itself
                ProgressBarModule.CancelOperation = True

                Dim f = _form
                If f IsNot Nothing Then
                    Try
                        If f.IsHandleCreated AndAlso Not f.IsDisposed Then
                            ' Request close on the form’s thread
                            f.BeginInvoke(New System.Action(Sub()
                                                                Try
                                                                    If Not f.IsDisposed Then f.Close()
                                                                Catch
                                                                End Try
                                                            End Sub))
                        End If
                    Catch
                        ' Ignore cross-thread or shutdown races
                    End Try
                End If
            Finally
                ' Give the UI thread a moment to exit cleanly
                Try
                    If _uiThread IsNot Nothing AndAlso _uiThread.IsAlive Then
                        If Not _uiThread.Join(1000) Then
                            Try : _uiThread.Interrupt() : Catch : End Try
                        End If
                    End If
                Catch
                End Try
            End Try
        End Sub
    End Class

    Public Function GetWordDefaultInterfaceLanguage() As String
        Try
            ' Get the language ID of the Word user interface
            Dim uiLanguageID As Integer = Globals.ThisAddIn.Application.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI)

            ' Convert the language ID to a human-readable name
            Dim cultureInfo As Globalization.CultureInfo = New Globalization.CultureInfo(uiLanguageID)

            ' Return the language display name
            Return cultureInfo.DisplayName
        Catch ex As System.Exception
            Return "English"
        End Try
    End Function

    Private Function CodeAPIKey(ByVal apiKey As String) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String
        Dim HadPrefix As Boolean = False

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to encode. Select the API Key you wish to encode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Encryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Encryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the API key has the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            HadPrefix = True
            ' Encrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Encrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Encrypt the modified key (without the prefix)
        resultKey = CodeString(modifiedKey, xcodebasis)

        ' Add the prefix back if it was present
        If HadPrefix Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function
    Private Function DeCodeAPIKey(ByVal apiKey As String) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to decode. Select the API Key you wish to decode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Decryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Decryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the key starts with the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            ' Decrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Decrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Decrypt the modified key (without the prefix)
        resultKey = DecodeString(modifiedKey, xcodebasis)

        ' Add the prefix back only if it was in the original key
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function


    Public Function VBAModuleWorking() As Boolean

        Dim xlApp As Microsoft.Office.Interop.Word.Application = Me.Application

        Try
            ' Call the VBA function
            Dim HelperVersion As Integer = CType(xlApp.Run("CheckAppHelper"), Integer)

            If HelperVersion >= MinHelperVersion Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
        End Try

    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Public Shared Function GetAsyncKeyState(ByVal vKey As System.Int32) As System.Int16
    End Function

    ' Convenience constant 
    Private Const VK_ESCAPE As System.Int32 = &H1B



End Class
